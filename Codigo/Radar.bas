Attribute VB_Name = "Radar"
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
' Modulo del radar de minimapa.
' - Perfil (NpcRadarActive=1, plan radar-v2 E3): lo que pidan los efectos de radar activos y los
'   objetos equipados (RadarVe/RadarRadio/RadarProtege/RadarPenetra). Ver RefreshRadarProfile.
' - Para GMs (NpcRadarActive=2): Hostile, NeutralNpc, GuardiaReal, GuardiaCaos, Mascota,
'   UserGm, UserCiudadano (incluye Armada/consejo) y UserCriminal (incluye Caos/concilio).
' - Siempre activos para todos: QuestNpc, NpcSymbol, PartyMember, ClanMember.
Option Explicit

' NPC_RADAR_TICK_MS (300 ms) vive ahora en modTickRules, junto a RadarDue.
Private Const MAX_RADAR_ENTITIES As Integer = 100

' Plan radar-v2 E1: NPCs agrupados por mapa. Se reconstruye como mucho cada RADAR_BUCKET_MS y cada
' jugador recorre solo los NPCs de su mapa (antes recorria NpcList entero 2 a 7 veces por tick).
' Orden estable por NpcIndex: mismo orden (y mismo hash) que el recorrido viejo.
Private Const RADAR_BUCKET_MS As Long = 100
' Plan radar-v2 E6: una marca de minimapa cada 3 s como mucho por jugador (el cliente espera 3000 ms; el server
' acepta desde 2800 para que el retraso de red no descarte una marca que el jugador ya vio en su minimapa).
Private Const MINIMAP_PING_COOLDOWN_MS As Long = 2800
Private RadarMapStart() As Long
Private RadarMapNpcs() As Integer
Private RadarBucketTick As Long
Private RadarBucketReady As Boolean

' Estados que viajan en el packet eNpcRadarQuestNpc (Byte por NPC):
'   0 = plano (sin glifo)
'   1 = quest disponible (! amarillo)
'   2 = quest en progreso (? gris)
'   3 = quest lista para entregar (? amarillo)
Public Const QUEST_STATE_NONE       As Byte = 0
Public Const QUEST_STATE_AVAILABLE  As Byte = 1
Public Const QUEST_STATE_INPROGRESS As Byte = 2
Public Const QUEST_STATE_READY      As Byte = 3

Public Const NPC_SYMBOL_NONE      As Byte = 0
Public Const NPC_SYMBOL_SACERDOTE As Byte = 1
Public Const NPC_SYMBOL_BANQUERO  As Byte = 2

' ============================================================================
' Clasificacion
' ============================================================================

' Devuelve e_RadarCategory o -1 si el NPC no debe mostrarse.
Private Function ClassifyNpc(ByVal NpcIndex As Integer) As Integer
    On Error GoTo ClassifyNpc_Err
    With NpcList(NpcIndex)
        If Not .flags.NPCActive Then
            ClassifyNpc = -1
            Exit Function
        End If
        ' NPCs no combatientes (banqueros, comerciantes, guardias estaticos) tienen MaxHp=0.
        ' Solo descartar como 'muerto' cuando tenian HP y la perdieron.
        If .Stats.MaxHp > 0 And .Stats.MinHp <= 0 Then
            ClassifyNpc = -1
            Exit Function
        End If
        ' Guardias antes que Mascota: una guardia nunca es mascota.
        Select Case .npcType
            Case e_NPCType.GuardiasCaos
                ClassifyNpc = e_RadarCategory.eRadarGuardiaCaos
                Exit Function
            Case e_NPCType.GuardiaReal, e_NPCType.GuardiaNpc, e_NPCType.Pretoriano, e_NPCType.ArenaGuard
                ClassifyNpc = e_RadarCategory.eRadarGuardiaReal
                Exit Function
            Case e_NPCType.Revividor, e_NPCType.Banquero
                ClassifyNpc = e_RadarCategory.eRadarNpcSymbol
                Exit Function
            Case e_NPCType.Quest
                ClassifyNpc = e_RadarCategory.eRadarQuestNpc
                Exit Function
        End Select
        ' Mascota invocada por un usuario (esqueletos, lobos, etc.)
        If .MaestroUser.ArrayIndex > 0 Then
            ClassifyNpc = e_RadarCategory.eRadarMascota
            Exit Function
        End If
        ' Plan radar-v2 E4: jefe marcado en npcs.dat (RadarJefe=1). Viaja aparte, con su icono.
        If .RadarJefe Then
            ClassifyNpc = e_RadarCategory.eRadarJefe
            Exit Function
        End If
        If .Hostile = 1 Then
            ClassifyNpc = e_RadarCategory.eRadarHostile
            Exit Function
        End If
        ClassifyNpc = e_RadarCategory.eRadarNeutralNpc
    End With
    Exit Function
ClassifyNpc_Err:
    ClassifyNpc = -1
    Call TraceError(Err.Number, Err.Description, "Radar.ClassifyNpc", Erl)
End Function

' Devuelve e_RadarCategory (UserGm/UserCiudadano/UserCriminal) o -1 si no es visible.
Private Function ClassifyUser(ByVal ObserverIndex As Integer, ByVal TargetIndex As Integer) As Integer
    On Error GoTo ClassifyUser_Err
    ClassifyUser = -1
    With UserList(TargetIndex)
        If Not .flags.UserLogged Then Exit Function
        Dim hidden As Boolean
        hidden = (.flags.AdminInvisible = 1) Or (.flags.invisible > 0) Or (.flags.Oculto = 1)
        If hidden Then
            If CompararPrivilegiosUser(ObserverIndex, TargetIndex) < 0 Then Exit Function
        End If
        If .flags.Privilegios <> e_PlayerType.User Then
            ClassifyUser = e_RadarCategory.eRadarUserGm
        ElseIf .Faccion.Status = e_Facciones.Criminal Or .Faccion.Status = e_Facciones.Caos Or .Faccion.Status = e_Facciones.concilio Then
            ' Plan radar-v2 D2: Caos/concilio se ven como criminal; Armada/consejo caen en ciudadano.
            ClassifyUser = e_RadarCategory.eRadarUserCriminal
        Else
            ClassifyUser = e_RadarCategory.eRadarUserCiudadano
        End If
    End With
    Exit Function
ClassifyUser_Err:
    ClassifyUser = -1
    Call TraceError(Err.Number, Err.Description, "Radar.ClassifyUser", Erl)
End Function

' Devuelve e_RadarCategory granular (4 facciones reales) o -1 si no es visible.
' Usado por el modo NpcRadarActive=3 (hechizo de deteccion de usuarios).
Private Function ClassifyUserDetailed(ByVal ObserverIndex As Integer, ByVal TargetIndex As Integer) As Integer
    On Error GoTo ClassifyUserDetailed_Err
    ClassifyUserDetailed = -1
    With UserList(TargetIndex)
        If Not .flags.UserLogged Then Exit Function
        If .flags.Privilegios <> e_PlayerType.User Then Exit Function
        Dim hidden As Boolean
        hidden = (.flags.AdminInvisible = 1) Or (.flags.invisible > 0) Or (.flags.Oculto = 1)
        ' Si el target esta invisible/oculto, no mostrarlo en el radar.
        ' Excepcion: solo lo ve un observador con privilegios estrictamente mayores.
        If hidden Then
            If CompararPrivilegiosUser(ObserverIndex, TargetIndex) <= 0 Then Exit Function
        End If
        ' Plan radar-v2 E3: proteccion contra la deteccion (objetos/efectos). Solo el radar por perfil
        ' usa esta funcion: el GM usa ClassifyUser y party/clan tienen sus propias reglas.
        If .flags.RadarProtection > UserList(ObserverIndex).flags.RadarPenetration Then Exit Function
        Select Case .Faccion.Status
            Case e_Facciones.Criminal
                ClassifyUserDetailed = e_RadarCategory.eRadarUserCriminal
            Case e_Facciones.Ciudadano
                ClassifyUserDetailed = e_RadarCategory.eRadarUserCiudadano
            Case e_Facciones.Armada, e_Facciones.consejo
                ClassifyUserDetailed = e_RadarCategory.eRadarUserImperial
            Case e_Facciones.Caos, e_Facciones.concilio
                ClassifyUserDetailed = e_RadarCategory.eRadarUserCaos
            Case Else
                ClassifyUserDetailed = e_RadarCategory.eRadarUserCiudadano
        End Select
    End With
    Exit Function
ClassifyUserDetailed_Err:
    ClassifyUserDetailed = -1
    Call TraceError(Err.Number, Err.Description, "Radar.ClassifyUserDetailed", Erl)
End Function

Private Function PacketIdForCategory(ByVal Category As e_RadarCategory) As ServerPacketID
    Select Case Category
        Case e_RadarCategory.eRadarHostile:        PacketIdForCategory = ServerPacketID.eNpcRadar
        Case e_RadarCategory.eRadarNeutralNpc:     PacketIdForCategory = ServerPacketID.eNpcRadarNeutralNpc
        Case e_RadarCategory.eRadarQuestNpc:       PacketIdForCategory = ServerPacketID.eNpcRadarQuestNpc
        Case e_RadarCategory.eRadarGuardiaReal:    PacketIdForCategory = ServerPacketID.eNpcRadarGuardiaReal
        Case e_RadarCategory.eRadarGuardiaCaos:    PacketIdForCategory = ServerPacketID.eNpcRadarGuardiaCaos
        Case e_RadarCategory.eRadarMascota:        PacketIdForCategory = ServerPacketID.eNpcRadarMascota
        Case e_RadarCategory.eRadarUserGm:         PacketIdForCategory = ServerPacketID.eNpcRadarUserGm
        Case e_RadarCategory.eRadarUserCiudadano:  PacketIdForCategory = ServerPacketID.eNpcRadarUserCiudadano
        Case e_RadarCategory.eRadarUserCriminal:   PacketIdForCategory = ServerPacketID.eNpcRadarUserCriminal
        Case e_RadarCategory.eRadarNpcSymbol:      PacketIdForCategory = ServerPacketID.eNpcRadarNpcSymbol
        Case e_RadarCategory.eRadarUserImperial:   PacketIdForCategory = ServerPacketID.eNpcRadarUserImperial
        Case e_RadarCategory.eRadarUserCaos:       PacketIdForCategory = ServerPacketID.eNpcRadarUserCaos
        Case e_RadarCategory.eRadarPartyMember:    PacketIdForCategory = ServerPacketID.eNpcRadarPartyMember
        Case e_RadarCategory.eRadarClanMember:     PacketIdForCategory = ServerPacketID.eNpcRadarClanMember
        Case e_RadarCategory.eRadarJefe:           PacketIdForCategory = ServerPacketID.eNpcRadarJefe
    End Select
End Function

' ============================================================================
' Recoleccion
' ============================================================================

Private Sub EnsureNpcBuckets(ByVal nowTick As Long)
    On Error GoTo EnsureNpcBuckets_Err
    If RadarBucketReady Then
        If CooldownActive(nowTick, RadarBucketTick, RADAR_BUCKET_MS) Then Exit Sub
    End If
    RadarBucketReady = False
    If NumMaps <= 0 Then Exit Sub
    Dim m As Long
    Dim n As Long
    ReDim RadarMapStart(1 To NumMaps + 1) As Long
    For n = 1 To LastNPC
        m = NpcList(n).pos.Map
        If m >= 1 And m <= NumMaps Then
            If NpcList(n).flags.NPCActive Then RadarMapStart(m) = RadarMapStart(m) + 1
        End If
    Next n
    Dim mapCount As Long
    Dim nextStart As Long
    nextStart = 1
    For m = 1 To NumMaps
        mapCount = RadarMapStart(m)
        RadarMapStart(m) = nextStart
        nextStart = nextStart + mapCount
    Next m
    RadarMapStart(NumMaps + 1) = nextStart
    If nextStart > 1 Then
        ReDim RadarMapNpcs(1 To nextStart - 1) As Integer
        Dim fillPos() As Long
        ReDim fillPos(1 To NumMaps) As Long
        For m = 1 To NumMaps
            fillPos(m) = RadarMapStart(m)
        Next m
        For n = 1 To LastNPC
            m = NpcList(n).pos.Map
            If m >= 1 And m <= NumMaps Then
                If NpcList(n).flags.NPCActive Then
                    RadarMapNpcs(fillPos(m)) = n
                    fillPos(m) = fillPos(m) + 1
                End If
            End If
        Next n
    End If
    RadarBucketTick = NonZeroTickFrom(nowTick)
    RadarBucketReady = True
    Exit Sub
EnsureNpcBuckets_Err:
    RadarBucketReady = False
    Call TraceError(Err.Number, Err.Description, "Radar.EnsureNpcBuckets", Erl)
End Sub

' Rango [firstK, lastK] de RadarMapNpcs para el mapa. Vacio (1, 0) si no hay baldes.
Private Sub BucketRange(ByVal userMap As Integer, ByRef firstK As Long, ByRef lastK As Long)
    firstK = 1
    lastK = 0
    Call EnsureNpcBuckets(GetTickCountRaw())
    If Not RadarBucketReady Then Exit Sub
    If userMap < 1 Or userMap > UBound(RadarMapStart) - 1 Then Exit Sub
    firstK = RadarMapStart(userMap)
    lastK = RadarMapStart(userMap + 1) - 1
End Sub

' Plan radar-v2 D4: un invisible u oculto solo aparece para un observador de rango estrictamente
' mayor (regla de ClassifyUserDetailed). Vale para party y clan.
Private Function HiddenFromObserver(ByVal ObserverIndex As Integer, ByVal TargetIndex As Integer) As Boolean
    With UserList(TargetIndex)
        If (.flags.AdminInvisible = 1) Or (.flags.invisible > 0) Or (.flags.Oculto = 1) Then
            HiddenFromObserver = (CompararPrivilegiosUser(ObserverIndex, TargetIndex) <= 0)
        End If
    End With
End Function

' Recolecta NPCs del mapa que matchean Category. Devuelve un hash.
Private Function CollectRadarNpcs(ByVal UserIndex As Integer, ByVal Category As e_RadarCategory, ByRef Xs() As Byte, ByRef Ys() As Byte, ByRef count As Integer) As Long
    On Error GoTo CollectRadarNpcs_Err
    Dim userMap As Integer
    userMap = UserList(UserIndex).pos.Map
    count = 0
    Dim h As Long
    h = (CLng(userMap) Mod 32768) * 17 + CLng(Category) * 31
    If userMap <= 0 Then
        CollectRadarNpcs = h
        Exit Function
    End If
    Dim NpcIndex As Integer
    Dim px As Byte, py As Byte
    Dim cat As Integer
    Dim firstK As Long, lastK As Long, k As Long
    Call BucketRange(userMap, firstK, lastK)
    For k = firstK To lastK
        NpcIndex = RadarMapNpcs(k)
        With NpcList(NpcIndex)
            If .pos.Map = userMap Then
                cat = ClassifyNpc(NpcIndex)
                If cat = Category And InRadarRange(UserIndex, .pos.x, .pos.y) Then
                    count = count + 1
                    px = CByte(.pos.x)
                    py = CByte(.pos.y)
                    Xs(count) = px
                    Ys(count) = py
                    h = ((h Mod 1048576) * 17) Xor (CLng(px) * 256& + CLng(py))
                    If count >= MAX_RADAR_ENTITIES Then Exit For
                End If
            End If
        End With
    Next k
    CollectRadarNpcs = h
    Exit Function
CollectRadarNpcs_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.CollectRadarNpcs", Erl)
End Function

' Recolecta usuarios del mapa que matchean Category (UserGm/UserCiudadano/UserCriminal).
' Respeta privacidad de invisibles (regla CompararPrivilegiosUser).
Private Function CollectRadarUsersByCategory(ByVal ObserverIndex As Integer, ByVal Category As e_RadarCategory, ByRef Xs() As Byte, ByRef Ys() As Byte, ByRef count As Integer) As Long
    On Error GoTo CollectRadarUsersByCategory_Err
    Dim userMap As Integer
    userMap = UserList(ObserverIndex).pos.Map
    count = 0
    Dim h As Long
    h = (CLng(userMap) Mod 32768) * 17 + CLng(Category) * 31
    If userMap <= 0 Then
        CollectRadarUsersByCategory = h
        Exit Function
    End If
    Dim i As Integer
    Dim px As Byte, py As Byte
    Dim cat As Integer
    For i = 1 To LastUser
        If i <> ObserverIndex Then
            If UserList(i).pos.Map = userMap Then
                If UserList(ObserverIndex).flags.NpcRadarActive = 2 Then
                    ' Radar GM: 3 categorias de usuario. Antes ciudadano/criminal pasaban por
                    ' ClassifyUserDetailed y Armada/Caos no aparecian (plan radar-v2 D2).
                    cat = ClassifyUser(ObserverIndex, i)
                Else
                    cat = ClassifyUserDetailed(ObserverIndex, i)
                End If
                If cat = Category And InRadarRange(ObserverIndex, UserList(i).pos.x, UserList(i).pos.y) Then
                    count = count + 1
                    px = CByte(UserList(i).pos.x)
                    py = CByte(UserList(i).pos.y)
                    Xs(count) = px
                    Ys(count) = py
                    h = ((h Mod 1048576) * 17) Xor (CLng(px) * 256& + CLng(py))
                    If count >= MAX_RADAR_ENTITIES Then Exit For
                End If
            End If
        End If
    Next i
    CollectRadarUsersByCategory = h
    Exit Function
CollectRadarUsersByCategory_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.CollectRadarUsersByCategory", Erl)
End Function

' Recolecta miembros de la party del observador que estan en el mismo mapa.
' Respeta invisibilidad con CompararPrivilegiosUser, igual que CollectRadarUsersByCategory.
Private Function CollectRadarPartyMembers(ByVal ObserverIndex As Integer, ByRef Xs() As Byte, ByRef Ys() As Byte, ByRef count As Integer) As Long
    On Error GoTo CollectRadarPartyMembers_Err
    Dim userMap As Integer
    userMap = UserList(ObserverIndex).pos.Map
    count = 0
    Dim h As Long
    h = (CLng(userMap) Mod 32768) * 17 + CLng(e_RadarCategory.eRadarPartyMember) * 31
    If userMap <= 0 Then CollectRadarPartyMembers = h: Exit Function
    If Not UserList(ObserverIndex).Grupo.EnGrupo Then CollectRadarPartyMembers = h: Exit Function
    Dim leaderIndex As Integer
    leaderIndex = UserList(ObserverIndex).Grupo.Lider.ArrayIndex
    If leaderIndex <= 0 Then CollectRadarPartyMembers = h: Exit Function
    Dim cantidad As Byte
    cantidad = UserList(leaderIndex).Grupo.CantidadMiembros
    Dim i As Integer
    Dim memberIndex As Integer
    Dim px As Byte, py As Byte
    For i = 1 To cantidad
        memberIndex = UserList(leaderIndex).Grupo.Miembros(i).ArrayIndex
        ' Plan 27.001 (05.002 wave 7): the party radar showed the exact position of a rival who
        ' shares a party inside an event or duel. GrupoCruzaBandos is False outside events.
        ' IsInObserverParty is NOT touched on purpose: it keeps returning True for the member hidden
        ' here, which is what stops CollectRadarClanMembers from showing him through the clan radar.
        If memberIndex > 0 And memberIndex <> ObserverIndex And Not UserMod.GrupoCruzaBandos(ObserverIndex, memberIndex) Then
            If UserList(memberIndex).flags.UserLogged Then
                If UserList(memberIndex).pos.Map = userMap Then
                    If Not HiddenFromObserver(ObserverIndex, memberIndex) Then
                        count = count + 1
                        px = CByte(UserList(memberIndex).pos.x)
                        py = CByte(UserList(memberIndex).pos.y)
                        Xs(count) = px
                        Ys(count) = py
                        h = ((h Mod 1048576) * 17) Xor (CLng(px) * 256& + CLng(py))
                        If count >= MAX_RADAR_ENTITIES Then Exit For
                    End If
                End If
            End If
        End If
    Next i
    CollectRadarPartyMembers = h
    Exit Function
CollectRadarPartyMembers_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.CollectRadarPartyMembers", Erl)
End Function

' Devuelve True si memberIndex es miembro de la party del observador.
Private Function IsInObserverParty(ByVal ObserverIndex As Integer, ByVal memberIndex As Integer) As Boolean
    If Not UserList(ObserverIndex).Grupo.EnGrupo Then Exit Function
    Dim leaderIndex As Integer
    leaderIndex = UserList(ObserverIndex).Grupo.Lider.ArrayIndex
    If leaderIndex <= 0 Then Exit Function
    Dim cantidad As Byte
    cantidad = UserList(leaderIndex).Grupo.CantidadMiembros
    Dim i As Integer
    For i = 1 To cantidad
        If UserList(leaderIndex).Grupo.Miembros(i).ArrayIndex = memberIndex Then
            IsInObserverParty = True
            Exit Function
        End If
    Next i
End Function

' Recolecta miembros del clan del observador que estan en el mismo mapa.
' Excluye miembros de party (party tiene prioridad visual sobre clan).
Private Function CollectRadarClanMembers(ByVal ObserverIndex As Integer, ByRef Xs() As Byte, ByRef Ys() As Byte, ByRef count As Integer) As Long
    On Error GoTo CollectRadarClanMembers_Err
    Dim userMap As Integer
    userMap = UserList(ObserverIndex).pos.Map
    count = 0
    Dim h As Long
    h = (CLng(userMap) Mod 32768) * 17 + CLng(e_RadarCategory.eRadarClanMember) * 31
    If userMap <= 0 Then CollectRadarClanMembers = h: Exit Function
    Dim observerGuild As Integer
    observerGuild = UserList(ObserverIndex).GuildIndex
    If observerGuild <= 0 Then CollectRadarClanMembers = h: Exit Function
    ' Plan radar-v2 D5: dentro de un evento o reto no hay radar de clan (dos del mismo clan pueden
    ' estar en bandos opuestos). La party sigue con su regla de bandos (GrupoCruzaBandos).
    If CustomScenarios.IsUserInMatch(ObserverIndex) Then CollectRadarClanMembers = h: Exit Function
    Dim i As Integer
    Dim px As Byte, py As Byte
    For i = 1 To LastUser
        If i <> ObserverIndex Then
            If UserList(i).flags.UserLogged Then
                If UserList(i).GuildIndex = observerGuild Then
                    If UserList(i).pos.Map = userMap Then
                        If Not IsInObserverParty(ObserverIndex, i) And Not CustomScenarios.IsUserInMatch(i) Then
                            If Not HiddenFromObserver(ObserverIndex, i) Then
                                count = count + 1
                                px = CByte(UserList(i).pos.x)
                                py = CByte(UserList(i).pos.y)
                                Xs(count) = px
                                Ys(count) = py
                                h = ((h Mod 1048576) * 17) Xor (CLng(px) * 256& + CLng(py))
                                If count >= MAX_RADAR_ENTITIES Then Exit For
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
    CollectRadarClanMembers = h
    Exit Function
CollectRadarClanMembers_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.CollectRadarClanMembers", Erl)
End Function

' ============================================================================
' Estado de quest por usuario (siempre activo para todos los jugadores)
' ============================================================================

' Disponible: no la hizo, cumple nivel, quest previa y clase. Lee QuestList sin copiar la quest
' (plan radar-v2 E1: antes se copiaba el UDT entero por cada NPC de quest, jugador y tick).
Private Function IsQuestAvailableForUser(ByVal UserIndex As Integer, ByVal qIdx As Integer) As Boolean
    If qIdx < LBound(QuestList) Or qIdx > UBound(QuestList) Then Exit Function
    If UserDoneQuest(UserIndex, qIdx) Then Exit Function
    With QuestList(qIdx)
        If UserList(UserIndex).Stats.ELV < .RequiredLevel Then Exit Function
        If .RequiredQuest <> 0 Then
            If Not UserDoneQuest(UserIndex, .RequiredQuest) Then Exit Function
        End If
        If .RequiredClassesCount > 0 Then
            Dim ci As Integer
            For ci = 1 To .RequiredClassesCount
                If UserList(UserIndex).clase = .RequiredClass(ci) Then
                    IsQuestAvailableForUser = True
                    Exit Function
                End If
            Next ci
            Exit Function
        End If
    End With
    IsQuestAvailableForUser = True
End Function

' Calcula el estado de quest agregado para un NPC respecto a un usuario.
' Prioridad: Ready > InProgress > Available > None.
' Reusa TieneQuest / FinishQuestCheck / UserDoneQuest de ModQuest.bas.
Private Function GetQuestStateForUser(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Byte
    On Error GoTo GetQuestStateForUser_Err
    Dim hasReady As Boolean
    Dim hasInProgress As Boolean
    Dim hasAvailable As Boolean
    With NpcList(NpcIndex)
        If .NumQuest <= 0 Then
            GetQuestStateForUser = QUEST_STATE_NONE
            Exit Function
        End If
        Dim q As Integer
        Dim qIdx As Integer
        Dim slot As Byte
        For q = 1 To .NumQuest
            qIdx = .QuestNumber(q)
            If qIdx > 0 Then
                slot = TieneQuest(UserIndex, qIdx)
                If slot > 0 Then
                    If FinishQuestCheck(UserIndex, qIdx, slot) Then
                        hasReady = True
                    Else
                        hasInProgress = True
                    End If
                Else
                    If IsQuestAvailableForUser(UserIndex, qIdx) Then hasAvailable = True
                End If
            End If
        Next q
    End With
    If hasReady Then
        GetQuestStateForUser = QUEST_STATE_READY
    ElseIf hasInProgress Then
        GetQuestStateForUser = QUEST_STATE_INPROGRESS
    ElseIf hasAvailable Then
        GetQuestStateForUser = QUEST_STATE_AVAILABLE
    Else
        GetQuestStateForUser = QUEST_STATE_NONE
    End If
    Exit Function
GetQuestStateForUser_Err:
    GetQuestStateForUser = QUEST_STATE_NONE
    Call TraceError(Err.Number, Err.Description, "Radar.GetQuestStateForUser", Erl)
End Function

' Recolecta NPCs eRadarQuestNpc del mapa, con estado de quest por usuario.
' Hash incorpora state -> cualquier cambio dispara reenvio.
Private Function CollectQuestNpcRadar(ByVal UserIndex As Integer, ByRef Xs() As Byte, ByRef Ys() As Byte, ByRef Ss() As Byte, ByRef count As Integer) As Long
    On Error GoTo CollectQuestNpcRadar_Err
    Dim userMap As Integer
    userMap = UserList(UserIndex).pos.Map
    count = 0
    Dim h As Long
    h = (CLng(userMap) Mod 32768) * 17 + CLng(e_RadarCategory.eRadarQuestNpc) * 31
    If userMap <= 0 Then
        CollectQuestNpcRadar = h
        Exit Function
    End If
    Dim NpcIndex As Integer
    Dim px As Byte, py As Byte
    Dim cat As Integer
    Dim state As Byte
    Dim firstK As Long, lastK As Long, k As Long
    Call BucketRange(userMap, firstK, lastK)
    For k = firstK To lastK
        NpcIndex = RadarMapNpcs(k)
        With NpcList(NpcIndex)
            If .pos.Map = userMap Then
                cat = ClassifyNpc(NpcIndex)
                If cat = e_RadarCategory.eRadarQuestNpc Then
                    state = GetQuestStateForUser(UserIndex, NpcIndex)
                    count = count + 1
                    px = CByte(.pos.x)
                    py = CByte(.pos.y)
                    Xs(count) = px
                    Ys(count) = py
                    Ss(count) = state
                    h = ((h Mod 1048576) * 17) Xor (CLng(px) * 256& + CLng(py))
                    h = ((h Mod 1048576) * 31) Xor CLng(state)
                    If count >= MAX_RADAR_ENTITIES Then Exit For
                End If
            End If
        End With
    Next k
    CollectQuestNpcRadar = h
    Exit Function
CollectQuestNpcRadar_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.CollectQuestNpcRadar", Erl)
End Function

Private Function GetNpcSymbolSubtype(ByVal NpcIndex As Integer) As Byte
    Select Case NpcList(NpcIndex).npcType
        Case e_NPCType.Revividor: GetNpcSymbolSubtype = NPC_SYMBOL_SACERDOTE
        Case e_NPCType.Banquero:  GetNpcSymbolSubtype = NPC_SYMBOL_BANQUERO
        Case Else:                GetNpcSymbolSubtype = NPC_SYMBOL_NONE
    End Select
End Function

Private Function CollectNpcSymbolRadar(ByVal UserIndex As Integer, ByRef Xs() As Byte, ByRef Ys() As Byte, ByRef Ss() As Byte, ByRef count As Integer) As Long
    On Error GoTo CollectNpcSymbolRadar_Err
    Dim userMap As Integer
    userMap = UserList(UserIndex).pos.Map
    count = 0
    Dim h As Long
    h = (CLng(userMap) Mod 32768) * 17 + CLng(e_RadarCategory.eRadarNpcSymbol) * 31
    If userMap <= 0 Then CollectNpcSymbolRadar = h: Exit Function
    Dim NpcIndex As Integer
    Dim px As Byte, py As Byte
    Dim cat As Integer
    Dim sub_ As Byte
    Dim firstK As Long, lastK As Long, k As Long
    Call BucketRange(userMap, firstK, lastK)
    For k = firstK To lastK
        NpcIndex = RadarMapNpcs(k)
        With NpcList(NpcIndex)
            If .pos.Map = userMap Then
                cat = ClassifyNpc(NpcIndex)
                If cat = e_RadarCategory.eRadarNpcSymbol Then
                    sub_ = GetNpcSymbolSubtype(NpcIndex)
                    count = count + 1
                    px = CByte(.pos.x)
                    py = CByte(.pos.y)
                    Xs(count) = px
                    Ys(count) = py
                    Ss(count) = sub_
                    h = ((h Mod 1048576) * 17) Xor (CLng(px) * 256& + CLng(py))
                    h = ((h Mod 1048576) * 31) Xor CLng(sub_)
                    If count >= MAX_RADAR_ENTITIES Then Exit For
                End If
            End If
        End With
    Next k
    CollectNpcSymbolRadar = h
    Exit Function
CollectNpcSymbolRadar_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.CollectNpcSymbolRadar", Erl)
End Function

' Plan radar-v2 E4: jefes del mapa con su icono (0 = sin icono: el cliente dibuja un punto grande).
Private Function CollectJefeRadar(ByVal UserIndex As Integer, ByRef Xs() As Byte, ByRef Ys() As Byte, ByRef Ss() As Byte, ByRef count As Integer) As Long
    On Error GoTo CollectJefeRadar_Err
    Dim userMap As Integer
    userMap = UserList(UserIndex).pos.Map
    count = 0
    Dim h As Long
    h = (CLng(userMap) Mod 32768) * 17 + CLng(e_RadarCategory.eRadarJefe) * 31
    If userMap <= 0 Then CollectJefeRadar = h: Exit Function
    Dim NpcIndex As Integer
    Dim px As Byte, py As Byte
    Dim firstK As Long, lastK As Long, k As Long
    Call BucketRange(userMap, firstK, lastK)
    For k = firstK To lastK
        NpcIndex = RadarMapNpcs(k)
        With NpcList(NpcIndex)
            If .pos.Map = userMap Then
                If ClassifyNpc(NpcIndex) = e_RadarCategory.eRadarJefe And InRadarRange(UserIndex, .pos.x, .pos.y) Then
                    count = count + 1
                    px = CByte(.pos.x)
                    py = CByte(.pos.y)
                    Xs(count) = px
                    Ys(count) = py
                    Ss(count) = .RadarIcono
                    h = ((h Mod 1048576) * 17) Xor (CLng(px) * 256& + CLng(py))
                    h = ((h Mod 1048576) * 31) Xor CLng(.RadarIcono)
                    If count >= MAX_RADAR_ENTITIES Then Exit For
                End If
            End If
        End With
    Next k
    CollectJefeRadar = h
    Exit Function
CollectJefeRadar_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.CollectJefeRadar", Erl)
End Function

' ============================================================================
' Hash storage (por categoria)
' ============================================================================

Private Function GetCategoryHash(ByVal UserIndex As Integer, ByVal Category As e_RadarCategory) As Long
    Select Case Category
        Case e_RadarCategory.eRadarHostile:        GetCategoryHash = UserList(UserIndex).Counters.NpcRadarHashHostile
        Case e_RadarCategory.eRadarNeutralNpc:     GetCategoryHash = UserList(UserIndex).Counters.NpcRadarHashNeutralNpc
        Case e_RadarCategory.eRadarQuestNpc:       GetCategoryHash = UserList(UserIndex).Counters.NpcRadarHashQuestNpc
        Case e_RadarCategory.eRadarGuardiaReal:    GetCategoryHash = UserList(UserIndex).Counters.NpcRadarHashGuardiaReal
        Case e_RadarCategory.eRadarGuardiaCaos:    GetCategoryHash = UserList(UserIndex).Counters.NpcRadarHashGuardiaCaos
        Case e_RadarCategory.eRadarMascota:        GetCategoryHash = UserList(UserIndex).Counters.NpcRadarHashMascota
        Case e_RadarCategory.eRadarUserGm:         GetCategoryHash = UserList(UserIndex).Counters.NpcRadarHashUserGm
        Case e_RadarCategory.eRadarUserCiudadano:  GetCategoryHash = UserList(UserIndex).Counters.NpcRadarHashUserCiudadano
        Case e_RadarCategory.eRadarUserCriminal:   GetCategoryHash = UserList(UserIndex).Counters.NpcRadarHashUserCriminal
        Case e_RadarCategory.eRadarNpcSymbol:      GetCategoryHash = UserList(UserIndex).Counters.NpcRadarHashNpcSymbol
        Case e_RadarCategory.eRadarUserImperial:   GetCategoryHash = UserList(UserIndex).Counters.NpcRadarHashUserImperial
        Case e_RadarCategory.eRadarUserCaos:       GetCategoryHash = UserList(UserIndex).Counters.NpcRadarHashUserCaos
        Case e_RadarCategory.eRadarPartyMember:    GetCategoryHash = UserList(UserIndex).Counters.NpcRadarHashPartyMember
        Case e_RadarCategory.eRadarClanMember:     GetCategoryHash = UserList(UserIndex).Counters.NpcRadarHashClanMember
        Case e_RadarCategory.eRadarJefe:           GetCategoryHash = UserList(UserIndex).Counters.NpcRadarHashJefe
    End Select
End Function

Private Sub SetCategoryHash(ByVal UserIndex As Integer, ByVal Category As e_RadarCategory, ByVal h As Long)
    Select Case Category
        Case e_RadarCategory.eRadarHostile:        UserList(UserIndex).Counters.NpcRadarHashHostile = h
        Case e_RadarCategory.eRadarNeutralNpc:     UserList(UserIndex).Counters.NpcRadarHashNeutralNpc = h
        Case e_RadarCategory.eRadarQuestNpc:       UserList(UserIndex).Counters.NpcRadarHashQuestNpc = h
        Case e_RadarCategory.eRadarGuardiaReal:    UserList(UserIndex).Counters.NpcRadarHashGuardiaReal = h
        Case e_RadarCategory.eRadarGuardiaCaos:    UserList(UserIndex).Counters.NpcRadarHashGuardiaCaos = h
        Case e_RadarCategory.eRadarMascota:        UserList(UserIndex).Counters.NpcRadarHashMascota = h
        Case e_RadarCategory.eRadarUserGm:         UserList(UserIndex).Counters.NpcRadarHashUserGm = h
        Case e_RadarCategory.eRadarUserCiudadano:  UserList(UserIndex).Counters.NpcRadarHashUserCiudadano = h
        Case e_RadarCategory.eRadarUserCriminal:   UserList(UserIndex).Counters.NpcRadarHashUserCriminal = h
        Case e_RadarCategory.eRadarNpcSymbol:      UserList(UserIndex).Counters.NpcRadarHashNpcSymbol = h
        Case e_RadarCategory.eRadarUserImperial:   UserList(UserIndex).Counters.NpcRadarHashUserImperial = h
        Case e_RadarCategory.eRadarUserCaos:       UserList(UserIndex).Counters.NpcRadarHashUserCaos = h
        Case e_RadarCategory.eRadarPartyMember:    UserList(UserIndex).Counters.NpcRadarHashPartyMember = h
        Case e_RadarCategory.eRadarClanMember:     UserList(UserIndex).Counters.NpcRadarHashClanMember = h
        Case e_RadarCategory.eRadarJefe:           UserList(UserIndex).Counters.NpcRadarHashJefe = h
    End Select
End Sub

Private Sub ResetAllHashes(ByVal UserIndex As Integer)
    Call ResetCategoryHashes(UserIndex)
    UserList(UserIndex).Counters.NpcRadarLastMap = -1
End Sub

' Pone en 0 el hash de todas las categorias: el proximo tick de cada una reenvia su lista.
Private Sub ResetCategoryHashes(ByVal UserIndex As Integer)
    With UserList(UserIndex).Counters
        .NpcRadarHashNpcSymbol = 0
        .NpcRadarHashHostile = 0
        .NpcRadarHashNeutralNpc = 0
        .NpcRadarHashQuestNpc = 0
        .NpcRadarHashGuardiaReal = 0
        .NpcRadarHashGuardiaCaos = 0
        .NpcRadarHashMascota = 0
        .NpcRadarHashUserGm = 0
        .NpcRadarHashUserCiudadano = 0
        .NpcRadarHashUserCriminal = 0
        .NpcRadarHashUserImperial = 0
        .NpcRadarHashUserCaos = 0
        .NpcRadarHashPartyMember = 0
        .NpcRadarHashClanMember = 0
        .NpcRadarHashJefe = 0
    End With
End Sub

' ============================================================================
' Envio condicional por categoria (con dedup)
' ============================================================================

Private Sub SendCategoryUpdate(ByVal UserIndex As Integer, ByVal Category As e_RadarCategory, ByVal forceSend As Boolean)
    On Error GoTo SendCategoryUpdate_Err
    Dim Xs(1 To MAX_RADAR_ENTITIES) As Byte
    Dim Ys(1 To MAX_RADAR_ENTITIES) As Byte
    Dim count As Integer
    Dim newHash As Long
    If Category = e_RadarCategory.eRadarQuestNpc Then
        ' QuestNpc: formato extendido con estado por NPC.
        Dim Ss(1 To MAX_RADAR_ENTITIES) As Byte
        newHash = CollectQuestNpcRadar(UserIndex, Xs, Ys, Ss, count)
        If newHash = 0 Then newHash = 1   ' 0 = "sin enviar" (ResetCategoryHashes)
        If (Not forceSend) And newHash = GetCategoryHash(UserIndex, Category) Then
            Exit Sub
        End If
        Call SetCategoryHash(UserIndex, Category, newHash)
        Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageRadarQuestNpc(PacketIdForCategory(Category), CByte(count), Xs, Ys, Ss))
        Exit Sub
    End If
    If Category = e_RadarCategory.eRadarNpcSymbol Then
        Dim Ss2(1 To MAX_RADAR_ENTITIES) As Byte
        newHash = CollectNpcSymbolRadar(UserIndex, Xs, Ys, Ss2, count)
        If newHash = 0 Then newHash = 1   ' 0 = "sin enviar" (ResetCategoryHashes)
        If (Not forceSend) And newHash = GetCategoryHash(UserIndex, Category) Then
            Exit Sub
        End If
        Call SetCategoryHash(UserIndex, Category, newHash)
        Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageRadarQuestNpc(PacketIdForCategory(Category), CByte(count), Xs, Ys, Ss2))
        Exit Sub
    End If
    If Category = e_RadarCategory.eRadarJefe Then
        Dim Ss3(1 To MAX_RADAR_ENTITIES) As Byte
        newHash = CollectJefeRadar(UserIndex, Xs, Ys, Ss3, count)
        If newHash = 0 Then newHash = 1   ' 0 = "sin enviar" (ResetCategoryHashes)
        If (Not forceSend) And newHash = GetCategoryHash(UserIndex, Category) Then
            Exit Sub
        End If
        Call SetCategoryHash(UserIndex, Category, newHash)
        Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageRadarQuestNpc(PacketIdForCategory(Category), CByte(count), Xs, Ys, Ss3))
        Exit Sub
    End If
    If Category = e_RadarCategory.eRadarPartyMember Then
        newHash = CollectRadarPartyMembers(UserIndex, Xs, Ys, count)
    ElseIf Category = e_RadarCategory.eRadarClanMember Then
        newHash = CollectRadarClanMembers(UserIndex, Xs, Ys, count)
    ElseIf Category >= e_RadarCategory.eRadarUserGm Then
        newHash = CollectRadarUsersByCategory(UserIndex, Category, Xs, Ys, count)
    Else
        newHash = CollectRadarNpcs(UserIndex, Category, Xs, Ys, count)
    End If
    If newHash = 0 Then newHash = 1   ' 0 = "sin enviar" (ResetCategoryHashes)
    If (Not forceSend) And newHash = GetCategoryHash(UserIndex, Category) Then
        Exit Sub
    End If
    Call SetCategoryHash(UserIndex, Category, newHash)
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageRadarCategory(PacketIdForCategory(Category), CByte(count), Xs, Ys))
    Exit Sub
SendCategoryUpdate_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.SendCategoryUpdate", Erl)
End Sub

' Envia paquetes vacios para las 9 categorias. Se usa al apagar el radar GM.
Public Sub SendRadarClearAll(ByVal UserIndex As Integer)
    On Error GoTo SendRadarClearAll_Err
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageRadarCategoryEmpty(ServerPacketID.eNpcRadar))
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageRadarCategoryEmpty(ServerPacketID.eNpcRadarNeutralNpc))
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageRadarCategoryEmpty(ServerPacketID.eNpcRadarQuestNpc))
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageRadarCategoryEmpty(ServerPacketID.eNpcRadarGuardiaReal))
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageRadarCategoryEmpty(ServerPacketID.eNpcRadarGuardiaCaos))
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageRadarCategoryEmpty(ServerPacketID.eNpcRadarMascota))
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageRadarCategoryEmpty(ServerPacketID.eNpcRadarUserGm))
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageRadarCategoryEmpty(ServerPacketID.eNpcRadarUserCiudadano))
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageRadarCategoryEmpty(ServerPacketID.eNpcRadarUserCriminal))
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageRadarCategoryEmpty(ServerPacketID.eNpcRadarUserImperial))
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageRadarCategoryEmpty(ServerPacketID.eNpcRadarUserCaos))
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageRadarCategoryEmpty(ServerPacketID.eNpcRadarPartyMember))
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageRadarCategoryEmpty(ServerPacketID.eNpcRadarClanMember))
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageRadarCategoryEmpty(ServerPacketID.eNpcRadarJefe))
    Call ResetAllHashes(UserIndex)
    Exit Sub
SendRadarClearAll_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.SendRadarClearAll", Erl)
End Sub

' ============================================================================
' Senders publicos (llamados por el game loop y el toggle)
' ============================================================================

' Perfil (NpcRadarActive=1): las categorias de flags.RadarMask, con radio y penetracion.
Public Sub SendNpcRadarUpdateProfile(ByVal UserIndex As Integer)
    On Error GoTo SendNpcRadarUpdateProfile_Err
    Dim mask As Long
    mask = UserList(UserIndex).flags.RadarMask
    Dim mapChanged As Boolean
    mapChanged = (UserList(UserIndex).Counters.NpcRadarLastMap <> UserList(UserIndex).pos.Map)
    Dim cats As Variant
    cats = Array(e_RadarCategory.eRadarHostile, e_RadarCategory.eRadarNeutralNpc, e_RadarCategory.eRadarGuardiaReal, _
                 e_RadarCategory.eRadarGuardiaCaos, e_RadarCategory.eRadarMascota, e_RadarCategory.eRadarUserCiudadano, _
                 e_RadarCategory.eRadarUserCriminal, e_RadarCategory.eRadarUserImperial, e_RadarCategory.eRadarUserCaos)
    Dim c As Long
    For c = LBound(cats) To UBound(cats)
        If (mask And RadarBit(cats(c))) <> 0 Then Call SendCategoryUpdate(UserIndex, cats(c), mapChanged)
    Next c
    ' Plan radar-v2 D9: quien ve hostiles ve tambien a los jefes (punto grande / icono propio).
    If (mask And (RadarBit(e_RadarCategory.eRadarHostile) Or RadarBit(e_RadarCategory.eRadarJefe))) <> 0 Then
        Call SendCategoryUpdate(UserIndex, e_RadarCategory.eRadarJefe, mapChanged)
    End If
    UserList(UserIndex).Counters.NpcRadarLastMap = UserList(UserIndex).pos.Map
    Exit Sub
SendNpcRadarUpdateProfile_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.SendNpcRadarUpdateProfile", Erl)
End Sub

' Para GMs (NpcRadarActive=2) - 8 categorias (QuestNpc esta en el loop siempre-activo).
Public Sub SendNpcRadarUpdateGm(ByVal UserIndex As Integer)
    On Error GoTo SendNpcRadarUpdateGm_Err
    Dim mapChanged As Boolean
    mapChanged = (UserList(UserIndex).Counters.NpcRadarLastMap <> UserList(UserIndex).pos.Map)
    Call SendCategoryUpdate(UserIndex, e_RadarCategory.eRadarHostile, mapChanged)
    Call SendCategoryUpdate(UserIndex, e_RadarCategory.eRadarJefe, mapChanged)
    Call SendCategoryUpdate(UserIndex, e_RadarCategory.eRadarNeutralNpc, mapChanged)
    Call SendCategoryUpdate(UserIndex, e_RadarCategory.eRadarGuardiaReal, mapChanged)
    Call SendCategoryUpdate(UserIndex, e_RadarCategory.eRadarGuardiaCaos, mapChanged)
    Call SendCategoryUpdate(UserIndex, e_RadarCategory.eRadarMascota, mapChanged)
    Call SendCategoryUpdate(UserIndex, e_RadarCategory.eRadarUserGm, mapChanged)
    Call SendCategoryUpdate(UserIndex, e_RadarCategory.eRadarUserCiudadano, mapChanged)
    Call SendCategoryUpdate(UserIndex, e_RadarCategory.eRadarUserCriminal, mapChanged)
    UserList(UserIndex).Counters.NpcRadarLastMap = UserList(UserIndex).pos.Map
    Exit Sub
SendNpcRadarUpdateGm_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.SendNpcRadarUpdateGm", Erl)
End Sub

' Fuerza un envio inmediato segun el modo del usuario.
Public Sub ForceNpcRadarUpdate(ByVal UserIndex As Integer)
    On Error GoTo ForceNpcRadarUpdate_Err
    If Not IsFeatureEnabled("npc_minimap_radar") Then Exit Sub
    Call ResetAllHashes(UserIndex)
    If UserList(UserIndex).flags.NpcRadarActive = 2 Then
        Call SendNpcRadarUpdateGm(UserIndex)
    ElseIf UserList(UserIndex).flags.NpcRadarActive = 1 Then
        Call SendNpcRadarUpdateProfile(UserIndex)
    End If
    UserList(UserIndex).Counters.NpcRadarTick = GetTickCountRaw()
    Exit Sub
ForceNpcRadarUpdate_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.ForceNpcRadarUpdate", Erl)
End Sub

' Tick del main loop. Refresca:
'  - QuestNpc: SIEMPRE para todos los usuarios logueados (con hash-dedup).
'  - Resto del radar: solo si NpcRadarActive > 0 (hechizo o GM).
Public Sub UpdateNpcRadarTick()
    On Error GoTo UpdateNpcRadarTick_Err
    If Not IsFeatureEnabled("npc_minimap_radar") Then Exit Sub
    Dim nowTick As Long
    nowTick = GetTickCountRaw()
    Dim i As Integer
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            ' Plan radar-v2 E1 (bug 1): el cambio de mapa se detecta aca, para todos. Antes solo lo
            ' anotaban los senders de radar activo y al jugador sin radar se le forzaba el reenvio de
            ' quest/simbolos/party/clan en cada tick. El cliente borra los puntos al cambiar de mapa:
            ' con los hashes en 0 cada categoria se reenvia una vez en su proximo tick.
            If UserList(i).Counters.NpcRadarLastMap <> UserList(i).pos.Map Then
                Call ResetCategoryHashes(i)
                UserList(i).Counters.NpcRadarLastMap = UserList(i).pos.Map
            End If
            ' QuestNpc siempre activo para todos los jugadores logueados.
            If RadarDue(nowTick, UserList(i).Counters.QuestRadarTick) Then
                Call SendCategoryUpdate(i, e_RadarCategory.eRadarQuestNpc, False)
                UserList(i).Counters.QuestRadarTick = nowTick
            End If
            ' NpcSymbol siempre activo para todos los jugadores logueados.
            If RadarDue(nowTick, UserList(i).Counters.NpcSymbolRadarTick) Then
                Call SendCategoryUpdate(i, e_RadarCategory.eRadarNpcSymbol, False)
                UserList(i).Counters.NpcSymbolRadarTick = nowTick
            End If
            ' Party/Clan radar siempre activo para todos los jugadores logueados.
            If RadarDue(nowTick, UserList(i).Counters.SocialRadarTick) Then
                Call SendCategoryUpdate(i, e_RadarCategory.eRadarPartyMember, False)
                Call SendCategoryUpdate(i, e_RadarCategory.eRadarClanMember, False)
                UserList(i).Counters.SocialRadarTick = nowTick
            End If
            ' Resto del radar: perfil (efectos + objetos equipados, recalculado cada tick) o GM.
            If RadarDue(nowTick, UserList(i).Counters.NpcRadarTick) Then
                If UserList(i).flags.NpcRadarActive <> 2 Then Call RefreshRadarProfile(i)
                If UserList(i).flags.NpcRadarActive = 2 Then
                    Call SendNpcRadarUpdateGm(i)
                ElseIf UserList(i).flags.NpcRadarActive = 1 Then
                    Call SendNpcRadarUpdateProfile(i)
                End If
                UserList(i).Counters.NpcRadarTick = nowTick
            End If
        End If
    Next i
    Exit Sub
UpdateNpcRadarTick_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.UpdateNpcRadarTick", Erl)
End Sub

' ============================================================================
' Motor de perfiles de radar (plan 2026-11-07-radar-v2, E3)
' ============================================================================
' Un perfil se define en datos y se aplica igual desde un efecto con duracion o un objeto:
'   EffectsOverTime.dat [EOTn] de Type=18/19 -> hechizos (EOTID) y consumibles (ApplyEffectId)
'   obj.dat [OBJn]                         -> mientras el objeto este equipado
' Claves: RadarVe=hostiles,neutrales,mascotas,guardias,ciudadanos,criminales,imperiales,caos,jefes
'         ("hostiles" incluye a los jefes; "jefes" solo a los jefes)
'         RadarRadio=tiles (0 = mapa entero)  RadarProtege=0-3  RadarPenetra=0-3

Public Function RadarBit(ByVal Category As Long) As Long
    RadarBit = CLng(2 ^ Category)
End Function

Private Function ClampRadarValue(ByVal value As Long, ByVal minValue As Long, ByVal maxValue As Long) As Long
    If value < minValue Then value = minValue
    If value > maxValue Then value = maxValue
    ClampRadarValue = value
End Function

Private Function ParseRadarCategories(ByVal text As String, ByVal Section As String) As Long
    On Error GoTo ParseRadarCategories_Err
    Dim tokens() As String
    Dim t As Long
    Dim token As String
    text = LCase$(Replace$(text, " ", ""))
    If LenB(text) = 0 Then Exit Function
    tokens = Split(text, ",")
    For t = LBound(tokens) To UBound(tokens)
        token = tokens(t)
        Select Case token
            Case "": ' coma sobrante
            Case "hostiles":   ParseRadarCategories = ParseRadarCategories Or RadarBit(e_RadarCategory.eRadarHostile)
            Case "neutrales":  ParseRadarCategories = ParseRadarCategories Or RadarBit(e_RadarCategory.eRadarNeutralNpc)
            Case "mascotas":   ParseRadarCategories = ParseRadarCategories Or RadarBit(e_RadarCategory.eRadarMascota)
            Case "guardias":   ParseRadarCategories = ParseRadarCategories Or RadarBit(e_RadarCategory.eRadarGuardiaReal) Or RadarBit(e_RadarCategory.eRadarGuardiaCaos)
            Case "ciudadanos": ParseRadarCategories = ParseRadarCategories Or RadarBit(e_RadarCategory.eRadarUserCiudadano)
            Case "criminales": ParseRadarCategories = ParseRadarCategories Or RadarBit(e_RadarCategory.eRadarUserCriminal)
            Case "imperiales": ParseRadarCategories = ParseRadarCategories Or RadarBit(e_RadarCategory.eRadarUserImperial)
            Case "caos":       ParseRadarCategories = ParseRadarCategories Or RadarBit(e_RadarCategory.eRadarUserCaos)
            Case "jefes":      ParseRadarCategories = ParseRadarCategories Or RadarBit(e_RadarCategory.eRadarJefe)
            Case Else
                Call TraceError(0, "RadarVe con categoria desconocida '" & token & "' en [" & Section & "]", "Radar.ParseRadarCategories", 0)
        End Select
    Next t
    Exit Function
ParseRadarCategories_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.ParseRadarCategories", Erl)
End Function

' Lee el perfil de una seccion de EffectsOverTime.dat o de obj.dat. Claves ausentes = 0.
Public Sub ReadRadarProfile(ByVal Leer As clsIniManager, ByVal Section As String, ByRef profile As t_RadarProfile)
    On Error GoTo ReadRadarProfile_Err
    profile.Mask = ParseRadarCategories(Leer.GetValue(Section, "RadarVe"), Section)
    profile.Radius = ClampRadarValue(val(Leer.GetValue(Section, "RadarRadio")), 0, 100)
    profile.Protection = ClampRadarValue(val(Leer.GetValue(Section, "RadarProtege")), 0, 3)
    profile.Penetration = ClampRadarValue(val(Leer.GetValue(Section, "RadarPenetra")), 0, 3)
    Exit Sub
ReadRadarProfile_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.ReadRadarProfile " & Section, Erl)
End Sub

' Compatibilidad con los datos viejos, despues de cargar hechizos y EOT:
' - un hechizo con DetectaCiuda/Crimi/Impe/Caos le pasa esas categorias a su EOT si este no tiene RadarVe;
' - un EOT de tipo 18 sin perfil detecta hostiles (era "Detectar el Mal").
Public Sub ApplyLegacyRadarProfiles()
    On Error GoTo ApplyLegacyRadarProfiles_Err
    Dim h As Long
    Dim eot As Long
    Dim legacy As Byte
    For h = 1 To NumeroHechizos
        legacy = Hechizos(h).UserDetectMask
        eot = Hechizos(h).EotId
        If legacy <> 0 And eot >= LBound(EffectOverTime) And eot <= UBound(EffectOverTime) Then
            If EffectOverTime(eot).Radar.Mask = 0 Then
                If (legacy And 1) <> 0 Then EffectOverTime(eot).Radar.Mask = EffectOverTime(eot).Radar.Mask Or RadarBit(e_RadarCategory.eRadarUserCiudadano)
                If (legacy And 2) <> 0 Then EffectOverTime(eot).Radar.Mask = EffectOverTime(eot).Radar.Mask Or RadarBit(e_RadarCategory.eRadarUserCriminal)
                If (legacy And 4) <> 0 Then EffectOverTime(eot).Radar.Mask = EffectOverTime(eot).Radar.Mask Or RadarBit(e_RadarCategory.eRadarUserImperial)
                If (legacy And 8) <> 0 Then EffectOverTime(eot).Radar.Mask = EffectOverTime(eot).Radar.Mask Or RadarBit(e_RadarCategory.eRadarUserCaos)
            End If
        End If
    Next h
    For eot = LBound(EffectOverTime) To UBound(EffectOverTime)
        With EffectOverTime(eot)
            If .Type = e_EffectOverTimeType.eMinimapRadar And .Radar.Mask = 0 And .Radar.Protection = 0 Then
                .Radar.Mask = RadarBit(e_RadarCategory.eRadarHostile)
            End If
        End With
    Next eot
    Exit Sub
ApplyLegacyRadarProfiles_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.ApplyLegacyRadarProfiles", Erl)
End Sub

Private Function IsRadarEffect(ByRef Effect As IBaseEffectOverTime) As Boolean
    If Effect Is Nothing Then Exit Function
    If Effect.RemoveMe Then Exit Function
    IsRadarEffect = (Effect.TypeId = e_EffectOverTimeType.eMinimapRadar Or Effect.TypeId = e_EffectOverTimeType.eMinimapUserDetect)
End Function

' Plan radar-v2 D6: al aplicar un efecto de radar, los efectos anteriores que detectan se van si el nuevo
' detecta, y los que protegen se van si el nuevo protege (un efecto con las dos cosas se va entero).
Public Sub MarkReplacedRadarEffects(ByVal UserIndex As Integer, ByVal NewEotId As Integer, ByVal KeepUniqueId As Long)
    On Error GoTo MarkReplacedRadarEffects_Err
    Dim k As Long
    Dim newProfile As t_RadarProfile
    Dim oldProfile As t_RadarProfile
    newProfile = EffectOverTime(NewEotId).Radar
    With UserList(UserIndex).EffectOverTime
        For k = 0 To .EffectCount - 1
            If IsRadarEffect(.EffectList(k)) Then
                If .EffectList(k).UniqueId <> KeepUniqueId Then
                    oldProfile = EffectOverTime(.EffectList(k).EotId).Radar
                    If (newProfile.Mask <> 0 And oldProfile.Mask <> 0) Or (newProfile.Protection > 0 And oldProfile.Protection > 0) Then
                        .EffectList(k).RemoveMe = True
                    End If
                End If
            End If
        Next k
    End With
    Exit Sub
MarkReplacedRadarEffects_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.MarkReplacedRadarEffects", Erl)
End Sub

Private Sub MergeRadarProfile(ByRef src As t_RadarProfile, ByRef mask As Long, ByRef radius As Integer, ByRef wholeMap As Boolean, ByRef penetration As Byte, ByRef protection As Byte)
    If src.Mask <> 0 Then
        mask = mask Or src.Mask
        If src.Radius = 0 Then
            wholeMap = True
        ElseIf src.Radius > radius Then
            radius = src.Radius
        End If
        If src.Penetration > penetration Then penetration = src.Penetration
    End If
    If src.Protection > protection Then protection = src.Protection
End Sub

Private Sub MergeObjRadar(ByVal ObjIndex As Integer, ByRef mask As Long, ByRef radius As Integer, ByRef wholeMap As Boolean, ByRef penetration As Byte, ByRef protection As Byte)
    If ObjIndex <= 0 Or ObjIndex > UBound(ObjData) Then Exit Sub
    Call MergeRadarProfile(ObjData(ObjIndex).Radar, mask, radius, wholeMap, penetration, protection)
End Sub

' Vacia el perfil guardado: el proximo RefreshRadarProfile lo vuelve a calcular y avisa al cliente.
Public Sub ResetRadarProfileState(ByVal UserIndex As Integer)
    With UserList(UserIndex).flags
        .RadarMask = 0
        .RadarRadius = 0
        .RadarPenetration = 0
        .RadarProtection = 0
    End With
End Sub

' Junta el perfil de los efectos de radar activos y de los objetos equipados. Si cambio lo que el
' jugador detecta, limpia el minimapa y manda lo nuevo. El GM (NpcRadarActive=2) no cambia de modo.
Public Sub RefreshRadarProfile(ByVal UserIndex As Integer)
    On Error GoTo RefreshRadarProfile_Err
    Dim mask As Long
    Dim radius As Integer
    Dim wholeMap As Boolean
    Dim penetration As Byte
    Dim protection As Byte
    Dim k As Long
    With UserList(UserIndex).EffectOverTime
        For k = 0 To .EffectCount - 1
            If IsRadarEffect(.EffectList(k)) Then
                Call MergeRadarProfile(EffectOverTime(.EffectList(k).EotId).Radar, mask, radius, wholeMap, penetration, protection)
            End If
        Next k
    End With
    With UserList(UserIndex).invent
        Call MergeObjRadar(.EquippedWeaponObjIndex, mask, radius, wholeMap, penetration, protection)
        Call MergeObjRadar(.EquippedArmorObjIndex, mask, radius, wholeMap, penetration, protection)
        Call MergeObjRadar(.EquippedShieldObjIndex, mask, radius, wholeMap, penetration, protection)
        Call MergeObjRadar(.EquippedHelmetObjIndex, mask, radius, wholeMap, penetration, protection)
        Call MergeObjRadar(.EquippedRingAccesoryObjIndex, mask, radius, wholeMap, penetration, protection)
        Call MergeObjRadar(.EquippedAmuletAccesoryObjIndex, mask, radius, wholeMap, penetration, protection)
        Call MergeObjRadar(.EquippedBackpackObjIndex, mask, radius, wholeMap, penetration, protection)
    End With
    If wholeMap Then radius = 0
    Dim hadMask As Boolean
    With UserList(UserIndex).flags
        .RadarProtection = protection
        If .NpcRadarActive = 2 Then Exit Sub
        If mask = .RadarMask And radius = .RadarRadius And penetration = .RadarPenetration Then Exit Sub
        hadMask = (.RadarMask <> 0)
        .RadarMask = mask
        .RadarRadius = radius
        .RadarPenetration = penetration
        If mask <> 0 Then
            .NpcRadarActive = 1
        Else
            .NpcRadarActive = 0
        End If
        If Not .UserLogged Then Exit Sub
    End With
    If Not IsFeatureEnabled("npc_minimap_radar") Then Exit Sub
    If hadMask Then Call SendRadarClearAll(UserIndex)
    If mask <> 0 Then Call ForceNpcRadarUpdate(UserIndex)
    Exit Sub
RefreshRadarProfile_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.RefreshRadarProfile", Erl)
End Sub

' Radio del perfil: solo aplica al radar por perfil (NpcRadarActive=1). 0 = mapa entero.
Private Function InRadarRange(ByVal ObserverIndex As Integer, ByVal x As Long, ByVal y As Long) As Boolean
    InRadarRange = True
    With UserList(ObserverIndex)
        If .flags.NpcRadarActive <> 1 Then Exit Function
        If .flags.RadarRadius <= 0 Then Exit Function
        If Abs(x - CLng(.pos.x)) > .flags.RadarRadius Or Abs(y - CLng(.pos.y)) > .flags.RadarRadius Then InRadarRange = False
    End With
End Function

' ============================================================================
' Marcar un punto (plan 2026-11-07-radar-v2, E6)
' ============================================================================
' El cliente manda (x, y); el server se lo reenvia a la party del mismo mapa con el nombre de quien marco.
' Sin party, fuera de rango, antes de 3 s o con el radar apagado por toggle: se descarta en silencio.
' En eventos y retos no cruza bandos (GrupoCruzaBandos, como el radar de party).
Public Sub HandleMinimapPing(ByVal UserIndex As Integer)
    On Error GoTo HandleMinimapPing_Err
    Dim x As Byte
    Dim y As Byte
    x = reader.ReadInt8
    y = reader.ReadInt8
    If Not IsFeatureEnabled("npc_minimap_radar") Then Exit Sub
    Dim nowTick As Long
    Dim leaderIndex As Integer
    Dim userMap As Integer
    With UserList(UserIndex)
        If Not .flags.UserLogged Then Exit Sub
        If x < 1 Or x > 100 Or y < 1 Or y > 100 Then Exit Sub
        If Not .Grupo.EnGrupo Then Exit Sub
        nowTick = GetTickCountRaw()
        If CooldownActive(nowTick, .Counters.MinimapPingTick, MINIMAP_PING_COOLDOWN_MS) Then Exit Sub
        .Counters.MinimapPingTick = NonZeroTickFrom(nowTick)
        leaderIndex = .Grupo.Lider.ArrayIndex
        userMap = .pos.Map
    End With
    If leaderIndex <= 0 Then Exit Sub
    Dim senderName As String
    senderName = GetUserDisplayName(UserIndex)
    Dim i As Integer
    Dim memberIndex As Integer
    For i = 1 To UserList(leaderIndex).Grupo.CantidadMiembros
        memberIndex = UserList(leaderIndex).Grupo.Miembros(i).ArrayIndex
        If memberIndex > 0 And memberIndex <> UserIndex Then
            If UserList(memberIndex).flags.UserLogged And UserList(memberIndex).pos.Map = userMap Then
                If Not UserMod.GrupoCruzaBandos(UserIndex, memberIndex) Then
                    Call WriteMinimapPingShow(memberIndex, x, y, senderName)
                End If
            End If
        End If
    Next i
    Exit Sub
HandleMinimapPing_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.HandleMinimapPing", Erl)
End Sub
