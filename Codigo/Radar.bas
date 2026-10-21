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
' - Para jugadores normales (NpcRadarActive=1): solo categoria Hostile.
' - Para GMs (NpcRadarActive=2): 9 categorias (Hostile, NeutralNpc, QuestNpc,
'   GuardiaReal, GuardiaCaos, Mascota, UserGm, UserCiudadano, UserCriminal).
Option Explicit

' NPC_RADAR_TICK_MS (300 ms) vive ahora en modTickRules, junto a RadarDue.
Private Const MAX_RADAR_ENTITIES As Integer = 100

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
        ElseIf .Faccion.Status = e_Facciones.Criminal Then
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
    End Select
End Function

' ============================================================================
' Recoleccion
' ============================================================================

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
    For NpcIndex = 1 To LastNPC
        With NpcList(NpcIndex)
            If .pos.Map = userMap Then
                cat = ClassifyNpc(NpcIndex)
                If cat = Category Then
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
    Next NpcIndex
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
                If Category = e_RadarCategory.eRadarUserGm Then
                    cat = ClassifyUser(ObserverIndex, i)
                Else
                    cat = ClassifyUserDetailed(ObserverIndex, i)
                End If
                If cat = Category Then
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
        If memberIndex > 0 And memberIndex <> ObserverIndex Then
            If UserList(memberIndex).flags.UserLogged Then
                If UserList(memberIndex).pos.Map = userMap Then
                    Dim hidden As Boolean
                    hidden = (UserList(memberIndex).flags.AdminInvisible = 1) Or (UserList(memberIndex).flags.invisible > 0) Or (UserList(memberIndex).flags.Oculto = 1)
                    Dim show As Boolean
                    show = True
                    If hidden Then
                        If CompararPrivilegiosUser(ObserverIndex, memberIndex) < 0 Then show = False
                    End If
                    If show Then
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
    Dim i As Integer
    Dim px As Byte, py As Byte
    For i = 1 To LastUser
        If i <> ObserverIndex Then
            If UserList(i).flags.UserLogged Then
                If UserList(i).GuildIndex = observerGuild Then
                    If UserList(i).pos.Map = userMap Then
                        If Not IsInObserverParty(ObserverIndex, i) Then
                            Dim hidden As Boolean
                            hidden = (UserList(i).flags.AdminInvisible = 1) Or (UserList(i).flags.invisible > 0) Or (UserList(i).flags.Oculto = 1)
                            Dim show As Boolean
                            show = True
                            If hidden Then
                                If CompararPrivilegiosUser(ObserverIndex, i) < 0 Then show = False
                            End If
                            If show Then
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
                    ' Disponible: no la hizo, cumple prerequisitos
                    If qIdx >= LBound(QuestList) And qIdx <= UBound(QuestList) Then
                        Dim qData As t_Quest
                        qData = QuestList(qIdx)
                        Dim okClass As Boolean
                        If qData.RequiredClassesCount > 0 Then
                            Dim ci As Integer
                            For ci = 1 To qData.RequiredClassesCount
                                If UserList(UserIndex).clase = qData.RequiredClass(ci) Then
                                    okClass = True
                                    Exit For
                                End If
                            Next ci
                        Else
                            okClass = True
                        End If
                        If okClass Then
                            If Not UserDoneQuest(UserIndex, qIdx) Then
                                Dim okPrev As Boolean
                                If qData.RequiredQuest = 0 Then
                                    okPrev = True
                                Else
                                    okPrev = UserDoneQuest(UserIndex, qData.RequiredQuest)
                                End If
                                If okPrev Then
                                    If UserList(UserIndex).Stats.ELV >= qData.RequiredLevel Then
                                        hasAvailable = True
                                    End If
                                End If
                            End If
                        End If
                    End If
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
    For NpcIndex = 1 To LastNPC
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
    Next NpcIndex
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
    For NpcIndex = 1 To LastNPC
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
    Next NpcIndex
    CollectNpcSymbolRadar = h
    Exit Function
CollectNpcSymbolRadar_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.CollectNpcSymbolRadar", Erl)
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
    End Select
End Sub

Private Sub ResetAllHashes(ByVal UserIndex As Integer)
    With UserList(UserIndex).Counters
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
        .NpcRadarLastMap = -1
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
        If (Not forceSend) And newHash = GetCategoryHash(UserIndex, Category) Then
            Exit Sub
        End If
        Call SetCategoryHash(UserIndex, Category, newHash)
        Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageRadarQuestNpc(PacketIdForCategory(Category), CByte(count), Xs, Ys, Ss2))
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
    Call ResetAllHashes(UserIndex)
    Exit Sub
SendRadarClearAll_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.SendRadarClearAll", Erl)
End Sub

' ============================================================================
' Senders publicos (llamados por el game loop y el toggle)
' ============================================================================

' Para jugadores normales (NpcRadarActive=1) - solo hostiles.
Public Sub SendNpcRadarUpdate(ByVal UserIndex As Integer)
    On Error GoTo SendNpcRadarUpdate_Err
    Dim mapChanged As Boolean
    mapChanged = (UserList(UserIndex).Counters.NpcRadarLastMap <> UserList(UserIndex).pos.Map)
    Call SendCategoryUpdate(UserIndex, e_RadarCategory.eRadarHostile, mapChanged)
    UserList(UserIndex).Counters.NpcRadarLastMap = UserList(UserIndex).pos.Map
    Exit Sub
SendNpcRadarUpdate_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.SendNpcRadarUpdate", Erl)
End Sub

' Para GMs (NpcRadarActive=2) - 8 categorias (QuestNpc esta en el loop siempre-activo).
Public Sub SendNpcRadarUpdateGm(ByVal UserIndex As Integer)
    On Error GoTo SendNpcRadarUpdateGm_Err
    Dim mapChanged As Boolean
    mapChanged = (UserList(UserIndex).Counters.NpcRadarLastMap <> UserList(UserIndex).pos.Map)
    Call SendCategoryUpdate(UserIndex, e_RadarCategory.eRadarHostile, mapChanged)
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

' Para hechizo de deteccion de usuarios (NpcRadarActive=3) - solo categorias en UserDetectMask.
Public Sub SendNpcRadarUpdateUserDetect(ByVal UserIndex As Integer)
    On Error GoTo SendNpcRadarUpdateUserDetect_Err
    Dim mask As Byte
    mask = UserList(UserIndex).flags.UserDetectMask
    Dim mapChanged As Boolean
    mapChanged = (UserList(UserIndex).Counters.NpcRadarLastMap <> UserList(UserIndex).pos.Map)
    If (mask And 1) <> 0 Then Call SendCategoryUpdate(UserIndex, e_RadarCategory.eRadarUserCiudadano, mapChanged)
    If (mask And 2) <> 0 Then Call SendCategoryUpdate(UserIndex, e_RadarCategory.eRadarUserCriminal, mapChanged)
    If (mask And 4) <> 0 Then Call SendCategoryUpdate(UserIndex, e_RadarCategory.eRadarUserImperial, mapChanged)
    If (mask And 8) <> 0 Then Call SendCategoryUpdate(UserIndex, e_RadarCategory.eRadarUserCaos, mapChanged)
    UserList(UserIndex).Counters.NpcRadarLastMap = UserList(UserIndex).pos.Map
    Exit Sub
SendNpcRadarUpdateUserDetect_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.SendNpcRadarUpdateUserDetect", Erl)
End Sub

' Fuerza un envio inmediato segun el modo del usuario.
Public Sub ForceNpcRadarUpdate(ByVal UserIndex As Integer)
    On Error GoTo ForceNpcRadarUpdate_Err
    If Not IsFeatureEnabled("npc_minimap_radar") Then Exit Sub
    Call ResetAllHashes(UserIndex)
    If UserList(UserIndex).flags.NpcRadarActive = 2 Then
        Call SendNpcRadarUpdateGm(UserIndex)
    ElseIf UserList(UserIndex).flags.NpcRadarActive = 3 Then
        Call SendNpcRadarUpdateUserDetect(UserIndex)
    Else
        Call SendNpcRadarUpdate(UserIndex)
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
    Dim mapChanged As Boolean
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            ' QuestNpc siempre activo para todos los jugadores logueados.
            If RadarDue(nowTick, UserList(i).Counters.QuestRadarTick) Then
                mapChanged = (UserList(i).Counters.NpcRadarLastMap <> UserList(i).pos.Map)
                Call SendCategoryUpdate(i, e_RadarCategory.eRadarQuestNpc, mapChanged)
                UserList(i).Counters.QuestRadarTick = nowTick
            End If
            ' NpcSymbol siempre activo para todos los jugadores logueados.
            If RadarDue(nowTick, UserList(i).Counters.NpcSymbolRadarTick) Then
                mapChanged = (UserList(i).Counters.NpcRadarLastMap <> UserList(i).pos.Map)
                Call SendCategoryUpdate(i, e_RadarCategory.eRadarNpcSymbol, mapChanged)
                UserList(i).Counters.NpcSymbolRadarTick = nowTick
            End If
            ' Party/Clan radar siempre activo para todos los jugadores logueados.
            If RadarDue(nowTick, UserList(i).Counters.SocialRadarTick) Then
                mapChanged = (UserList(i).Counters.NpcRadarLastMap <> UserList(i).pos.Map)
                Call SendCategoryUpdate(i, e_RadarCategory.eRadarPartyMember, mapChanged)
                Call SendCategoryUpdate(i, e_RadarCategory.eRadarClanMember, mapChanged)
                UserList(i).Counters.SocialRadarTick = nowTick
            End If
            ' Resto del radar: solo si esta activo por hechizo (1) o GM (2).
            If UserList(i).flags.NpcRadarActive > 0 Then
                If RadarDue(nowTick, UserList(i).Counters.NpcRadarTick) Then
                    If UserList(i).flags.NpcRadarActive = 2 Then
                        Call SendNpcRadarUpdateGm(i)
                    ElseIf UserList(i).flags.NpcRadarActive = 3 Then
                        Call SendNpcRadarUpdateUserDetect(i)
                    Else
                        Call SendNpcRadarUpdate(i)
                    End If
                    UserList(i).Counters.NpcRadarTick = nowTick
                End If
            End If
        End If
    Next i
    Exit Sub
UpdateNpcRadarTick_Err:
    Call TraceError(Err.Number, Err.Description, "Radar.UpdateNpcRadarTick", Erl)
End Sub
