Attribute VB_Name = "modProfesiones"
' Argentum 20 Game Server
'
'    Copyright (C) 2026 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
Option Explicit

' Sistema de Profesiones Aprendibles.
' Gateado por el feature toggle "professions_learnable".
' Ver: ia/plans/2026-04-22-sistema-profesiones-aprendibles.md

Public Const PROF_MAX_SLOTS As Byte = 7
Public Const PROF_MIN_ID    As Byte = 17
Public Const PROF_MAX_ID    As Byte = 23

Private Type t_ProfesionConfig
    SkillId            As Integer
    NpcMaestroId       As Integer
    ItemManualId       As Integer
    ItemPocionOlvidoId As Integer
    SkillParaMax       As Integer
    CantidadMin        As Integer
    CantidadMax        As Integer
    ProbBase           As Single
    ProbPorSkill       As Single
    CantidadFija       As Integer
    EsExtraccion       As Boolean
    PenalizacionZonaSegura As Integer
    Cargada            As Boolean
End Type

Private ProfesionConfig(PROF_MIN_ID To PROF_MAX_ID) As t_ProfesionConfig
Private MaxOlvidosPorPersonaje    As Byte
Private MaxProfesionesAprendidas  As Byte

Public Sub LoadProfesionesConfig()
    On Error GoTo LoadProfesionesConfig_Err
    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    Call Lector.Initialize(App.Path & "\profesiones.ini")
    MaxOlvidosPorPersonaje = CByte(val(Lector.GetValue("General", "MaxOlvidosPorPersonaje")))
    MaxProfesionesAprendidas = CByte(val(Lector.GetValue("General", "MaxProfesionesAprendidas")))
    ' MaxOlvidosPorPersonaje = 0 => olvidos ILIMITADOS (no se fuerza default)
    If MaxProfesionesAprendidas = 0 Then MaxProfesionesAprendidas = 2
    If MaxProfesionesAprendidas > PROF_MAX_SLOTS Then MaxProfesionesAprendidas = PROF_MAX_SLOTS
    Call LoadSeccionProfesion(Lector, "Carpinteria", e_Skill.Carpinteria, False)
    Call LoadSeccionProfesion(Lector, "Herreria", e_Skill.Herreria, False)
    Call LoadSeccionProfesion(Lector, "Sastreria", e_Skill.Sastreria, False)
    Call LoadSeccionProfesion(Lector, "Alquimia", e_Skill.Alquimia, False)
    Call LoadSeccionProfesion(Lector, "Pesca", e_Skill.Pescar, True)
    Call LoadSeccionProfesion(Lector, "Tala", e_Skill.Talar, True)
    Call LoadSeccionProfesion(Lector, "Mineria", e_Skill.Mineria, True)
    Set Lector = Nothing
    Exit Sub
LoadProfesionesConfig_Err:
    Set Lector = Nothing
    Call TraceError(Err.Number, Err.Description, "modProfesiones.LoadProfesionesConfig", Erl)
End Sub

Private Sub LoadSeccionProfesion(ByRef Lector As clsIniManager, ByVal seccion As String, ByVal skillId As Integer, ByVal esExtraccion As Boolean)
    On Error GoTo LoadSeccionProfesion_Err
    With ProfesionConfig(skillId)
        .SkillId = skillId
        .NpcMaestroId = CInt(val(Lector.GetValue(seccion, "NpcMaestroId")))
        .ItemManualId = CInt(val(Lector.GetValue(seccion, "ItemManualId")))
        .ItemPocionOlvidoId = CInt(val(Lector.GetValue(seccion, "ItemPocionOlvidoId")))
        .EsExtraccion = esExtraccion
        If esExtraccion Then
            .SkillParaMax = CInt(val(Lector.GetValue(seccion, "SkillParaMax")))
            .CantidadMin = CInt(val(Lector.GetValue(seccion, "CantidadMin")))
            .CantidadMax = CInt(val(Lector.GetValue(seccion, "CantidadMax")))
            .ProbBase = CSng(val(Lector.GetValue(seccion, "ProbBase")))
            .ProbPorSkill = CSng(val(Lector.GetValue(seccion, "ProbPorSkill")))
            .PenalizacionZonaSegura = CInt(val(Lector.GetValue(seccion, "PenalizacionZonaSegura", "0")))
            If .SkillParaMax <= 0 Then .SkillParaMax = 100
            If .PenalizacionZonaSegura < 0 Then .PenalizacionZonaSegura = 0
            If .PenalizacionZonaSegura > 100 Then .PenalizacionZonaSegura = 100
        Else
            .CantidadFija = CInt(val(Lector.GetValue(seccion, "CantidadFija")))
            If .CantidadFija = 0 Then .CantidadFija = 1
        End If
        .Cargada = True
    End With
    Exit Sub
LoadSeccionProfesion_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.LoadSeccionProfesion", Erl)
End Sub

Public Function NombreProfesion(ByVal ProfesionId As Integer) As String
    If ProfesionId < LBound(SkillsNames) Or ProfesionId > UBound(SkillsNames) Then
        NombreProfesion = "?"
    Else
        NombreProfesion = SkillsNames(ProfesionId)
    End If
End Function

Public Function TieneProfesionAprendida(ByVal UserIndex As Integer, ByVal ProfesionId As Integer) As Boolean
    On Error GoTo TieneProfesionAprendida_Err
    TieneProfesionAprendida = False
    If Not IsFeatureEnabled("professions_learnable") Then
        TieneProfesionAprendida = True
        Exit Function
    End If
    ' Bypass total de staff (plan 04.001): un GM trabaja y usa toda herramienta siempre.
    ' Los chequeos internos de aprender/olvidar usan TieneProfesionEnSlots (sin bypass).
    If EsGM(UserIndex) Then
        TieneProfesionAprendida = True
        Exit Function
    End If
    TieneProfesionAprendida = TieneProfesionEnSlots(UserIndex, ProfesionId)
    Exit Function
TieneProfesionAprendida_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.TieneProfesionAprendida", Erl)
End Function

Private Function TieneProfesionEnSlots(ByVal UserIndex As Integer, ByVal ProfesionId As Integer) As Boolean
    ' Mira los slots reales, sin bypass de staff ni de toggle.
    Dim i As Byte
    TieneProfesionEnSlots = False
    For i = 1 To PROF_MAX_SLOTS
        If UserList(UserIndex).Professions(i) = ProfesionId Then
            TieneProfesionEnSlots = True
            Exit Function
        End If
    Next i
End Function

Public Function SlotsOcupados(ByVal UserIndex As Integer) As Byte
    Dim i     As Byte
    Dim count As Byte
    count = 0
    For i = 1 To PROF_MAX_SLOTS
        If UserList(UserIndex).Professions(i) > 0 Then count = count + 1
    Next i
    SlotsOcupados = count
End Function

Public Function PuedeAprenderProfesion(ByVal UserIndex As Integer, ByVal ProfesionId As Integer, Optional ByVal EsAccionDeStaff As Boolean = False) As Boolean
    On Error GoTo PuedeAprenderProfesion_Err
    PuedeAprenderProfesion = False
    If Not IsFeatureEnabled("professions_learnable") Then Exit Function
    If ProfesionId < PROF_MIN_ID Or ProfesionId > PROF_MAX_ID Then Exit Function
    If TieneProfesionEnSlots(UserIndex, ProfesionId) Then
        If Not EsAccionDeStaff Then Call WriteLocaleMsg(UserIndex, MSG_PROF_YA_APRENDIDA, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    ' GMs y acciones de staff exentos de los limites de config (plan 04.001).
    ' El tope fisico PROF_MAX_SLOTS lo garantiza AprenderProfesion al buscar slot libre.
    If Not (EsGM(UserIndex) Or EsAccionDeStaff) Then
        ' El tope de olvidos frena el re-aprendizaje por si solo (antes exigia ademas
        ' slots llenos, condicion que la guarda siguiente ya cubria: nunca disparaba).
        If MaxOlvidosPorPersonaje > 0 And UserList(UserIndex).ProfessionForgotCount >= MaxOlvidosPorPersonaje Then
            Call WriteLocaleMsg(UserIndex, MSG_PROF_SLOT_BLOQUEADO_POR_OLVIDOS, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        If SlotsOcupados(UserIndex) >= MaxProfesionesAprendidas Then
            Call WriteLocaleMsg(UserIndex, MSG_PROF_YA_TIENE_2, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    End If
    PuedeAprenderProfesion = True
    Exit Function
PuedeAprenderProfesion_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.PuedeAprenderProfesion", Erl)
End Function

Public Function PuedeOlvidarProfesion(ByVal UserIndex As Integer, ByVal ProfesionId As Integer, Optional ByVal EsAccionDeStaff As Boolean = False) As Boolean
    On Error GoTo PuedeOlvidarProfesion_Err
    PuedeOlvidarProfesion = False
    If Not IsFeatureEnabled("professions_learnable") Then Exit Function
    If ProfesionId < PROF_MIN_ID Or ProfesionId > PROF_MAX_ID Then Exit Function
    If Not TieneProfesionEnSlots(UserIndex, ProfesionId) Then
        If Not EsAccionDeStaff Then Call WriteLocaleMsg(UserIndex, MSG_PROF_NO_APRENDIDA, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If Not (EsGM(UserIndex) Or EsAccionDeStaff) Then
        If MaxOlvidosPorPersonaje > 0 And UserList(UserIndex).ProfessionForgotCount >= MaxOlvidosPorPersonaje Then
            Call WriteLocaleMsg(UserIndex, MSG_PROF_MAX_OLVIDOS_PERSONAJE, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    End If
    PuedeOlvidarProfesion = True
    Exit Function
PuedeOlvidarProfesion_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.PuedeOlvidarProfesion", Erl)
End Function

Public Function AprenderProfesion(ByVal UserIndex As Integer, ByVal ProfesionId As Integer, Optional ByVal EsAccionDeStaff As Boolean = False) As Boolean
    On Error GoTo AprenderProfesion_Err
    AprenderProfesion = False
    If Not IsFeatureEnabled("professions_learnable") Then Exit Function
    If Not PuedeAprenderProfesion(UserIndex, ProfesionId, EsAccionDeStaff) Then Exit Function
    Dim i As Byte
    Dim almacenada As Boolean
    almacenada = False
    For i = 1 To PROF_MAX_SLOTS
        If UserList(UserIndex).Professions(i) = 0 Then
            UserList(UserIndex).Professions(i) = ProfesionId
            almacenada = True
            Exit For
        End If
    Next i
    If Not almacenada Then
        ' Sin slot libre (limite real = PROF_MAX_SLOTS). No reportar exito ni consumir el item.
        Call WriteLocaleMsg(UserIndex, MSG_PROF_YA_TIENE_2, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    Call Execute("INSERT OR REPLACE INTO user_professions (user_id, profession_id, learned_at) VALUES (?, ?, ?);", UserList(UserIndex).Id, CInt(ProfesionId), CLng(GetTickCountRaw() \ 1000))
    Call RefrescarHerramientasInventario(UserIndex, ProfesionId)
    Call WriteProfessionsUpdate(UserIndex)
    Call WriteLocaleMsg(UserIndex, MSG_PROF_APRENDIDA_OK, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO, NombreProfesion(ProfesionId))
    AprenderProfesion = True
    Exit Function
AprenderProfesion_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.AprenderProfesion", Erl)
End Function

Public Function OlvidarProfesion(ByVal UserIndex As Integer, ByVal ProfesionId As Integer, Optional ByVal EsAccionDeStaff As Boolean = False) As Boolean
    On Error GoTo OlvidarProfesion_Err
    OlvidarProfesion = False
    If Not IsFeatureEnabled("professions_learnable") Then Exit Function
    If Not PuedeOlvidarProfesion(UserIndex, ProfesionId, EsAccionDeStaff) Then Exit Function
    Dim i            As Byte
    For i = 1 To PROF_MAX_SLOTS
        If UserList(UserIndex).Professions(i) = ProfesionId Then
            UserList(UserIndex).Professions(i) = 0
            Exit For
        End If
    Next i
    ' T5 (usuario 2026-06-03): olvidar quita el permiso pero CONSERVA el skill ya progresado.
    ' Las acciones de staff no queman el contador de olvidos del objetivo (plan 04.001).
    If Not EsAccionDeStaff Then
        ' Saturar en 255: el campo es Byte y el olvido numero 256 desbordaba (error 6).
        If UserList(UserIndex).ProfessionForgotCount < 255 Then
            UserList(UserIndex).ProfessionForgotCount = UserList(UserIndex).ProfessionForgotCount + 1
        End If
        Call Execute("UPDATE user SET profession_forgot_count = ? WHERE id = ?;", CLng(UserList(UserIndex).ProfessionForgotCount), UserList(UserIndex).Id)
    End If
    Call Execute("DELETE FROM user_professions WHERE user_id = ? AND profession_id = ?;", UserList(UserIndex).Id, CInt(ProfesionId))
    Call RefrescarHerramientasInventario(UserIndex, ProfesionId)
    Call WriteProfessionsUpdate(UserIndex)
    Call WriteLocaleMsg(UserIndex, MSG_PROF_OLVIDADA_OK, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO, NombreProfesion(ProfesionId))
    OlvidarProfesion = True
    Exit Function
OlvidarProfesion_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.OlvidarProfesion", Erl)
End Function

Public Function CalcularExitoExtraccion(ByVal UserIndex As Integer, ByVal ProfesionId As Integer) As Boolean
    On Error GoTo CalcularExitoExtraccion_Err
    CalcularExitoExtraccion = True
    If ProfesionId < PROF_MIN_ID Or ProfesionId > PROF_MAX_ID Then Exit Function
    If Not ProfesionConfig(ProfesionId).EsExtraccion Then Exit Function
    Dim skillActual As Integer
    Dim prob        As Single
    skillActual = UserList(UserIndex).Stats.UserSkills(ProfesionId)
    With ProfesionConfig(ProfesionId)
        prob = .ProbBase + (.ProbPorSkill * CSng(skillActual))
    End With
    If prob > 100 Then prob = 100
    If prob < 0 Then prob = 0
    CalcularExitoExtraccion = (RandomNumber(1, 100) <= CInt(prob))
    Exit Function
CalcularExitoExtraccion_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.CalcularExitoExtraccion", Erl)
    CalcularExitoExtraccion = True
End Function

Public Function CalcularCantidadExtraccion(ByVal UserIndex As Integer, ByVal ProfesionId As Integer) As Integer
    On Error GoTo CalcularCantidadExtraccion_Err
    CalcularCantidadExtraccion = 1
    If ProfesionId < PROF_MIN_ID Or ProfesionId > PROF_MAX_ID Then Exit Function
    If Not ProfesionConfig(ProfesionId).EsExtraccion Then Exit Function
    Dim skillActual As Integer
    Dim rango       As Integer
    Dim resultado   As Single
    skillActual = UserList(UserIndex).Stats.UserSkills(ProfesionId)
    With ProfesionConfig(ProfesionId)
        If skillActual > .SkillParaMax Then skillActual = .SkillParaMax
        rango = .CantidadMax - .CantidadMin
        If rango < 0 Then rango = 0
        resultado = CSng(.CantidadMin) + (CSng(rango) * CSng(skillActual) / CSng(.SkillParaMax))
        ' Plan 15.002 (Ola 2, G1): el "grifo de oro" de extraer en zona segura -- sin
        ' riesgo ni penalizacion, la pesca AFK en zona segura rendia mas oro/hora que
        ' cualquier otra fuente. No toca la probabilidad de exito, solo la cantidad.
        If MapInfo(UserList(UserIndex).pos.Map).Seguro = 1 And .PenalizacionZonaSegura > 0 Then
            resultado = Int(resultado * CSng(100 - .PenalizacionZonaSegura) / 100)
        End If
    End With
    If resultado < 1 Then resultado = 1
    CalcularCantidadExtraccion = CInt(resultado)
    Exit Function
CalcularCantidadExtraccion_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.CalcularCantidadExtraccion", Erl)
End Function

Public Function CantidadFijaArtesania(ByVal ProfesionId As Integer) As Integer
    If ProfesionId < PROF_MIN_ID Or ProfesionId > PROF_MAX_ID Then
        CantidadFijaArtesania = 1
    Else
        CantidadFijaArtesania = ProfesionConfig(ProfesionId).CantidadFija
        If CantidadFijaArtesania < 1 Then CantidadFijaArtesania = 1
    End If
End Function

Public Function ProfesionDelItem(ByVal ObjIndex As Integer) As Integer
    On Error GoTo ProfesionDelItem_Err
    ProfesionDelItem = 0
    If ObjIndex <= 0 Then Exit Function
    ' Manuales (otUseOnce) y pociones de olvido (otPotions) usan ProfesionId con OTRA
    ' semantica: "ensena/olvida esta profesion", no "requiere esta profesion".
    If ObjData(ObjIndex).OBJType = e_OBJType.otUseOnce Or ObjData(ObjIndex).OBJType = e_OBJType.otPotions Then Exit Function
    ' Fuente primaria (plan 04.001): el campo ProfesionId de obj.dat. Cubre armaduras,
    ' pergaminos de receta y items de oficio que el mapeo por Subtipo no alcanzaba.
    If ObjData(ObjIndex).ProfesionId >= PROF_MIN_ID And ObjData(ObjIndex).ProfesionId <= PROF_MAX_ID Then
        ProfesionDelItem = ObjData(ObjIndex).ProfesionId
        Exit Function
    End If
    ' Fallback legacy: herramientas de trabajo por subtipo.
    If ObjData(ObjIndex).OBJType <> e_OBJType.otWorkingTools Then Exit Function
    Select Case ObjData(ObjIndex).Subtipo
        Case e_WorkingToolSubType.FishingRod, e_WorkingToolSubType.FishingNet, e_WorkingToolSubType.FishingLine
            ProfesionDelItem = e_Skill.Pescar
        Case e_WorkingToolSubType.AlchemyScissors, e_WorkingToolSubType.AlchemyCauldron
            ProfesionDelItem = e_Skill.Alquimia
        Case e_WorkingToolSubType.CarpentryHacksaw
            ProfesionDelItem = e_Skill.Carpinteria
        Case e_WorkingToolSubType.FellingAxe
            ProfesionDelItem = e_Skill.Talar
        Case e_WorkingToolSubType.SmithHammer
            ProfesionDelItem = e_Skill.Herreria
        Case e_WorkingToolSubType.MinerPickaxe
            ProfesionDelItem = e_Skill.Mineria
        Case e_WorkingToolSubType.TailorSewingbox
            ProfesionDelItem = e_Skill.Sastreria
    End Select
    Exit Function
ProfesionDelItem_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.ProfesionDelItem", Erl)
End Function

Public Sub UsarManualProfesion(ByVal UserIndex As Integer, ByVal Slot As Byte)
    On Error GoTo UsarManualProfesion_Err
    If Not IsFeatureEnabled("professions_learnable") Then
        Call WriteLocaleMsg(UserIndex, MSG_PROF_ITEM_NO_USABLE, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    Dim ObjIndex    As Integer
    Dim ProfesionId As Integer
    ObjIndex = UserList(UserIndex).invent.Object(Slot).ObjIndex
    If ObjIndex <= 0 Then Exit Sub
    ProfesionId = ObjData(ObjIndex).ProfesionId
    If ProfesionId < PROF_MIN_ID Or ProfesionId > PROF_MAX_ID Then Exit Sub
    If Not PuedeAprenderProfesion(UserIndex, ProfesionId) Then Exit Sub
    If Not AprenderProfesion(UserIndex, ProfesionId) Then Exit Sub
    Call QuitarUserInvItem(UserIndex, Slot, 1)
    Call UpdateUserInv(False, UserIndex, Slot)
    Exit Sub
UsarManualProfesion_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.UsarManualProfesion", Erl)
End Sub

Public Sub UsarPocionOlvidoProfesion(ByVal UserIndex As Integer, ByVal Slot As Byte)
    On Error GoTo UsarPocionOlvidoProfesion_Err
    If Not IsFeatureEnabled("professions_learnable") Then
        Call WriteLocaleMsg(UserIndex, MSG_PROF_ITEM_NO_USABLE, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    Dim ObjIndex    As Integer
    Dim ProfesionId As Integer
    ObjIndex = UserList(UserIndex).invent.Object(Slot).ObjIndex
    If ObjIndex <= 0 Then Exit Sub
    ProfesionId = ObjData(ObjIndex).ProfesionId
    If ProfesionId < PROF_MIN_ID Or ProfesionId > PROF_MAX_ID Then Exit Sub
    If Not PuedeOlvidarProfesion(UserIndex, ProfesionId) Then Exit Sub
    ' La pocion solo se consume si el olvido se aplico de verdad (plan 04.001).
    If Not OlvidarProfesion(UserIndex, ProfesionId) Then Exit Sub
    Call QuitarUserInvItem(UserIndex, Slot, 1)
    Call UpdateUserInv(False, UserIndex, Slot)
    Exit Sub
UsarPocionOlvidoProfesion_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.UsarPocionOlvidoProfesion", Erl)
End Sub

Private Sub RefrescarHerramientasInventario(ByVal UserIndex As Integer, ByVal ProfesionId As Integer)
    On Error GoTo RefrescarHerramientasInventario_Err
    Dim slot   As Byte
    Dim ObjIdx As Integer
    For slot = 1 To MAX_INVENTORY_SLOTS
        ObjIdx = UserList(UserIndex).invent.Object(slot).ObjIndex
        If ObjIdx > 0 Then
            If ProfesionDelItem(ObjIdx) = ProfesionId Then
                ' Si perdio el permiso y la tiene equipada, desequipar (plan 04.001):
                ' antes quedaba puesta y era imposible sacarla (CanUseObject la bloqueaba).
                If UserList(UserIndex).invent.Object(slot).Equipped Then
                    If Not TieneProfesionAprendida(UserIndex, ProfesionId) Then
                        Call Desequipar(UserIndex, slot)
                    End If
                End If
                Call WriteChangeInventorySlot(UserIndex, slot)
            End If
        End If
    Next slot
    Exit Sub
RefrescarHerramientasInventario_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.RefrescarHerramientasInventario", Erl)
End Sub

Public Sub ResetUserProfessions(ByVal UserIndex As Integer)
    ' QA (plan 04.001): /RESET limpia slots, contador de olvidos y persistencia.
    On Error GoTo ResetUserProfessions_Err
    Dim i      As Byte
    Dim ObjIdx As Integer
    For i = 1 To PROF_MAX_SLOTS
        UserList(UserIndex).Professions(i) = 0
    Next i
    UserList(UserIndex).ProfessionForgotCount = 0
    Call Execute("DELETE FROM user_professions WHERE user_id = ?;", UserList(UserIndex).Id)
    Call Execute("UPDATE user SET profession_forgot_count = 0 WHERE id = ?;", UserList(UserIndex).Id)
    ' Desequipar y refrescar las herramientas que quedaron sin permiso.
    For i = 1 To MAX_INVENTORY_SLOTS
        ObjIdx = UserList(UserIndex).invent.Object(i).ObjIndex
        If ObjIdx > 0 Then
            If ProfesionDelItem(ObjIdx) > 0 Then
                If UserList(UserIndex).invent.Object(i).Equipped Then
                    If Not TieneProfesionAprendida(UserIndex, ProfesionDelItem(ObjIdx)) Then
                        Call Desequipar(UserIndex, i)
                    End If
                End If
                Call WriteChangeInventorySlot(UserIndex, i)
            End If
        End If
    Next i
    Call WriteProfessionsUpdate(UserIndex)
    Exit Sub
ResetUserProfessions_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.ResetUserProfessions", Erl)
End Sub

Public Function HandleProfesionChatCommand(ByVal UserIndex As Integer, ByVal chat As String) As Boolean
    On Error GoTo HandleProfesionChatCommand_Err
    HandleProfesionChatCommand = False
    Dim body  As String
    Dim parts() As String
    Dim cmd   As String
    Dim targetName As String
    Dim profArg    As String
    Dim profId     As Integer
    Dim tUserRef   As t_UserReference
    body = mid$(chat, 2)
    body = Trim$(body)
    If LenB(body) = 0 Then Exit Function
    parts = Split(body, " ")
    If UBound(parts) < 2 Then Exit Function
    cmd = UCase$(parts(0))
    If cmd <> "APRENDE" And cmd <> "OLVIDA" Then Exit Function
    HandleProfesionChatCommand = True
    targetName = Replace(parts(1), "+", " ")
    profArg = parts(2)
    profId = ResolverProfesionId(profArg)
    If profId < PROF_MIN_ID Or profId > PROF_MAX_ID Then
        Call WriteConsoleMsg(UserIndex, "Profesion invalida: " & profArg & ". Usa nombre (Carpinteria, Tala, ...) o id (17-23).", e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UCase$(targetName) = "YO" Then
        Call SetUserRef(tUserRef, UserIndex)
    Else
        tUserRef = NameIndex(targetName)
    End If
    If Not IsValidUserRef(tUserRef) Then
        Call WriteConsoleMsg(UserIndex, "Usuario offline: " & targetName, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    ' Plan 04.001: operar sobre OTRO personaje exige Dios o Admin; sobre staff superior, solo Admin.
    If tUserRef.ArrayIndex <> UserIndex Then
        If (UserList(UserIndex).flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin)) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Solo Dios o Admin pueden modificar profesiones de otro personaje.", e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        If (UserList(tUserRef.ArrayIndex).flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin)) <> 0 And (UserList(UserIndex).flags.Privilegios And e_PlayerType.Admin) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Solo un Admin puede modificar profesiones de un Dios o Admin.", e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    End If
    If cmd = "APRENDE" Then
        If AprenderProfesion(tUserRef.ArrayIndex, profId, True) Then
            Call WriteConsoleMsg(UserIndex, "Aprendio " & NombreProfesion(profId) & " para " & targetName, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No se pudo aprender " & NombreProfesion(profId) & " para " & targetName & " (ya aprendida, slots llenos o limite de olvidos).", e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        If OlvidarProfesion(tUserRef.ArrayIndex, profId, True) Then
            Call WriteConsoleMsg(UserIndex, "Olvido " & NombreProfesion(profId) & " para " & targetName, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No se pudo olvidar " & NombreProfesion(profId) & " para " & targetName & " (no la tiene o limite de olvidos).", e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End If
    Call LogGM(GetUserRealName(UserIndex), "/" & cmd & " " & targetName & " " & NombreProfesion(profId))
    Exit Function
HandleProfesionChatCommand_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.HandleProfesionChatCommand", Erl)
End Function

Private Function ResolverProfesionId(ByVal arg As String) As Integer
    Dim n As Integer
    n = CInt(val(arg))
    If n >= PROF_MIN_ID And n <= PROF_MAX_ID Then
        ResolverProfesionId = n
        Exit Function
    End If
    Select Case UCase$(arg)
        Case "TALA", "TALAR"
            ResolverProfesionId = e_Skill.Talar
        Case "PESCA", "PESCAR"
            ResolverProfesionId = e_Skill.Pescar
        Case "MINERIA"
            ResolverProfesionId = e_Skill.Mineria
        Case "HERRERIA"
            ResolverProfesionId = e_Skill.Herreria
        Case "CARPINTERIA"
            ResolverProfesionId = e_Skill.Carpinteria
        Case "ALQUIMIA"
            ResolverProfesionId = e_Skill.Alquimia
        Case "SASTRERIA"
            ResolverProfesionId = e_Skill.Sastreria
        Case Else
            ResolverProfesionId = 0
    End Select
End Function
