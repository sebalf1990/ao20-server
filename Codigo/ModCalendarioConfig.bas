Attribute VB_Name = "ModCalendarioConfig"
Option Explicit

' Argentum 20 Game Server - automatic events calendar, configuration side.
'
' Stage 2 of plan 2026-11-24-revision-eventos-automaticos. This module has NO game state: types, the
' parser/validator of Dat\CalendarioEventos.dat and the PURE functions that decide when a turno fires
' (they receive the clock as a parameter, so they are unit tested without waiting for anything).
' The scheduler itself lives in ModCalendarioEventos.

Public Const CAL_MAX_TURNOS As Long = 48
Public Const CAL_VENTANA_MIN_DEFAULT As Long = 30
Public Const CAL_REINTENTO_MIN_DEFAULT As Long = 10
Public Const CAL_AVISO_PREVIO_DEFAULT As Long = 5
Public Const CAL_MAX_DESCRIPCION As Long = 40
Public Const CAL_MAX_CUOTA As Long = 10000000

' Day bits: Weekday(date, vbSunday) - 1, so Sunday = bit 0.
Public Const CAL_DIA_DOM As Long = 1
Public Const CAL_DIA_LUN As Long = 2
Public Const CAL_DIA_MAR As Long = 4
Public Const CAL_DIA_MIE As Long = 8
Public Const CAL_DIA_JUE As Long = 16
Public Const CAL_DIA_VIE As Long = 32
Public Const CAL_DIA_SAB As Long = 64
Public Const CAL_DIAS_TODOS As Long = 127

Public Enum e_CalModo
    CalModoNinguno = 0
    CalModoDeathMatch = 1
    CalModoCaceria = 2
    CalModoCaptura = 3
    CalModoNaval = 4
    CalModoInvasion = 5
End Enum

Public Type t_CalConfig
    Version As Long
    VentanaMin As Long
    ReintentoMin As Long
    AvisoPrevioMin As Long
    BloquearConInvasion As Boolean
End Type

' Raw values of a variant as read from the file. They are Long on purpose: the range checks run BEFORE
' anything is cast to the Byte fields of t_NewScenearioSettings (a Byte cast of 300 raises error 6).
Public Type t_CalVarianteRaw
    Modo As Long
    MinLevel As Long
    MaxLevel As Long
    MinPlayers As Long
    MaxPlayers As Long
    TeamSize As Long
    TeamType As Long
    RoundNumber As Long
    InscriptionFee As Long
    MinOnline As Long
    Description As String
End Type

Public Type t_CalVariante
    Nombre As String
    Valida As Boolean
    Modo As Long
    Settings As t_NewScenearioSettings
    MinOnline As Long
End Type

Public Type t_CalTurno
    Clave As String
    DiasMask As Long
    MinutoDelDia As Long
    VentanaMin As Long
    VarIdx As Long
    Texto As String
End Type

Public Type t_CalResumen
    ArchivoExiste As Boolean
    TurnosCargados As Long
    TurnosDescartados As Long
    VariantesValidas As Long
    VariantesDescartadas As Long
    Advertencias As Long
End Type

' ------------------------------------------------------------------------------------------------
' Pure helpers
' ------------------------------------------------------------------------------------------------

' True for an optional "-" followed by 1 to 9 digits (so CLng can never overflow).
Public Function CalEsEnteroTexto(ByVal texto As String) As Boolean
    Dim i      As Long
    Dim inicio As Long
    texto = Trim$(texto)
    If LenB(texto) = 0 Then Exit Function
    inicio = 1
    If Left$(texto, 1) = "-" Then inicio = 2
    If Len(texto) < inicio Then Exit Function
    If Len(texto) - inicio + 1 > 9 Then Exit Function
    For i = inicio To Len(texto)
        If Mid$(texto, i, 1) < "0" Or Mid$(texto, i, 1) > "9" Then Exit Function
    Next i
    CalEsEnteroTexto = True
End Function

Private Function CalBitDia(ByVal abreviatura As String) As Long
    Select Case abreviatura
        Case "DOM"
            CalBitDia = CAL_DIA_DOM
        Case "LUN"
            CalBitDia = CAL_DIA_LUN
        Case "MAR"
            CalBitDia = CAL_DIA_MAR
        Case "MIE"
            CalBitDia = CAL_DIA_MIE
        Case "JUE"
            CalBitDia = CAL_DIA_JUE
        Case "VIE"
            CalBitDia = CAL_DIA_VIE
        Case "SAB"
            CalBitDia = CAL_DIA_SAB
        Case Else
            CalBitDia = 0
    End Select
End Function

' "TODOS" or a comma separated list of DOM,LUN,MAR,MIE,JUE,VIE,SAB. Returns the day mask.
Public Function CalParseDias(ByVal texto As String, ByRef mascara As Long) As Boolean
    Dim partes()  As String
    Dim i         As Long
    Dim bit       As Long
    Dim resultado As Long
    mascara = 0
    texto = UCase$(Trim$(texto))
    If LenB(texto) = 0 Then Exit Function
    If texto = "TODOS" Then
        mascara = CAL_DIAS_TODOS
        CalParseDias = True
        Exit Function
    End If
    partes = Split(texto, ",")
    For i = 0 To UBound(partes)
        bit = CalBitDia(Trim$(partes(i)))
        If bit = 0 Then Exit Function
        resultado = resultado Or bit
    Next i
    mascara = resultado
    CalParseDias = True
End Function

' "HH:MM" (HH 0-23 with one or two digits, MM exactly two digits 00-59) -> minute of the day.
Public Function CalParseHora(ByVal texto As String, ByRef minutoDelDia As Long) As Boolean
    Dim partes() As String
    Dim hh       As Long
    Dim mm       As Long
    texto = Trim$(texto)
    partes = Split(texto, ":")
    If UBound(partes) <> 1 Then Exit Function
    partes(0) = Trim$(partes(0))
    partes(1) = Trim$(partes(1))
    If Len(partes(0)) < 1 Or Len(partes(0)) > 2 Or Len(partes(1)) <> 2 Then Exit Function
    If Not CalEsEnteroTexto(partes(0)) Or Not CalEsEnteroTexto(partes(1)) Then Exit Function
    If Left$(partes(0), 1) = "-" Or Left$(partes(1), 1) = "-" Then Exit Function
    hh = CLng(partes(0))
    mm = CLng(partes(1))
    If hh > 23 Or mm > 59 Then Exit Function
    minutoDelDia = hh * 60 + mm
    CalParseHora = True
End Function

Public Function CalParseModo(ByVal texto As String) As Long
    Select Case UCase$(Trim$(texto))
        Case "DEATHMATCH"
            CalParseModo = e_CalModo.CalModoDeathMatch
        Case "CACERIA"
            CalParseModo = e_CalModo.CalModoCaceria
        Case "CAPTURA"
            CalParseModo = e_CalModo.CalModoCaptura
        Case "NAVAL"
            CalParseModo = e_CalModo.CalModoNaval
        Case "INVASION"
            CalParseModo = e_CalModo.CalModoInvasion
        Case Else
            CalParseModo = e_CalModo.CalModoNinguno
    End Select
End Function

Public Function CalNombreModo(ByVal modo As Long) As String
    Select Case modo
        Case e_CalModo.CalModoDeathMatch
            CalNombreModo = "DEATHMATCH"
        Case e_CalModo.CalModoCaceria
            CalNombreModo = "CACERIA"
        Case e_CalModo.CalModoCaptura
            CalNombreModo = "CAPTURA"
        Case e_CalModo.CalModoNaval
            CalNombreModo = "NAVAL"
        Case e_CalModo.CalModoInvasion
            CalNombreModo = "INVASION"
        Case Else
            CalNombreModo = "?"
    End Select
End Function

' e_CalModo -> e_EventType of ModLobby.
Public Function CalModoAEventType(ByVal modo As Long) As Long
    Select Case modo
        Case e_CalModo.CalModoDeathMatch
            CalModoAEventType = e_EventType.DeathMatch
        Case e_CalModo.CalModoCaceria
            CalModoAEventType = e_EventType.NpcHunt
        Case e_CalModo.CalModoCaptura
            CalModoAEventType = e_EventType.CaptureTheFlag
        Case e_CalModo.CalModoNaval
            CalModoAEventType = e_EventType.NavalBattle
        Case Else
            CalModoAEventType = -1
    End Select
End Function

Public Function CalHoraTexto(ByVal minutoDelDia As Long) As String
    CalHoraTexto = Format$(minutoDelDia \ 60, "00") & ":" & Format$(minutoDelDia Mod 60, "00")
End Function

Public Function CalDiasTexto(ByVal mascara As Long) As String
    Dim nombres As Variant
    Dim i       As Long
    Dim texto   As String
    If mascara = CAL_DIAS_TODOS Then
        CalDiasTexto = "TODOS"
        Exit Function
    End If
    nombres = Array("DOM", "LUN", "MAR", "MIE", "JUE", "VIE", "SAB")
    For i = 0 To 6
        If (mascara And CLng(2 ^ i)) <> 0 Then
            If LenB(texto) <> 0 Then texto = texto & ","
            texto = texto & nombres(i)
        End If
    Next i
    CalDiasTexto = texto
End Function

' Stable key of a turno, derived from the day MASK and the minute (not from the text of the line): TODOS and
' the list of the 7 days are the same turno. Changing days or time of a turno makes a NEW key on purpose.
Public Function CalClaveTurno(ByVal mascara As Long, ByVal minutoDelDia As Long) As String
    CalClaveTurno = "D" & mascara & "-" & Format$(minutoDelDia \ 60, "00") & Format$(minutoDelDia Mod 60, "00")
End Function

Public Function CalBitDelDia(ByVal dia As Date) As Long
    CalBitDelDia = CLng(2 ^ (Weekday(dia, vbSunday) - 1))
End Function

' Whole minutes since 2000-01-01 00:00. Minute arithmetic on Long avoids comparing Doubles.
Public Function CalMinutosDesdeEpoca(ByVal momento As Date) As Long
    CalMinutosDesdeEpoca = CLng(DateDiff("n", DateSerial(2000, 1, 1), momento))
End Function

' True when `ahora` is inside [start, start + ventanaMin) of an occurrence of the turno that started TODAY or
' YESTERDAY (a window that crosses midnight belongs to the day it started). fechaInicio = "yyyy-mm-dd" of
' that occurrence (the day of the start, never the day of `ahora`).
Public Function CalTurnoEnVentana(ByVal diasMask As Long, ByVal minutoDelDia As Long, ByVal ventanaMin As Long, ByVal ahora As Date, ByRef fechaInicio As String) As Boolean
    Dim offsetDias As Long
    Dim dia        As Date
    Dim inicioMin  As Long
    Dim ahoraMin   As Long
    ahoraMin = CalMinutosDesdeEpoca(ahora)
    For offsetDias = 0 To 1
        dia = DateAdd("d", -offsetDias, DateValue(ahora))
        If (diasMask And CalBitDelDia(dia)) <> 0 Then
            inicioMin = CalMinutosDesdeEpoca(dia) + minutoDelDia
            If ahoraMin >= inicioMin And ahoraMin < inicioMin + ventanaMin Then
                fechaInicio = Format$(dia, "yyyy-mm-dd")
                CalTurnoEnVentana = True
                Exit Function
            End If
        End If
    Next offsetDias
End Function

' THE firing decision: inside the window of an occurrence that was not handled yet. `ultimaFechaProcesada` is
' the "yyyy-mm-dd" of the last occurrence of this turno that already has a row (opened, cancelled, skipped...):
' one fire per turno per date.
Public Function TurnoDebeDisparar(ByVal diasMask As Long, ByVal minutoDelDia As Long, ByVal ventanaMin As Long, ByVal ahora As Date, ByVal ultimaFechaProcesada As String, ByRef fechaInicio As String) As Boolean
    If Not CalTurnoEnVentana(diasMask, minutoDelDia, ventanaMin, ahora, fechaInicio) Then Exit Function
    If fechaInicio = ultimaFechaProcesada Then Exit Function
    TurnoDebeDisparar = True
End Function

' Next start of the turno that is still ahead of `ahora` (within 8 days), or 0 if the mask is empty.
Public Function CalProximoInicio(ByVal diasMask As Long, ByVal minutoDelDia As Long, ByVal ahora As Date) As Date
    Dim offsetDias As Long
    Dim dia        As Date
    For offsetDias = 0 To 7
        dia = DateAdd("d", offsetDias, DateValue(ahora))
        If (diasMask And CalBitDelDia(dia)) <> 0 Then
            If CalMinutosDesdeEpoca(dia) + minutoDelDia > CalMinutosDesdeEpoca(ahora) Then
                CalProximoInicio = DateAdd("n", minutoDelDia, dia)
                Exit Function
            End If
        End If
    Next offsetDias
End Function

' Validation of a variant. Returns False and the reason, or True (possibly with a warning).
Public Function CalValidarVariante(ByRef v As t_CalVarianteRaw, ByRef motivo As String, ByRef advertencia As String) As Boolean
    motivo = vbNullString
    advertencia = vbNullString
    Select Case v.Modo
        Case e_CalModo.CalModoDeathMatch, e_CalModo.CalModoCaceria, e_CalModo.CalModoCaptura, e_CalModo.CalModoNaval
            ' supported
        Case e_CalModo.CalModoInvasion
            motivo = "Modo=INVASION is not enabled in this build"
            Exit Function
        Case Else
            motivo = "unknown Modo"
            Exit Function
    End Select
    If v.MinLevel < 1 Or v.MinLevel > v.MaxLevel Or v.MaxLevel > 47 Then
        motivo = "levels must satisfy 1 <= MinLevel <= MaxLevel <= 47 (got " & v.MinLevel & "-" & v.MaxLevel & ")"
        Exit Function
    End If
    If v.MinPlayers < 2 Or v.MinPlayers > v.MaxPlayers Or v.MaxPlayers > 255 Then
        motivo = "players must satisfy 2 <= MinPlayers <= MaxPlayers <= 255 (got " & v.MinPlayers & "-" & v.MaxPlayers & ")"
        Exit Function
    End If
    If v.Modo = e_CalModo.CalModoCaptura And v.MaxPlayers > 48 Then
        motivo = "CAPTURA allows at most 48 players (got MaxPlayers=" & v.MaxPlayers & ")"
        Exit Function
    End If
    If v.Modo = e_CalModo.CalModoNaval Then
        If (v.MaxPlayers Mod 2) <> 0 Then advertencia = "NAVAL with an odd MaxPlayers (" & v.MaxPlayers & ")"
    End If
    If v.TeamType < 0 Or v.TeamType > 1 Then
        motivo = "TeamType must be 0 or 1 (got " & v.TeamType & ")"
        Exit Function
    End If
    If v.TeamSize < 0 Or v.TeamSize > 255 Then
        motivo = "TeamSize must be 0-255 (got " & v.TeamSize & ")"
        Exit Function
    End If
    If v.RoundNumber < 1 Or v.RoundNumber > 9 Then
        motivo = "RoundNumber must be 1-9 (got " & v.RoundNumber & ")"
        Exit Function
    End If
    If v.InscriptionFee < 0 Or v.InscriptionFee > CAL_MAX_CUOTA Then
        motivo = "InscriptionFee must be 0-" & CAL_MAX_CUOTA & " (got " & v.InscriptionFee & ")"
        Exit Function
    End If
    If v.MinOnline < 0 Or v.MinOnline > 255 Then
        motivo = "MinOnline must be 0-255 (got " & v.MinOnline & ")"
        Exit Function
    End If
    If Len(v.Description) > CAL_MAX_DESCRIPCION Then
        motivo = "Description is longer than " & CAL_MAX_DESCRIPCION & " characters"
        Exit Function
    End If
    If InStr(1, v.Description, Chr$(172)) > 0 Then
        motivo = "Description cannot contain the separator character"
        Exit Function
    End If
    CalValidarVariante = True
End Function

' Only call it after CalValidarVariante returned True (the casts below cannot overflow then).
Public Sub CalConstruirSettings(ByRef v As t_CalVarianteRaw, ByRef s As t_NewScenearioSettings)
    s.MinLevel = CByte(v.MinLevel)
    s.MaxLevel = CByte(v.MaxLevel)
    s.MinPlayers = CByte(v.MinPlayers)
    s.MaxPlayers = CByte(v.MaxPlayers)
    s.TeamSize = CByte(v.TeamSize)
    s.InscriptionFee = v.InscriptionFee
    s.ScenearioType = CByte(CalModoAEventType(v.Modo))
    s.TeamType = CByte(v.TeamType)
    s.RoundNumber = CByte(v.RoundNumber)
    s.Description = v.Description
    s.Password = vbNullString
End Sub

' ------------------------------------------------------------------------------------------------
' File loader
' ------------------------------------------------------------------------------------------------

' Loads Archivo into the arrays that the caller gives (temporaries: the live state is not touched here).
' Returns False if the file does not exist (one log line, nothing loaded). Every discard leaves ONE LogError
' line, here, at load time (never again on each tick).
Public Function CalendarioConfig_Cargar(ByVal archivo As String, ByRef cfg As t_CalConfig, ByRef turnos() As t_CalTurno, ByRef numTurnos As Long, ByRef variantes() As t_CalVariante, ByRef numVariantes As Long, ByRef resumen As t_CalResumen) As Boolean
    On Error GoTo CalendarioConfig_Cargar_Err
    Dim ini   As clsIniManager
    Dim n     As Long
    Dim linea As String
    Dim vacio As t_CalResumen
    numTurnos = 0
    numVariantes = 0
    ReDim turnos(1 To CAL_MAX_TURNOS) As t_CalTurno
    ReDim variantes(1 To CAL_MAX_TURNOS) As t_CalVariante
    resumen = vacio
    cfg.Version = 1
    cfg.VentanaMin = CAL_VENTANA_MIN_DEFAULT
    cfg.ReintentoMin = CAL_REINTENTO_MIN_DEFAULT
    cfg.AvisoPrevioMin = CAL_AVISO_PREVIO_DEFAULT
    cfg.BloquearConInvasion = True
    If Not FileExist(archivo) Then
        Call LogError("[Calendario] file not found: " & archivo & ". The calendar stays empty.")
        Exit Function
    End If
    resumen.ArchivoExiste = True
    Set ini = New clsIniManager
    Call ini.Initialize(archivo)
    Call CalLeerSeccionCalendario(ini, cfg)
    For n = 1 To CAL_MAX_TURNOS
        linea = Trim$(ini.GetValue("TURNOS", "Turno" & n))
        If LenB(linea) <> 0 Then
            Call CalProcesarTurno(ini, n, linea, cfg, turnos, numTurnos, variantes, numVariantes, resumen)
        End If
    Next n
    Call CalAdvertirSolapes(turnos, numTurnos, resumen)
    Set ini = Nothing
    CalendarioConfig_Cargar = True
    Exit Function
CalendarioConfig_Cargar_Err:
    Call TraceError(Err.Number, Err.Description, "ModCalendarioConfig.CalendarioConfig_Cargar", Erl)
End Function

Private Function CalLeerEnRango(ByRef ini As clsIniManager, ByVal clave As String, ByVal valorDefault As Long, ByVal minimo As Long, ByVal maximo As Long) As Long
    Dim texto As String
    Dim valor As Long
    CalLeerEnRango = valorDefault
    texto = Trim$(ini.GetValue("CALENDARIO", clave))
    If LenB(texto) = 0 Then Exit Function
    If Not CalEsEnteroTexto(texto) Then
        Call LogError("[Calendario] [CALENDARIO] " & clave & " is not a number (" & texto & "): using " & valorDefault & ".")
        Exit Function
    End If
    valor = CLng(texto)
    If valor < minimo Or valor > maximo Then
        Call LogError("[Calendario] [CALENDARIO] " & clave & " out of range " & minimo & "-" & maximo & " (" & valor & "): using " & valorDefault & ".")
        Exit Function
    End If
    CalLeerEnRango = valor
End Function

Private Sub CalLeerSeccionCalendario(ByRef ini As clsIniManager, ByRef cfg As t_CalConfig)
    cfg.Version = CalLeerEnRango(ini, "Version", 1, 1, 1)
    cfg.VentanaMin = CalLeerEnRango(ini, "VentanaMin", CAL_VENTANA_MIN_DEFAULT, 5, 180)
    cfg.ReintentoMin = CalLeerEnRango(ini, "ReintentoMin", CAL_REINTENTO_MIN_DEFAULT, 5, 60)
    ' AvisoPrevioMin is parsed and validated but no notice is sent yet (planned for after the beta opening).
    cfg.AvisoPrevioMin = CalLeerEnRango(ini, "AvisoPrevioMin", CAL_AVISO_PREVIO_DEFAULT, 0, 30)
    cfg.BloquearConInvasion = (CalLeerEnRango(ini, "BloquearConInvasion", 1, 0, 1) <> 0)
End Sub

Private Sub CalDescartarTurno(ByRef resumen As t_CalResumen, ByVal n As Long, ByVal motivo As String)
    resumen.TurnosDescartados = resumen.TurnosDescartados + 1
    Call LogError("[Calendario] Turno" & n & " discarded: " & motivo)
End Sub

Private Sub CalProcesarTurno(ByRef ini As clsIniManager, ByVal n As Long, ByVal linea As String, ByRef cfg As t_CalConfig, ByRef turnos() As t_CalTurno, ByRef numTurnos As Long, ByRef variantes() As t_CalVariante, ByRef numVariantes As Long, ByRef resumen As t_CalResumen)
    Dim partes()  As String
    Dim mascara   As Long
    Dim minuto    As Long
    Dim ventana   As Long
    Dim varIdx    As Long
    Dim nombreVar As String
    Dim clave     As String
    Dim i         As Long
    partes = Split(linea, "|")
    If UBound(partes) < 2 Or UBound(partes) > 3 Then
        Call CalDescartarTurno(resumen, n, "expected Dias|HH:MM|Variante[|VentanaMin]: " & linea)
        Exit Sub
    End If
    If Not CalParseDias(partes(0), mascara) Then
        Call CalDescartarTurno(resumen, n, "invalid Dias (" & Trim$(partes(0)) & ")")
        Exit Sub
    End If
    If Not CalParseHora(partes(1), minuto) Then
        Call CalDescartarTurno(resumen, n, "invalid HH:MM (" & Trim$(partes(1)) & ")")
        Exit Sub
    End If
    ventana = cfg.VentanaMin
    If UBound(partes) = 3 Then
        If Not CalEsEnteroTexto(partes(3)) Then
            Call CalDescartarTurno(resumen, n, "invalid VentanaMin (" & Trim$(partes(3)) & ")")
            Exit Sub
        End If
        ventana = CLng(Trim$(partes(3)))
        If ventana < 5 Or ventana > 180 Then
            Call CalDescartarTurno(resumen, n, "VentanaMin out of range 5-180 (" & ventana & ")")
            Exit Sub
        End If
    End If
    nombreVar = UCase$(Trim$(partes(2)))
    If LenB(nombreVar) = 0 Then
        Call CalDescartarTurno(resumen, n, "missing Variante")
        Exit Sub
    End If
    varIdx = CalObtenerVariante(ini, nombreVar, variantes, numVariantes, resumen)
    If varIdx = 0 Then
        Call CalDescartarTurno(resumen, n, "variant " & nombreVar & " is not valid (see the previous line)")
        Exit Sub
    End If
    clave = CalClaveTurno(mascara, minuto)
    For i = 1 To numTurnos
        If (turnos(i).DiasMask And mascara) <> 0 And turnos(i).MinutoDelDia = minuto Then
            Call CalDescartarTurno(resumen, n, "same day and minute as " & turnos(i).Clave & " (" & turnos(i).Texto & ")")
            Exit Sub
        End If
    Next i
    numTurnos = numTurnos + 1
    turnos(numTurnos).Clave = clave
    turnos(numTurnos).DiasMask = mascara
    turnos(numTurnos).MinutoDelDia = minuto
    turnos(numTurnos).VentanaMin = ventana
    turnos(numTurnos).VarIdx = varIdx
    turnos(numTurnos).Texto = "Turno" & n & " " & CalDiasTexto(mascara) & " " & CalHoraTexto(minuto) & " " & nombreVar
    resumen.TurnosCargados = resumen.TurnosCargados + 1
End Sub

' Index of the variant in the cache, loading it the first time (0 = invalid; the reason was logged ONCE).
Private Function CalObtenerVariante(ByRef ini As clsIniManager, ByVal nombre As String, ByRef variantes() As t_CalVariante, ByRef numVariantes As Long, ByRef resumen As t_CalResumen) As Long
    Dim i As Long
    For i = 1 To numVariantes
        If variantes(i).Nombre = nombre Then
            If variantes(i).Valida Then CalObtenerVariante = i
            Exit Function
        End If
    Next i
    numVariantes = numVariantes + 1
    variantes(numVariantes).Nombre = nombre
    variantes(numVariantes).Valida = CalCargarVariante(ini, nombre, variantes(numVariantes), resumen)
    If variantes(numVariantes).Valida Then
        resumen.VariantesValidas = resumen.VariantesValidas + 1
        CalObtenerVariante = numVariantes
    Else
        resumen.VariantesDescartadas = resumen.VariantesDescartadas + 1
    End If
End Function

Private Function CalLeerEntero(ByRef ini As clsIniManager, ByVal seccion As String, ByVal clave As String, ByVal valorDefault As Long, ByRef valor As Long) As Boolean
    Dim texto As String
    texto = Trim$(ini.GetValue(seccion, clave))
    If LenB(texto) = 0 Then
        valor = valorDefault
        CalLeerEntero = True
        Exit Function
    End If
    If Not CalEsEnteroTexto(texto) Then Exit Function
    valor = CLng(texto)
    CalLeerEntero = True
End Function

Private Function CalCargarVariante(ByRef ini As clsIniManager, ByVal nombre As String, ByRef v As t_CalVariante, ByRef resumen As t_CalResumen) As Boolean
    Dim seccion       As String
    Dim raw           As t_CalVarianteRaw
    Dim motivo        As String
    Dim advertencia   As String
    Dim textoModo     As String
    Dim textoOnline   As String
    Dim rondaDefault  As Long
    Dim ok            As Boolean
    seccion = "VARIANTE_" & nombre
    textoModo = Trim$(ini.GetValue(seccion, "Modo"))
    If LenB(textoModo) = 0 Then
        Call LogError("[Calendario] variant " & nombre & " discarded: section [" & seccion & "] is missing or has no Modo.")
        Exit Function
    End If
    raw.Modo = CalParseModo(textoModo)
    If raw.Modo = e_CalModo.CalModoNinguno Then
        Call LogError("[Calendario] variant " & nombre & " discarded: unknown Modo (" & textoModo & ").")
        Exit Function
    End If
    rondaDefault = 1
    If raw.Modo = e_CalModo.CalModoCaptura Then rondaDefault = 3
    ok = CalLeerEntero(ini, seccion, "MinLevel", 1, raw.MinLevel)
    ok = ok And CalLeerEntero(ini, seccion, "MaxLevel", 47, raw.MaxLevel)
    ok = ok And CalLeerEntero(ini, seccion, "MinPlayers", 2, raw.MinPlayers)
    ok = ok And CalLeerEntero(ini, seccion, "MaxPlayers", 20, raw.MaxPlayers)
    ok = ok And CalLeerEntero(ini, seccion, "TeamSize", 0, raw.TeamSize)
    ok = ok And CalLeerEntero(ini, seccion, "TeamType", 1, raw.TeamType)
    ok = ok And CalLeerEntero(ini, seccion, "RoundNumber", rondaDefault, raw.RoundNumber)
    ok = ok And CalLeerEntero(ini, seccion, "InscriptionFee", 0, raw.InscriptionFee)
    If Not ok Then
        Call LogError("[Calendario] variant " & nombre & " discarded: a numeric key is not a whole number of at most 9 digits.")
        Exit Function
    End If
    ' MinOnline defaults to MinPlayers.
    textoOnline = Trim$(ini.GetValue(seccion, "MinOnline"))
    If LenB(textoOnline) = 0 Then
        raw.MinOnline = raw.MinPlayers
    ElseIf CalEsEnteroTexto(textoOnline) Then
        raw.MinOnline = CLng(textoOnline)
    Else
        Call LogError("[Calendario] variant " & nombre & " discarded: MinOnline is not a whole number.")
        Exit Function
    End If
    raw.Description = Trim$(ini.GetValue(seccion, "Description"))
    If Not CalValidarVariante(raw, motivo, advertencia) Then
        Call LogError("[Calendario] variant " & nombre & " discarded: " & motivo & ".")
        Exit Function
    End If
    If LenB(advertencia) <> 0 Then
        resumen.Advertencias = resumen.Advertencias + 1
        Call LogInfoServidor("[Calendario] warning in variant " & nombre & ": " & advertencia & ".")
    End If
    Call CalConstruirSettings(raw, v.Settings)
    v.Modo = raw.Modo
    v.MinOnline = raw.MinOnline
    CalCargarVariante = True
End Function

' Overlapping windows are legal (the second turno just waits, retrying, until the first event ends) but
' they are usually a typo: one warning at load.
Private Sub CalAdvertirSolapes(ByRef turnos() As t_CalTurno, ByVal numTurnos As Long, ByRef resumen As t_CalResumen)
    Dim i As Long
    Dim j As Long
    For i = 1 To numTurnos - 1
        For j = i + 1 To numTurnos
            If (turnos(i).DiasMask And turnos(j).DiasMask) <> 0 Then
                If turnos(i).MinutoDelDia < turnos(j).MinutoDelDia + turnos(j).VentanaMin And turnos(j).MinutoDelDia < turnos(i).MinutoDelDia + turnos(i).VentanaMin Then
                    resumen.Advertencias = resumen.Advertencias + 1
                    Call LogInfoServidor("[Calendario] warning: the windows of " & turnos(i).Texto & " and " & turnos(j).Texto & " overlap.")
                End If
            End If
        Next j
    Next i
End Sub
