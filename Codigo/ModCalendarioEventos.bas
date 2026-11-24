Attribute VB_Name = "ModCalendarioEventos"
Option Explicit

' Argentum 20 Game Server - automatic events calendar, scheduler side.
'
' Stage 2 of plan 2026-11-24-revision-eventos-automaticos. Replaces the old one-hour-per-day timer that lived in frmMain
' (one fixed hour, one event type per weekday). The calendar is a list of TURNOS (days | HH:MM | variant)
' read from Dat\CalendarioEventos.dat; the pure rules (window, midnight, day mask, one fire per date) live in
' ModCalendarioConfig and are unit tested with an injected clock.
'
' Rules of the scheduler:
'  - CalendarioEventos_Tick runs once a minute from frmMain.Minuto_Timer, only with automatic_events=1.
'  - A turno opens at the first tick inside [HH:MM, HH:MM + window), never later. A blocked turno retries every
'    ReintentoMin minutes inside the window and ends OMITIDO (one staff notice) when the window closes.
'  - One row per (turno, date) in event_calendar_run (SQLite, written with the synchronous Query). Any row means
'    "already handled today": it survives a restart.
'  - The opened event is followed by a hook in ModLobby.ReleaseLobby (CalendarioEventos_LobbyLiberado), not by
'    polling: the lobby pool is LIFO and reuses an index right away, polling could confuse two events.
'  - Invasions are NOT scheduled here (Modo=INVASION is rejected by the loader).
'  - An error inside the tick is trapped here and never reaches Minuto_Timer.

Private Const ARCHIVO_CALENDARIO As String = "CalendarioEventos.dat"

Private m_Cargado          As Boolean
Private m_Cfg              As t_CalConfig
Private m_Turnos()         As t_CalTurno
Private m_NumTurnos        As Long
Private m_Variantes()      As t_CalVariante
Private m_NumVariantes     As Long
Private m_Resumen          As t_CalResumen

' Runtime state per turno (same index as m_Turnos). Rebuilt on every (re)load: the table is the truth.
Private m_FechaProcesada() As String   ' "yyyy-mm-dd" of the last occurrence that has a row or was opened
Private m_PendienteFecha() As String   ' occurrence that was blocked at least once and is still waiting
Private m_Intentos()       As Long
Private m_UltimoIntentoMin() As Long   ' CalMinutosDesdeEpoca of the last blocked attempt, -1 = none
Private m_UltimoMotivo()   As String

' The event this module opened and is still following.
Private m_SegActivo        As Boolean
Private m_SegLobby         As Integer
Private m_SegClave         As String
Private m_SegFecha         As String
Private m_SegVariante      As String
Private m_SegIntentos      As Long
Private m_SegPico          As Long

Private m_AvisoSinTabla    As Boolean

Private Function CalArchivo() As String
    CalArchivo = DatPath & ARCHIVO_CALENDARIO
End Function

Private Function CalAhora() As Date
    CalAhora = Now
End Function

Private Function CalMarcaTiempo() As String
    CalMarcaTiempo = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Function

Private Sub CalAvisarStaff(ByVal texto As String)
    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(texto, e_TextChannel.TEXTCHANNEL_SERVER_STAFF, e_FontTypeNames.FONTTYPE_INFO))
End Sub

Private Sub CalReiniciarEstado(ByVal cantidad As Long)
    ReDim m_FechaProcesada(0 To CAL_MAX_TURNOS) As String
    ReDim m_PendienteFecha(0 To CAL_MAX_TURNOS) As String
    ReDim m_Intentos(0 To CAL_MAX_TURNOS) As Long
    ReDim m_UltimoIntentoMin(0 To CAL_MAX_TURNOS) As Long
    ReDim m_UltimoMotivo(0 To CAL_MAX_TURNOS) As String
    Dim i As Long
    For i = 0 To CAL_MAX_TURNOS
        m_UltimoIntentoMin(i) = -1
    Next i
End Sub

' Copies the loaded temporaries over the live state. The running event (if any) is not touched.
Private Sub CalInstalar(ByRef cfg As t_CalConfig, ByRef turnos() As t_CalTurno, ByVal numTurnos As Long, ByRef variantes() As t_CalVariante, ByVal numVariantes As Long, ByRef resumen As t_CalResumen)
    Dim i As Long
    ReDim m_Turnos(1 To CAL_MAX_TURNOS) As t_CalTurno
    ReDim m_Variantes(1 To CAL_MAX_TURNOS) As t_CalVariante
    For i = 1 To numTurnos
        m_Turnos(i) = turnos(i)
    Next i
    For i = 1 To numVariantes
        m_Variantes(i) = variantes(i)
    Next i
    m_Cfg = cfg
    m_NumTurnos = numTurnos
    m_NumVariantes = numVariantes
    m_Resumen = resumen
    Call CalReiniciarEstado(numTurnos)
    m_Cargado = True
End Sub

Private Function CalTextoResumen(ByRef resumen As t_CalResumen) As String
    CalTextoResumen = "turnos " & resumen.TurnosCargados & " loaded / " & resumen.TurnosDescartados & " discarded, variants " & resumen.VariantesValidas & " valid / " & resumen.VariantesDescartadas & " discarded, warnings " & resumen.Advertencias & ", window " & m_Cfg.VentanaMin & " min, retry " & m_Cfg.ReintentoMin & " min"
End Function

' Loads the file into temporaries and swaps it in only if the file exists. Returns True if it was installed.
Private Function CalCargarArchivo(ByVal etiqueta As String) As Boolean
    On Error GoTo CalCargarArchivo_Err
    Dim cfg       As t_CalConfig
    Dim turnos()  As t_CalTurno
    Dim numTurnos As Long
    Dim vars()    As t_CalVariante
    Dim numVars   As Long
    Dim resumen   As t_CalResumen
    If Not CalendarioConfig_Cargar(CalArchivo(), cfg, turnos, numTurnos, vars, numVars, resumen) Then Exit Function
    Call CalInstalar(cfg, turnos, numTurnos, vars, numVars, resumen)
    Call LogServerLifecycle("[Calendario] " & etiqueta & ": " & CalTextoResumen(m_Resumen))
    CalCargarArchivo = True
    Exit Function
CalCargarArchivo_Err:
    Call TraceError(Err.Number, Err.Description, "ModCalendarioEventos.CalCargarArchivo", Erl)
End Function

' Server start (General.Main, after CargarInfoEventos). A missing file leaves an empty calendar (the server
' starts normally). Rows left OPEN by the previous run are closed: that event died with the process.
Public Sub CalendarioEventos_Cargar()
    On Error GoTo CalendarioEventos_Cargar_Err
    Dim rs As ADODB.Recordset
    Set rs = Query("UPDATE event_calendar_run SET estado = ?, detalle = ?, actualizado = ? WHERE estado = ?;", "CANCELADO", "server restarted with the event open", CalMarcaTiempo(), "ABIERTO")
    If rs Is Nothing Then
        Call CalAvisarSinPersistencia
    ElseIf rs.State = adStateOpen Then
        Call rs.Close
    End If
    If Not CalCargarArchivo("loaded") Then
        Dim vacioCfg   As t_CalConfig
        Dim vacioTur() As t_CalTurno
        Dim vacioVar() As t_CalVariante
        Dim vacioRes   As t_CalResumen
        ReDim vacioTur(1 To CAL_MAX_TURNOS) As t_CalTurno
        ReDim vacioVar(1 To CAL_MAX_TURNOS) As t_CalVariante
        vacioCfg.VentanaMin = CAL_VENTANA_MIN_DEFAULT
        vacioCfg.ReintentoMin = CAL_REINTENTO_MIN_DEFAULT
        vacioCfg.AvisoPrevioMin = CAL_AVISO_PREVIO_DEFAULT
        vacioCfg.BloquearConInvasion = True
        Call CalInstalar(vacioCfg, vacioTur, 0, vacioVar, 0, vacioRes)
        Call LogServerLifecycle("[Calendario] loaded: no file, empty calendar")
    End If
    Exit Sub
CalendarioEventos_Cargar_Err:
    Call TraceError(Err.Number, Err.Description, "ModCalendarioEventos.CalendarioEventos_Cargar", Erl)
End Sub

Private Sub CalAvisarSinPersistencia()
    If m_AvisoSinTabla Then Exit Sub
    m_AvisoSinTabla = True
    Call LogError("[Calendario] event_calendar_run is not available: the calendar runs in memory only (a restart inside a window can open the event again).")
End Sub

' ------------------------------------------------------------------------------------------------
' Persistence (synchronous Query, never the async Execute: this row is what stops a double open)
' ------------------------------------------------------------------------------------------------

Private Sub CalPersistirFila(ByVal clave As String, ByVal fecha As String, ByVal variante As String, ByVal estado As String, ByVal intentos As Long, ByVal detalle As String)
    On Error GoTo CalPersistirFila_Err
    Dim rs As ADODB.Recordset
    Set rs = Query("INSERT OR REPLACE INTO event_calendar_run (turno_key, fecha, variante, estado, intentos, detalle, actualizado) VALUES (?, ?, ?, ?, ?, ?, ?);", clave, fecha, variante, estado, intentos, detalle, CalMarcaTiempo())
    If rs Is Nothing Then
        Call CalAvisarSinPersistencia
    ElseIf rs.State = adStateOpen Then
        Call rs.Close
    End If
    Exit Sub
CalPersistirFila_Err:
    Call TraceError(Err.Number, Err.Description, "ModCalendarioEventos.CalPersistirFila", Erl)
End Sub

' Returns False if the table could not be read (memory only); existe = a row for (clave, fecha).
Private Function CalFilaExiste(ByVal clave As String, ByVal fecha As String, ByRef existe As Boolean) As Boolean
    On Error GoTo CalFilaExiste_Err
    Dim rs As ADODB.Recordset
    existe = False
    Set rs = Query("SELECT estado FROM event_calendar_run WHERE turno_key = ? AND fecha = ?;", clave, fecha)
    If rs Is Nothing Then
        Call CalAvisarSinPersistencia
        Exit Function
    End If
    existe = Not rs.EOF
    Call rs.Close
    CalFilaExiste = True
    Exit Function
CalFilaExiste_Err:
    Call TraceError(Err.Number, Err.Description, "ModCalendarioEventos.CalFilaExiste", Erl)
End Function

' ------------------------------------------------------------------------------------------------
' Tick
' ------------------------------------------------------------------------------------------------

' Players online EXCLUDING staff (an admin connected for the test must not make a "MinOnline 2" pass).
Private Function CalJugadoresConectados() As Long
    Dim i     As Long
    Dim total As Long
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            If (UserList(i).flags.Privilegios And (e_PlayerType.RoleMaster Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Dios Or e_PlayerType.Admin)) = 0 Then
                total = total + 1
            End If
        End If
    Next i
    CalJugadoresConectados = total
End Function

' Why the turno cannot open right now ("" = it can). Blocks of a lobby turno: server closing, an active public
' event (GM events included), an active invasion if BloquearConInvasion, fewer players than MinOnline.
' The hourly multiplier event (EventoActivo) does NOT block any more.
Private Function CalMotivoBloqueo(ByVal varIdx As Long) As String
    Dim online As Long
    If m_SegActivo Then
        CalMotivoBloqueo = "another calendar event is still being followed"
        Exit Function
    End If
    If ServidorCerrandoEventos() Then
        CalMotivoBloqueo = "the server is closing"
        Exit Function
    End If
    If ActiveEventType() <> NO_ACTIVE_EVENT Then
        CalMotivoBloqueo = "a public event is active (type " & ActiveEventType() & ")"
        Exit Function
    End If
    If m_Cfg.BloquearConInvasion Then
        If HayInvasionActiva() Then
            CalMotivoBloqueo = "an invasion is active"
            Exit Function
        End If
    End If
    online = CalJugadoresConectados()
    If online < m_Variantes(varIdx).MinOnline Then
        CalMotivoBloqueo = "only " & online & " players online, MinOnline is " & m_Variantes(varIdx).MinOnline
        Exit Function
    End If
End Function

Public Sub CalendarioEventos_Tick()
    On Error GoTo CalendarioEventos_Tick_Err
    Dim ahora       As Date
    Dim ahoraMin    As Long
    Dim i           As Long
    Dim fechaInicio As String
    If Not m_Cargado Then Exit Sub
    ' Peak of registered players of the event being followed (the hook only sees it after the lobby is emptied).
    If m_SegActivo Then
        If m_SegLobby < LBound(LobbyList) Or m_SegLobby > UBound(LobbyList) Then
            Call LogError("[Calendario] followed lobby index out of range (" & m_SegLobby & "): tracking dropped.")
            m_SegActivo = False
        ElseIf LobbyList(m_SegLobby).RegisteredPlayers > m_SegPico Then
            m_SegPico = LobbyList(m_SegLobby).RegisteredPlayers
        End If
    End If
    If m_NumTurnos <= 0 Then Exit Sub
    ahora = CalAhora()
    ahoraMin = CalMinutosDesdeEpoca(ahora)
    For i = 1 To m_NumTurnos
        If TurnoDebeDisparar(m_Turnos(i).DiasMask, m_Turnos(i).MinutoDelDia, m_Turnos(i).VentanaMin, ahora, m_FechaProcesada(i), fechaInicio) Then
            ' Every turno in its window is evaluated: while another event is followed it is blocked and registered as postponed.
            Call CalEvaluarTurno(i, ahoraMin, fechaInicio)
        ElseIf LenB(m_PendienteFecha(i)) <> 0 Then
            Call CalMarcarOmitido(i)
        End If
    Next i
    Exit Sub
CalendarioEventos_Tick_Err:
    Call TraceError(Err.Number, Err.Description, "ModCalendarioEventos.CalendarioEventos_Tick", Erl)
End Sub

Private Sub CalEvaluarTurno(ByVal i As Long, ByVal ahoraMin As Long, ByVal fechaInicio As String)
    On Error GoTo CalEvaluarTurno_Err
    Dim existe     As Boolean
    Dim motivo     As String
    Dim varIdx     As Long
    Dim abrio      As Boolean
    If m_PendienteFecha(i) <> fechaInicio Then
        ' First time this occurrence is looked at.
        m_Intentos(i) = 0
        m_UltimoIntentoMin(i) = -1
    End If
    ' Retry brake: after a blocked attempt, wait ReintentoMin minutes.
    If m_UltimoIntentoMin(i) >= 0 Then
        If ahoraMin - m_UltimoIntentoMin(i) < m_Cfg.ReintentoMin Then Exit Sub
    End If
    ' A row for this (turno, date) means it was handled already (the server restarted inside the window).
    If CalFilaExiste(m_Turnos(i).Clave, fechaInicio, existe) Then
        If existe Then
            m_FechaProcesada(i) = fechaInicio
            m_PendienteFecha(i) = vbNullString
            Call LogInfoServidor("[Calendario] " & m_Turnos(i).Texto & " (" & m_Turnos(i).Clave & ") already has a row for " & fechaInicio & ": not opened again.")
            Exit Sub
        End If
    End If
    varIdx = m_Turnos(i).VarIdx
    motivo = CalMotivoBloqueo(varIdx)
    If LenB(motivo) = 0 Then
        abrio = CreatePublicEvent(m_Variantes(varIdx).Settings, motivo)
    End If
    m_Intentos(i) = m_Intentos(i) + 1
    If Not abrio Then
        m_UltimoIntentoMin(i) = ahoraMin
        m_PendienteFecha(i) = fechaInicio
        m_UltimoMotivo(i) = motivo
        Call LogInfoServidor("[Calendario] " & m_Turnos(i).Texto & " postponed (attempt " & m_Intentos(i) & "): " & motivo & ".")
        Exit Sub
    End If
    m_FechaProcesada(i) = fechaInicio
    m_PendienteFecha(i) = vbNullString
    If GlobalLobbyIndex < LBound(LobbyList) Or GlobalLobbyIndex > UBound(LobbyList) Then
        ' Opened but the lobby is unknown: it cannot be followed. Leave the row and ONE log line, no tracking.
        Call LogError("[Calendario] " & m_Turnos(i).Texto & " opened but GlobalLobbyIndex is " & GlobalLobbyIndex & ": the event is not tracked.")
        Call CalPersistirFila(m_Turnos(i).Clave, fechaInicio, m_Variantes(varIdx).Nombre, "ABIERTO", m_Intentos(i), "lobby not tracked")
        Exit Sub
    End If
    m_SegActivo = True
    m_SegLobby = GlobalLobbyIndex
    m_SegClave = m_Turnos(i).Clave
    m_SegFecha = fechaInicio
    m_SegVariante = m_Variantes(varIdx).Nombre
    m_SegIntentos = m_Intentos(i)
    m_SegPico = 0
    Call CalPersistirFila(m_SegClave, m_SegFecha, m_SegVariante, "ABIERTO", m_SegIntentos, "lobby=" & m_SegLobby)
    Call LogAudit("EVENTO_AUTO", "ABIERTO turno=" & m_Turnos(i).Clave & " fecha=" & fechaInicio & " variante=" & m_SegVariante & " lobby=" & m_SegLobby & " intentos=" & m_SegIntentos)
    Exit Sub
CalEvaluarTurno_Err:
    Call TraceError(Err.Number, Err.Description, "ModCalendarioEventos.CalEvaluarTurno", Erl)
End Sub

' The window of an occurrence that was blocked every time is over: it ends OMITIDO, with ONE notice to the staff.
Private Sub CalMarcarOmitido(ByVal i As Long)
    On Error GoTo CalMarcarOmitido_Err
    Dim fecha As String
    Dim texto As String
    fecha = m_PendienteFecha(i)
    m_FechaProcesada(i) = fecha
    m_PendienteFecha(i) = vbNullString
    texto = "OMITIDO " & m_Turnos(i).Texto & " (" & m_Turnos(i).Clave & ") " & fecha & ": window over after " & m_Intentos(i) & " attempts, last reason: " & m_UltimoMotivo(i)
    Call CalPersistirFila(m_Turnos(i).Clave, fecha, m_Variantes(m_Turnos(i).VarIdx).Nombre, "OMITIDO", m_Intentos(i), m_UltimoMotivo(i))
    Call LogError("[Calendario] " & texto)
    Call LogAudit("EVENTO_AUTO", texto)
    Call CalAvisarStaff("[Calendario] El turno " & m_Turnos(i).Texto & " no se pudo abrir y quedo OMITIDO: " & m_UltimoMotivo(i))
    Exit Sub
CalMarcarOmitido_Err:
    Call TraceError(Err.Number, Err.Description, "ModCalendarioEventos.CalMarcarOmitido", Erl)
End Sub

' Hook called by ModLobby.ReleaseLobby for EVERY released lobby (the single exit of all the ways an event ends:
' normal end, cancel for lack of players, GM cancel, force reset, server close). It closes the row of the event
' this module opened. It runs whether or not automatic_events is on.
Public Sub CalendarioEventos_LobbyLiberado(ByVal lobbyIndex As Integer)
    On Error GoTo CalendarioEventos_LobbyLiberado_Err
    Dim estado  As String
    Dim detalle As String
    If Not m_SegActivo Then Exit Sub
    If lobbyIndex <> m_SegLobby Then Exit Sub
    If LobbyList(lobbyIndex).Canceled Then
        estado = "CANCELADO"
    Else
        estado = "TERMINADO"
    End If
    If LobbyList(lobbyIndex).RegisteredPlayers > m_SegPico Then m_SegPico = LobbyList(lobbyIndex).RegisteredPlayers
    detalle = "lobby=" & lobbyIndex & " inscriptos_pico=" & m_SegPico
    Call CalPersistirFila(m_SegClave, m_SegFecha, m_SegVariante, estado, m_SegIntentos, detalle)
    Call LogAudit("EVENTO_AUTO", estado & " turno=" & m_SegClave & " fecha=" & m_SegFecha & " variante=" & m_SegVariante & " " & detalle)
    m_SegActivo = False
    Exit Sub
CalendarioEventos_LobbyLiberado_Err:
    m_SegActivo = False
    Call TraceError(Err.Number, Err.Description, "ModCalendarioEventos.CalendarioEventos_LobbyLiberado", Erl)
End Sub

' ------------------------------------------------------------------------------------------------
' GM chat command: ~CALENDARIO [ESTADO] | ~CALENDARIO RECARGAR   (SemiDios / Dios / Admin)
' ------------------------------------------------------------------------------------------------

Private Sub CalDecir(ByVal userIndex As Integer, ByVal texto As String)
    Call WriteConsoleMsg(userIndex, texto, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
End Sub

Public Function HandleCalendarioChatCommand(ByVal userIndex As Integer, ByVal chat As String) As Boolean
    On Error GoTo HandleCalendarioChatCommand_Err
    Dim partes() As String
    Dim accion   As String
    partes = Split(Trim$(chat), " ")
    If UCase$(partes(0)) <> "~CALENDARIO" Then Exit Function
    ' The verb is ours whatever the subcommand: returning False would make HandleTalk say it out loud in the area.
    HandleCalendarioChatCommand = True
    If (UserList(userIndex).flags.Privilegios And (e_PlayerType.SemiDios Or e_PlayerType.Dios Or e_PlayerType.Admin)) = 0 Then Exit Function
    If UBound(partes) >= 1 Then accion = UCase$(Trim$(partes(1)))
    Select Case accion
        Case vbNullString, "ESTADO"
            Call LogGM(GetUserRealName(userIndex), "~CALENDARIO ESTADO")
            Call CalMostrarEstado(userIndex)
        Case "RECARGAR"
            Call LogGM(GetUserRealName(userIndex), "~CALENDARIO RECARGAR")
            Call CalRecargarPorGM(userIndex)
        Case Else
            Call CalDecir(userIndex, "[Calendario] Uso: ~CALENDARIO [ESTADO] | ~CALENDARIO RECARGAR")
    End Select
    Exit Function
HandleCalendarioChatCommand_Err:
    Call TraceError(Err.Number, Err.Description, "ModCalendarioEventos.HandleCalendarioChatCommand", Erl)
End Function

Private Sub CalRecargarPorGM(ByVal userIndex As Integer)
    On Error GoTo CalRecargarPorGM_Err
    If CalCargarArchivo("reloaded") Then
        Call CalDecir(userIndex, "[Calendario] Recargado: " & CalTextoResumen(m_Resumen))
        Call CalAvisarStaff("[Calendario] Recargado por " & GetUserRealName(userIndex) & ": " & CalTextoResumen(m_Resumen))
    Else
        Call CalDecir(userIndex, "[Calendario] No existe " & CalArchivo() & ": se mantiene el calendario actual.")
    End If
    Exit Sub
CalRecargarPorGM_Err:
    Call TraceError(Err.Number, Err.Description, "ModCalendarioEventos.CalRecargarPorGM", Erl)
End Sub

Private Sub CalMostrarEstado(ByVal userIndex As Integer)
    On Error GoTo CalMostrarEstado_Err
    Dim ahora       As Date
    Dim i           As Long
    Dim fechaInicio As String
    Dim linea       As String
    Dim rs          As ADODB.Recordset
    ahora = CalAhora()
    If IsFeatureEnabled("automatic_events") Then
        linea = "automatic_events=ON"
    Else
        linea = "automatic_events=OFF (nada se abre solo)"
    End If
    Call CalDecir(userIndex, "[Calendario] " & linea & ". Hora del server " & Format$(ahora, "yyyy-mm-dd hh:nn") & ". " & CalTextoResumen(m_Resumen))
    For i = 1 To m_NumTurnos
        linea = m_Turnos(i).Texto & " [" & m_Turnos(i).Clave & "] ventana " & m_Turnos(i).VentanaMin & " min, " & CalNombreModo(m_Variantes(m_Turnos(i).VarIdx).Modo) & ", MinOnline " & m_Variantes(m_Turnos(i).VarIdx).MinOnline
        If CalTurnoEnVentana(m_Turnos(i).DiasMask, m_Turnos(i).MinutoDelDia, m_Turnos(i).VentanaMin, ahora, fechaInicio) Then
            If fechaInicio = m_FechaProcesada(i) Then
                linea = linea & " -> EN VENTANA, ya atendido"
            Else
                linea = linea & " -> EN VENTANA, intentos " & m_Intentos(i)
                If LenB(m_UltimoMotivo(i)) <> 0 Then linea = linea & " (" & m_UltimoMotivo(i) & ")"
            End If
        Else
            linea = linea & " -> proximo " & Format$(CalProximoInicio(m_Turnos(i).DiasMask, m_Turnos(i).MinutoDelDia, ahora), "yyyy-mm-dd hh:nn")
        End If
        Call CalDecir(userIndex, linea)
    Next i
    If m_NumTurnos = 0 Then Call CalDecir(userIndex, "[Calendario] No hay turnos cargados.")
    If m_SegActivo Then Call CalDecir(userIndex, "[Calendario] Siguiendo el evento " & m_SegVariante & " (lobby " & m_SegLobby & ", turno " & m_SegClave & ").")
    Set rs = Query("SELECT turno_key, variante, estado, intentos FROM event_calendar_run WHERE fecha = ? ORDER BY turno_key;", Format$(ahora, "yyyy-mm-dd"))
    If rs Is Nothing Then
        Call CalDecir(userIndex, "[Calendario] La tabla event_calendar_run no esta disponible (modo memoria).")
        Exit Sub
    End If
    Do While Not rs.EOF
        Call CalDecir(userIndex, "[Calendario] Hoy: " & rs!turno_key & " " & rs!variante & " " & rs!estado & " intentos " & rs!intentos)
        Call rs.MoveNext
    Loop
    Call rs.Close
    Exit Sub
CalMostrarEstado_Err:
    Call TraceError(Err.Number, Err.Description, "ModCalendarioEventos.CalMostrarEstado", Erl)
End Sub
