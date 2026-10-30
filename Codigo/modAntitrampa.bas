Attribute VB_Name = "modAntitrampa"
' Argentum 20 Game Server
'
'    Copyright (C) 2026 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
' modAntitrampa (plan 2026-10-29-seguridad-antitrampas, etapa E1)
'
' 1. Log estructurado: Logs\Antitrampa\Antitrampa_AAAA-MM-DD.jsonl, una linea JSON por evento
'    (ts, tick, evento, accion, cuenta, pj, nick, ip, mapa, x, y, detalle). Se acumula en memoria y se
'    baja a disco cada FLUSH_MS o cada FLUSH_LINEAS: nunca Event Log ni RPC sincronica dentro del loop
'    single-thread (leccion de LogThis, docs/server/anti-macro-estado.md). Como mucho una linea por
'    jugador por tipo de evento cada m_LogCadaMs; las del medio se cuentan en la siguiente.
' 2. Anti-speedhack: credito en ms tolerante a rafagas (lag). Cada paquete de caminar suma el tiempo
'    real transcurrido (con tope m_CreditoMaxMs) y cada paso aceptado cuesta
'    IntervaloCaminar / velocidad * m_ToleranciaPct / 100. Sin credito el paso se rechaza
'    (WritePosUpdate) con CUALQUIER politica: la politica solo decide que pasa con el evento.
'    Ya NO depende de PacketRatePolicy.ini [Walk] Limit (que vale 0 y lo dejaba muerto).
'    Si el loop principal se traba mas de m_StallMs (guardado, disco), los pasos encolados llegan juntos:
'    AntitrampaTickLoop les devuelve el credito lleno a todos para no marcar a nadie por eso.
' 3. Rechazos de IntervaloPermite*: solo se cuentan; con m_IntervaloRechazosEvento en la ventana
'    queda un evento. No cambia que se rechaza.
' 4. Contadores anti-replay de CastSpell/Drop/EquipItem/Work: solo registro, nunca bloqueo.
' 5. Politica (Configuracion.ini [ANTITRAMPA] Accion, o en caliente con ~ANTITRAMPA, solo Dios/Admin):
'    log | avisar_gm | expulsar. Nunca banea. Solo el speedhack aplica la politica; 3 y 4 solo registran.
'    El anti-macro de verifyTimeStamp NO pasa por aca (decision del duenio 2026-08-08).
' 6. Anti-farmeo de kills (etapa E1b, duenio 2026-10-29: "no cuentan las muertes entre misma IP/PC, mas
'    el tope diario por pareja"). Lo consultan ContarMuerte y ActStats (Modulo_UsUaRiOs.bas):
'    - Misma IP PUBLICA real (o misma PC cuando exista E4) = la muerte no suma kills, puntos de faccion ni
'      experiencia; evento `kill_misma_ip`. Nunca cuentan como "misma IP": loopback, privadas, link-local,
'      CGNAT, IPv6 y las de KillIPsTunel (detras de playit todos llegan como 127.0.0.1, H12). El toggle
'      kill_regla_misma_ip apaga esta regla. KillIncluirLoopback=1 (o ~ANTITRAMPA kills loopback si) trata
'      127.x como publica: SOLO para las sondas antitrampa_kills / antitrampa_retos.
'    - Retos (duenio 2026-10-29): la misma regla de origen entre RIVALES apaga el ELO del reto entero
'      (ModRetos.RetoRivalesMismoOrigen, evento `reto_misma_ip`); la apuesta no cambia.
'    - Tope por pareja asesino->victima (Id de personaje): KillsMismaParejaPorDia (Balance.dat [EXTRA],
'      default 3, 0 = sin tope) muertes que cuentan por dia; la siguiente no suma (`kill_tope_diario`). En
'      memoria: se vacia al cambiar la fecha (00:00) y con un reinicio (aceptado, plan). FactionReKillTime
'      (10 min) se sigue aplicando antes: solo las muertes que ya contaban gastan tope.
'    Solo registro (no son trampas en si): sin aviso a GMs ni politica. Ante un error propio cuenta (falla
'    abierto: nunca le saca un kill a nadie por un bug de este modulo).
Option Explicit

Public Enum e_AntitrampaAccion
    eAntitrampaLog = 0
    eAntitrampaAvisarGm = 1
    eAntitrampaExpulsar = 2
End Enum

' Indice de t_AntitrampaUser.IntervaloRechazos (Declares.bas): 0 To 6.
Public Enum e_AntitrampaIntervalo
    eAtiAtacar = 0
    eAtiHechizo = 1
    eAtiArco = 2
    eAtiUsar = 3
    eAtiUsarClick = 4
    eAtiTirar = 5
    eAtiCombo = 6
End Enum

' Indice de t_AntitrampaUser.UltimoLog / SinLog (Declares.bas): 0 To 2.
Private Const TIPO_SPEEDHACK As Long = 0
Private Const TIPO_INTERVALO As Long = 1
Private Const TIPO_CONTADOR As Long = 2
Private Const TIPO_MAX As Long = 2

Private Const INTERVALO_CATEGORIA_MAX As Long = 6
Private Const FLUSH_MS As Long = 1000
Private Const FLUSH_LINEAS As Long = 64
Private Const LOG_ARBOL As String = "Antitrampa"
Private Const CONTADOR_TOPE As Long = 1000000
Private Const KILL_TOGGLE_IP As String = "kill_regla_misma_ip"

Private m_Accion As e_AntitrampaAccion
Private m_ToleranciaPct As Long
Private m_CreditoMaxMs As Long
Private m_SpeedRechazosEvento As Long
Private m_SpeedVentanaMs As Long
Private m_IntervaloRechazosEvento As Long
Private m_IntervaloVentanaMs As Long
Private m_LogCadaMs As Long
Private m_AvisoGmCadaMs As Long
Private m_StallMs As Long

Private m_Buffer As String
Private m_BufferLineas As Long
Private m_BufferFecha As String
Private m_UltimoFlush As Long
Private m_CarpetaLista As Boolean
Private m_UltimoFrame As Long

Private m_KillIncluirLoopback As Boolean
Private m_KillIPsTunel As String              ' ",ip1,ip2," en minuscula ("," = vacia)
Private m_KillsPareja As Dictionary          ' "asesinoId>victimaId" -> muertes del dia que llegaron al tope
Private m_KillsFecha As String               ' yyyymmdd de m_KillsPareja

Private Sub CargarDefaults()
    ' Si Configuracion.ini falla, el modulo queda con los valores del plan (falla abierto, no muerto).
    m_Accion = eAntitrampaAvisarGm
    m_ToleranciaPct = 85
    m_CreditoMaxMs = 5000
    m_SpeedRechazosEvento = 5
    m_SpeedVentanaMs = 10000
    m_IntervaloRechazosEvento = 20
    m_IntervaloVentanaMs = 10000
    m_LogCadaMs = 5000
    m_AvisoGmCadaMs = 60000
    m_StallMs = 1000
    m_KillIncluirLoopback = False
    m_KillIPsTunel = ","
End Sub

' Lee la configuracion. La llama General.Main despues de cargar Configuracion.ini.
Public Sub InitAntitrampa()
    On Error GoTo InitAntitrampa_Err
    Call CargarDefaults
    m_UltimoFlush = GetTickCountRaw()
    m_Accion = ParseAccion(SvrConfig.GetValue("AntitrampaAccion"))
    m_ToleranciaPct = SvrConfig.GetValue("AntitrampaSpeedToleranciaPct")
    m_CreditoMaxMs = SvrConfig.GetValue("AntitrampaSpeedCreditoMaxMs")
    m_SpeedRechazosEvento = SvrConfig.GetValue("AntitrampaSpeedRechazosEvento")
    m_SpeedVentanaMs = SvrConfig.GetValue("AntitrampaSpeedVentanaMs")
    m_IntervaloRechazosEvento = SvrConfig.GetValue("AntitrampaIntervaloRechazosEvento")
    m_IntervaloVentanaMs = SvrConfig.GetValue("AntitrampaIntervaloVentanaMs")
    m_LogCadaMs = SvrConfig.GetValue("AntitrampaLogCadaMs")
    m_AvisoGmCadaMs = SvrConfig.GetValue("AntitrampaAvisoGmCadaMs")
    m_StallMs = SvrConfig.GetValue("AntitrampaStallMs")
    m_KillIncluirLoopback = (SvrConfig.GetValue("AntitrampaKillIncluirLoopback") <> 0)
    m_KillIPsTunel = NormalizarListaIps(SvrConfig.GetValue("AntitrampaKillIPsTunel"))
    If m_SpeedRechazosEvento < 1 Then m_SpeedRechazosEvento = 1
    If m_IntervaloRechazosEvento < 1 Then m_IntervaloRechazosEvento = 1
    If m_StallMs < 100 Then m_StallMs = 100
    Call LogThis(0, "Antitrampa: accion=" & AccionTexto(m_Accion) & " tolerancia=" & m_ToleranciaPct & "% credito_max_ms=" & m_CreditoMaxMs & _
            " speed_rechazos=" & m_SpeedRechazosEvento & "/" & m_SpeedVentanaMs & "ms intervalo_rechazos=" & m_IntervaloRechazosEvento & "/" & _
            m_IntervaloVentanaMs & "ms log_cada_ms=" & m_LogCadaMs & " stall_ms=" & m_StallMs & " kills_tope_pareja=" & _
            KillsMismaParejaPorDia & " kills_ips_tunel=" & CantidadIpsTunel() & " kills_loopback=" & m_KillIncluirLoopback, _
            vbLogEventTypeInformation)
    Exit Sub
InitAntitrampa_Err:
    Call CargarDefaults
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.InitAntitrampa (quedan los defaults)", Erl)
End Sub

Public Function ParseAccion(ByVal texto As String) As e_AntitrampaAccion
    Select Case LCase$(Trim$(texto))
        Case "log"
            ParseAccion = eAntitrampaLog
        Case "expulsar"
            ParseAccion = eAntitrampaExpulsar
        Case Else
            ' Default de la beta (duenio 2026-10-29): registrar y avisar al GM.
            ParseAccion = eAntitrampaAvisarGm
    End Select
End Function

Public Function AccionTexto(ByVal accion As e_AntitrampaAccion) As String
    Select Case accion
        Case eAntitrampaLog
            AccionTexto = "log"
        Case eAntitrampaExpulsar
            AccionTexto = "expulsar"
        Case Else
            AccionTexto = "avisar_gm"
    End Select
End Function

' Campo por campo (como el resto de TCP.ResetContadores), sin asignar el UDT entero.
Public Sub AntitrampaResetUser(ByVal UserIndex As Integer)
    On Error GoTo AntitrampaResetUser_Err
    Dim i As Long
    With UserList(UserIndex).Counters.Antitrampa
        .SpeedRechazos = 0
        .SpeedVentanaInicio = 0
        For i = 0 To INTERVALO_CATEGORIA_MAX
            .IntervaloRechazos(i) = 0
        Next i
        .IntervaloVentanaInicio = 0
        For i = 0 To TIPO_MAX
            .UltimoLog(i) = 0
            .SinLog(i) = 0
        Next i
        .UltimoAvisoGm = 0
        .EventosSinAvisar = 0
    End With
    Exit Sub
AntitrampaResetUser_Err:
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.AntitrampaResetUser", Erl)
End Sub

Private Function TickNoCero(ByVal nowRaw As Long) As Long
    ' 0 significa "sin marca" en todos los campos de tick del modulo.
    TickNoCero = nowRaw
    If TickNoCero = 0 Then TickNoCero = 1
End Function

' Primera llamada de cada vuelta del loop principal (General.Main), ANTES de procesar la red. Si la
' vuelta anterior tardo mas de m_StallMs, los paquetes de caminar que se encolaron en el medio van a
' llegar juntos: a todos se les devuelve el credito lleno (LastStep = 0) para que no cuenten como
' speedhack. Costo: tras una traba, un tramposo tambien arranca con credito lleno.
Public Sub AntitrampaTickLoop()
    On Error GoTo AntitrampaTickLoop_Err
    Dim nowRaw As Long, LoopC As Long
    nowRaw = GetTickCountRaw()
    If m_UltimoFrame <> 0 Then
        If TicksElapsed(m_UltimoFrame, nowRaw) > m_StallMs Then
            For LoopC = 1 To LastUser
                If UserList(LoopC).flags.UserLogged Then UserList(LoopC).Counters.LastStep = 0
            Next LoopC
        End If
    End If
    m_UltimoFrame = TickNoCero(nowRaw)
    Exit Sub
AntitrampaTickLoop_Err:
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.AntitrampaTickLoop", Erl)
End Sub

' ---------------------------------------------------------------------------------------------
' Log estructurado
' ---------------------------------------------------------------------------------------------

Private Function JsonTexto(ByVal texto As String) As String
    Dim i As Long, c As Long, out As String
    out = """"
    For i = 1 To Len(texto)
        c = AscW(Mid$(texto, i, 1))
        Select Case c
            Case 34
                out = out & "\"""
            Case 92
                out = out & "\\"
            Case 0 To 31
                out = out & "\u" & Right$("000" & Hex$(c), 4)
            Case Else
                out = out & Mid$(texto, i, 1)
        End Select
    Next i
    JsonTexto = out & """"
End Function

Private Function JsonEntero(ByVal valor As Double) As String
    JsonEntero = CStr(Fix(valor))
End Function

' Registra un evento del jugador. accion = lo que se hizo con el (log, avisar_gm, expulsar).
Public Sub AntitrampaRegistrar(ByVal UserIndex As Integer, ByVal evento As String, ByVal accion As String, ByVal detalle As String)
    On Error GoTo AntitrampaRegistrar_Err
    Dim linea As String
    With UserList(UserIndex)
        linea = "{""ts"":" & JsonTexto(Format$(Now, "yyyy-mm-dd hh:nn:ss")) & _
                ",""tick"":" & CStr(GetTickCountRaw()) & _
                ",""evento"":" & JsonTexto(evento) & _
                ",""accion"":" & JsonTexto(accion) & _
                ",""cuenta"":" & .AccountID & _
                ",""pj"":" & .Id & _
                ",""nick"":" & JsonTexto(.name) & _
                ",""ip"":" & JsonTexto(.ConnectionDetails.IP) & _
                ",""mapa"":" & .pos.Map & _
                ",""x"":" & .pos.x & _
                ",""y"":" & .pos.y & _
                ",""detalle"":" & JsonTexto(detalle) & "}"
    End With
    Call AgregarLinea(linea)
    Exit Sub
AntitrampaRegistrar_Err:
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.AntitrampaRegistrar", Erl)
End Sub

' Como mucho una linea por jugador y tipo cada m_LogCadaMs (forzar = siempre, p. ej. una expulsion).
' Los eventos suprimidos se cuentan y salen como " sin_registrar=N" en la linea siguiente.
Private Sub RegistrarLimitado(ByVal UserIndex As Integer, ByVal tipo As Long, ByVal evento As String, ByVal accion As String, ByVal detalle As String, _
                              ByVal forzar As Boolean)
    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()
    With UserList(UserIndex).Counters.Antitrampa
        If Not forzar And .UltimoLog(tipo) <> 0 Then
            If TicksElapsed(.UltimoLog(tipo), nowRaw) < m_LogCadaMs Then
                If .SinLog(tipo) < CONTADOR_TOPE Then .SinLog(tipo) = .SinLog(tipo) + 1
                Exit Sub
            End If
        End If
        If .SinLog(tipo) > 0 Then detalle = detalle & " sin_registrar=" & .SinLog(tipo)
        .SinLog(tipo) = 0
        .UltimoLog(tipo) = TickNoCero(nowRaw)
    End With
    Call AntitrampaRegistrar(UserIndex, evento, accion, detalle)
End Sub

Private Sub AgregarLinea(ByVal linea As String)
    Dim hoy As String
    hoy = Format$(Now, "yyyy-mm-dd")
    ' Cambio de dia: lo acumulado va al archivo del dia en que paso.
    If m_BufferLineas > 0 And hoy <> m_BufferFecha Then Call FlushAntitrampa
    If m_BufferLineas = 0 Then m_BufferFecha = hoy
    m_Buffer = m_Buffer & linea & vbCrLf
    m_BufferLineas = m_BufferLineas + 1
    If m_BufferLineas >= FLUSH_LINEAS Then Call FlushAntitrampa
End Sub

' Una vez por vuelta del loop principal: barato si no hay nada que bajar.
Public Sub MaybeFlushAntitrampa()
    If m_BufferLineas = 0 Then Exit Sub
    If TicksElapsed(m_UltimoFlush, GetTickCountRaw()) < FLUSH_MS Then Exit Sub
    Call FlushAntitrampa
End Sub

' OJO: igual que Logging.WriteToLogFile, el manejador NO llama TraceError (un log que falla se traga).
Public Sub FlushAntitrampa()
    Dim fn As Integer
    On Error GoTo ErrHandler
    m_UltimoFlush = GetTickCountRaw()
    If m_BufferLineas = 0 Then Exit Sub
    Dim carpeta As String
    carpeta = App.Path & "\Logs\" & LOG_ARBOL
    ' La carpeta se verifica una vez (Dir$ tiene estado global: no se llama en cada flush).
    If Not m_CarpetaLista Then
        If LenB(Dir$(App.Path & "\Logs", vbDirectory)) = 0 Then MkDir App.Path & "\Logs"
        If LenB(Dir$(carpeta, vbDirectory)) = 0 Then MkDir carpeta
        m_CarpetaLista = True
    End If
    fn = FreeFile
    Open carpeta & "\" & LOG_ARBOL & "_" & m_BufferFecha & ".jsonl" For Append As #fn
    Print #fn, m_Buffer;
    Close #fn
    m_Buffer = vbNullString
    m_BufferLineas = 0
    Exit Sub
ErrHandler:
    On Error Resume Next
    If fn <> 0 Then Close #fn
    ' Si alguien borro la carpeta, se vuelve a crear en el proximo flush.
    m_CarpetaLista = False
    ' Se descarta para no crecer sin limite si el disco falla.
    m_Buffer = vbNullString
    m_BufferLineas = 0
End Sub

' ---------------------------------------------------------------------------------------------
' Politica y avisos
' ---------------------------------------------------------------------------------------------

' Aviso agregado: como mucho uno por jugador cada m_AvisoGmCadaMs; los eventos del medio se cuentan
' y salen en el aviso siguiente (riesgo "ruido a los GMs" del plan).
Private Sub AvisarGms(ByVal UserIndex As Integer, ByVal evento As String, ByVal detalle As String, ByVal forzar As Boolean)
    Dim nowRaw As Long, texto As String
    nowRaw = GetTickCountRaw()
    With UserList(UserIndex).Counters.Antitrampa
        If Not forzar And .UltimoAvisoGm <> 0 Then
            If TicksElapsed(.UltimoAvisoGm, nowRaw) < m_AvisoGmCadaMs Then
                If .EventosSinAvisar < CONTADOR_TOPE Then .EventosSinAvisar = .EventosSinAvisar + 1
                Exit Sub
            End If
        End If
        texto = "Antitrampa: " & UserList(UserIndex).name & " - " & evento & " (" & detalle & ")"
        If .EventosSinAvisar > 0 Then texto = texto & " +" & .EventosSinAvisar & " eventos sin avisar"
        texto = texto & ". Accion: " & AccionTexto(m_Accion) & ". Ver Logs\Antitrampa."
        .UltimoAvisoGm = TickNoCero(nowRaw)
        .EventosSinAvisar = 0
    End With
    Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg(texto, e_TextChannel.TEXTCHANNEL_SERVER_STAFF, e_FontTypeNames.FONTTYPE_SERVER))
End Sub

' Registra el evento y aplica la politica. Devuelve True si se cerro la conexion: el llamador tiene que
' salir sin volver a tocar el slot.
Private Function AplicarPolitica(ByVal UserIndex As Integer, ByVal tipo As Long, ByVal evento As String, ByVal detalle As String) As Boolean
    Dim accion As e_AntitrampaAccion
    accion = m_Accion
    Call RegistrarLimitado(UserIndex, tipo, evento, AccionTexto(accion), detalle, accion = eAntitrampaExpulsar)
    Select Case accion
        Case eAntitrampaAvisarGm
            Call AvisarGms(UserIndex, evento, detalle, False)
        Case eAntitrampaExpulsar
            Call AvisarGms(UserIndex, evento, detalle, True)
            Call LogAudit("ANTITRAMPA", "Expulsado (sin ban) pj=" & UserList(UserIndex).Id & " cuenta=" & UserList(UserIndex).AccountID & " evento=" & evento & " " & detalle)
            Call FlushAntitrampa
            Call CloseSocket(UserIndex)
            AplicarPolitica = True
    End Select
End Function

' ---------------------------------------------------------------------------------------------
' Anti-speedhack (HandleWalk)
' ---------------------------------------------------------------------------------------------

' Suma al credito el tiempo real desde el ultimo paquete de caminar (con tope) y devuelve el costo del
' paso si alcanza, o -1 si el paso tiene que rechazarse. No descuenta: eso lo hace AntitrampaPasoAceptado
' solo si el server realmente movio al personaje (un paso contra una pared no gasta credito).
Public Function AntitrampaCostoPaso(ByVal UserIndex As Integer, ByVal nowRaw As Long) As Double
    On Error GoTo AntitrampaCostoPaso_Err
    If m_ToleranciaPct <= 0 Then Exit Function
    Dim velocidad As Single, credito As Double, costo As Double, tope As Double
    With UserList(UserIndex)
        velocidad = .Char.speeding
        ' Recien logueado speeding llega en 0 (nadie recalcula la velocidad en el login) y el cliente
        ' camina a IntervaloCaminar: es la velocidad normal.
        If velocidad <= 0 Then velocidad = VelocidadNormal
        costo = CDbl(IntervaloCaminar) / CDbl(velocidad) * CDbl(m_ToleranciaPct) / 100#
        ' Con un tope menor al costo de un paso se rechazaria todo para siempre: minimo 3 pasos.
        tope = m_CreditoMaxMs
        If tope < costo * 3 Then tope = costo * 3
        If .Counters.LastStep = 0 Then
            credito = tope
        Else
            credito = CDbl(.Counters.SpeedHackCounter) + TicksElapsed(.Counters.LastStep, nowRaw)
        End If
        If credito > tope Then credito = tope
        .Counters.SpeedHackCounter = credito
        .Counters.LastStep = TickNoCero(nowRaw)
        If credito < costo Then
            AntitrampaCostoPaso = -1
        Else
            AntitrampaCostoPaso = costo
        End If
    End With
    Exit Function
AntitrampaCostoPaso_Err:
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.AntitrampaCostoPaso", Erl)
End Function

Public Sub AntitrampaPasoAceptado(ByVal UserIndex As Integer, ByVal costo As Double)
    If costo <= 0 Then Exit Sub
    With UserList(UserIndex).Counters
        .SpeedHackCounter = .SpeedHackCounter - costo
        If .SpeedHackCounter < 0 Then .SpeedHackCounter = 0
    End With
End Sub

' Cuenta el paso rechazado en la ventana; al llegar al umbral registra `speedhack` y aplica la politica.
' Devuelve True si el jugador fue expulsado.
Public Function AntitrampaPasoRechazado(ByVal UserIndex As Integer, ByVal nowRaw As Long) As Boolean
    On Error GoTo AntitrampaPasoRechazado_Err
    Dim detalle As String, velocidad As Single
    velocidad = UserList(UserIndex).Char.speeding
    If velocidad <= 0 Then velocidad = VelocidadNormal
    With UserList(UserIndex).Counters.Antitrampa
        If .SpeedVentanaInicio = 0 Or TicksElapsed(.SpeedVentanaInicio, nowRaw) >= m_SpeedVentanaMs Then
            .SpeedVentanaInicio = TickNoCero(nowRaw)
            .SpeedRechazos = 0
        End If
        .SpeedRechazos = .SpeedRechazos + 1
        If .SpeedRechazos < m_SpeedRechazosEvento Then Exit Function
        detalle = "pasos_rechazados=" & .SpeedRechazos & " en_ms=" & JsonEntero(TicksElapsed(.SpeedVentanaInicio, nowRaw)) & _
                " paso_ms=" & JsonEntero(CDbl(IntervaloCaminar) / CDbl(velocidad)) & _
                " velocidad=" & Replace$(CStr(UserList(UserIndex).Char.speeding), ",", ".")
        .SpeedRechazos = 0
        .SpeedVentanaInicio = TickNoCero(nowRaw)
    End With
    AntitrampaPasoRechazado = AplicarPolitica(UserIndex, TIPO_SPEEDHACK, "speedhack", detalle)
    Exit Function
AntitrampaPasoRechazado_Err:
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.AntitrampaPasoRechazado", Erl)
End Function

' ---------------------------------------------------------------------------------------------
' Rechazos de IntervaloPermite* (solo registro)
' ---------------------------------------------------------------------------------------------

Private Function NombreIntervalo(ByVal categoria As Long) As String
    Select Case categoria
        Case eAtiAtacar: NombreIntervalo = "atacar"
        Case eAtiHechizo: NombreIntervalo = "hechizo"
        Case eAtiArco: NombreIntervalo = "arco"
        Case eAtiUsar: NombreIntervalo = "usar"
        Case eAtiUsarClick: NombreIntervalo = "usar_click"
        Case eAtiTirar: NombreIntervalo = "tirar"
        Case Else: NombreIntervalo = "combo"
    End Select
End Function

Public Sub AntitrampaRechazoIntervalo(ByVal UserIndex As Integer, ByVal categoria As e_AntitrampaIntervalo)
    On Error GoTo AntitrampaRechazoIntervalo_Err
    If UserIndex <= 0 Then Exit Sub
    Dim nowRaw As Long, i As Long, total As Long, detalle As String
    nowRaw = GetTickCountRaw()
    With UserList(UserIndex).Counters.Antitrampa
        If .IntervaloVentanaInicio = 0 Or TicksElapsed(.IntervaloVentanaInicio, nowRaw) >= m_IntervaloVentanaMs Then
            For i = 0 To INTERVALO_CATEGORIA_MAX
                .IntervaloRechazos(i) = 0
            Next i
            .IntervaloVentanaInicio = TickNoCero(nowRaw)
        End If
        If .IntervaloRechazos(categoria) < 32000 Then .IntervaloRechazos(categoria) = .IntervaloRechazos(categoria) + 1
        For i = 0 To INTERVALO_CATEGORIA_MAX
            total = total + .IntervaloRechazos(i)
        Next i
        If total < m_IntervaloRechazosEvento Then Exit Sub
        detalle = "rechazos=" & total & " en_ms=" & JsonEntero(TicksElapsed(.IntervaloVentanaInicio, nowRaw))
        For i = 0 To INTERVALO_CATEGORIA_MAX
            If .IntervaloRechazos(i) > 0 Then detalle = detalle & " " & NombreIntervalo(i) & "=" & .IntervaloRechazos(i)
            .IntervaloRechazos(i) = 0
        Next i
        .IntervaloVentanaInicio = TickNoCero(nowRaw)
    End With
    Call RegistrarLimitado(UserIndex, TIPO_INTERVALO, "intervalo_rechazado", "log", detalle, False)
    Exit Sub
AntitrampaRechazoIntervalo_Err:
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.AntitrampaRechazoIntervalo", Erl)
End Sub

' ---------------------------------------------------------------------------------------------
' Contadores anti-replay de los handlers que antes solo los leian (H6): solo registro
' ---------------------------------------------------------------------------------------------

' Si el contador del cliente no crece (paquete grabado y reenviado) queda `contador_repetido` (con el
' limite por jugador de RegistrarLimitado). Nunca bloquea ni cierra (VerifyPacketSequence si cierra: no
' se usa aca).
Public Sub AntitrampaContador(ByVal UserIndex As Integer, ByVal slot As PacketNames, ByVal recibido As Long, ByVal paquete As String)
    On Error GoTo AntitrampaContador_Err
    Dim anterior As Long
    With UserList(UserIndex)
        anterior = .PacketCounters(slot)
        .PacketCounters(slot) = recibido
    End With
    If recibido > anterior Then Exit Sub
    Call RegistrarLimitado(UserIndex, TIPO_CONTADOR, "contador_repetido", "log", "paquete=" & paquete & " contador=" & recibido & " anterior=" & anterior, False)
    Exit Sub
AntitrampaContador_Err:
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.AntitrampaContador", Erl)
End Sub

' ---------------------------------------------------------------------------------------------
' Comando de GM "~ANTITRAMPA [log|avisar_gm|expulsar]" (viaja como chat, igual que ~LOGLEVEL).
' "~ANTITRAMPA kills [nuevo_dia | loopback si|no]": estado de la regla de kills (E1b); nuevo_dia vacia el
' tope diario (lo mismo que pasa a las 00:00); loopback es solo para pruebas y vuelve al .ini al reiniciar.
' Consultar: SemiDios/Dios/Admin (el gate de Protocol.HandleTalk). Cambiar: solo Dios/Admin.
' ---------------------------------------------------------------------------------------------

Private Sub ReiniciarLimitesDeLog()
    ' Tras cambiar la politica, el proximo evento de cada jugador se registra con la accion nueva.
    Dim LoopC As Long, t As Long
    For LoopC = 1 To LastUser
        With UserList(LoopC).Counters.Antitrampa
            For t = 0 To TIPO_MAX
                .UltimoLog(t) = 0
            Next t
        End With
    Next LoopC
End Sub

Public Function HandleAntitrampaChatCommand(ByVal UserIndex As Integer, ByVal chat As String) As Boolean
    On Error GoTo HandleAntitrampaChatCommand_Err
    Dim partes() As String
    partes = Split(Trim$(chat), " ")
    ' Split de "" devuelve un arreglo vacio (UBound = -1): partes(0) daria error 9.
    If UBound(partes) < 0 Then Exit Function
    If UCase$(Trim$(partes(0))) <> "~ANTITRAMPA" Then Exit Function
    HandleAntitrampaChatCommand = True
    If UBound(partes) >= 1 Then
        If LCase$(Trim$(partes(1))) = "kills" Then
            Call ComandoKills(UserIndex, partes)
            Exit Function
        End If
    End If
    If UBound(partes) >= 1 Then
        If (UserList(UserIndex).flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin)) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Antitrampa: solo Dios/Admin pueden cambiar la accion.", e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
        Else
            Select Case LCase$(Trim$(partes(1)))
                Case "log", "avisar_gm", "expulsar"
                    m_Accion = ParseAccion(partes(1))
                    Call ReiniciarLimitesDeLog
                    Call LogAudit("ANTITRAMPA", "GM " & UserList(UserIndex).name & " cambio la accion a " & AccionTexto(m_Accion) & " (hasta el proximo reinicio)")
                Case Else
                    Call WriteConsoleMsg(UserIndex, "Uso: ~ANTITRAMPA log|avisar_gm|expulsar", e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Function
            End Select
        End If
    End If
    Call WriteConsoleMsg(UserIndex, "Antitrampa: accion = " & AccionTexto(m_Accion), e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
    Exit Function
HandleAntitrampaChatCommand_Err:
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.HandleAntitrampaChatCommand", Erl)
End Function

' ---------------------------------------------------------------------------------------------
' Anti-farmeo de kills (E1b)
' ---------------------------------------------------------------------------------------------

Private Function NormalizarListaIps(ByVal lista As String) As String
    Dim partes() As String, i As Long, salida As String
    salida = ","
    partes = Split(Replace$(lista, " ", vbNullString), ",")
    ' Split de "" devuelve UBound = -1: el For no entra.
    For i = 0 To UBound(partes)
        If LenB(partes(i)) > 0 Then salida = salida & LCase$(partes(i)) & ","
    Next i
    NormalizarListaIps = salida
End Function

Private Function CantidadIpsTunel() As Long
    CantidadIpsTunel = Len(m_KillIPsTunel) - Len(Replace$(m_KillIPsTunel, ",", vbNullString)) - 1
    If CantidadIpsTunel < 0 Then CantidadIpsTunel = 0
End Function

' True solo si la IP identifica a una conexion propia: IPv4 publica que no esta en KillIPsTunel. Loopback
' (salvo KillIncluirLoopback), 0/8, 10/8, 100.64/10 (CGNAT), 169.254/16, 172.16/12, 192.168/16 e IPv6 no:
' varios jugadores distintos pueden llegar con la misma.
Public Function AntitrampaIpEsPublica(ByVal ip As String) As Boolean
    On Error GoTo AntitrampaIpEsPublica_Err
    Dim partes() As String, i As Long, octeto As Long, primero As Long, segundo As Long
    ip = LCase$(Trim$(ip))
    If LenB(ip) = 0 Then Exit Function
    If InStr(1, m_KillIPsTunel, "," & ip & ",", vbBinaryCompare) > 0 Then Exit Function
    ' IPv6 o "ip:puerto": Aurora.Network hoy entrega IPv4 pelada; lo demas no se compara.
    If InStr(ip, ":") > 0 Then Exit Function
    partes = Split(ip, ".")
    If UBound(partes) <> 3 Then Exit Function
    For i = 0 To 3
        If LenB(partes(i)) = 0 Or Not IsNumeric(partes(i)) Then Exit Function
        octeto = CLng(val(partes(i)))
        If octeto < 0 Or octeto > 255 Then Exit Function
    Next i
    primero = CLng(val(partes(0)))
    segundo = CLng(val(partes(1)))
    Select Case primero
        Case 127
            AntitrampaIpEsPublica = m_KillIncluirLoopback
        Case 0, 10
            AntitrampaIpEsPublica = False
        Case 100
            AntitrampaIpEsPublica = Not (segundo >= 64 And segundo <= 127)
        Case 169
            AntitrampaIpEsPublica = (segundo <> 254)
        Case 172
            AntitrampaIpEsPublica = Not (segundo >= 16 And segundo <= 31)
        Case 192
            AntitrampaIpEsPublica = (segundo <> 168)
        Case Else
            AntitrampaIpEsPublica = True
    End Select
    Exit Function
AntitrampaIpEsPublica_Err:
    AntitrampaIpEsPublica = False
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.AntitrampaIpEsPublica", Erl)
End Function

' Gancho para E4 (guardian): cuando el server conozca un id de PC por conexion, compararlo aca y devolver True
' si coinciden (evento kill_misma_pc). Hoy el server no conoce ningun id de maquina (H8): nunca coincide.
Private Function KillMismaPc(ByVal asesino As Integer, ByVal victima As Integer) As Boolean
    KillMismaPc = False
End Function

' Evento por el que la muerte NO cuenta por mismo origen ("kill_misma_pc" / "kill_misma_ip"), o "" si cuenta.
' No registra nada: ContarMuerte registra, ActStats solo pregunta.
Public Function AntitrampaKillMotivoMismoOrigen(ByVal asesino As Integer, ByVal victima As Integer) As String
    On Error GoTo AntitrampaKillMotivoMismoOrigen_Err
    Dim ipAsesino As String
    If KillMismaPc(asesino, victima) Then
        AntitrampaKillMotivoMismoOrigen = "kill_misma_pc"
        Exit Function
    End If
    If Not IsFeatureEnabled(KILL_TOGGLE_IP) Then Exit Function
    ipAsesino = LCase$(Trim$(UserList(asesino).ConnectionDetails.IP))
    If ipAsesino <> LCase$(Trim$(UserList(victima).ConnectionDetails.IP)) Then Exit Function
    If Not AntitrampaIpEsPublica(ipAsesino) Then Exit Function
    AntitrampaKillMotivoMismoOrigen = "kill_misma_ip"
    Exit Function
AntitrampaKillMotivoMismoOrigen_Err:
    AntitrampaKillMotivoMismoOrigen = vbNullString
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.AntitrampaKillMotivoMismoOrigen", Erl)
End Function

' Una linea en el log antitrampa a nombre del asesino, con la victima en el detalle. Solo log.
Public Sub AntitrampaKillRegistrar(ByVal asesino As Integer, ByVal victima As Integer, ByVal evento As String, ByVal extra As String)
    On Error GoTo AntitrampaKillRegistrar_Err
    Dim detalle As String
    With UserList(victima)
        detalle = "victima_pj=" & .Id & " victima_nick=" & .name & " victima_cuenta=" & .AccountID & " victima_ip=" & .ConnectionDetails.IP
    End With
    If LenB(extra) > 0 Then detalle = detalle & " " & extra
    Call AntitrampaRegistrar(asesino, evento, "log", detalle)
    Exit Sub
AntitrampaKillRegistrar_Err:
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.AntitrampaKillRegistrar", Erl)
End Sub

' Igual para un reto (ModRetos): el rival en el detalle. Solo log.
Public Sub AntitrampaRetoRegistrar(ByVal jugador As Integer, ByVal rival As Integer, ByVal evento As String, ByVal extra As String)
    On Error GoTo AntitrampaRetoRegistrar_Err
    Dim detalle As String
    With UserList(rival)
        detalle = "rival_pj=" & .Id & " rival_nick=" & .name & " rival_cuenta=" & .AccountID & " rival_ip=" & .ConnectionDetails.IP
    End With
    If LenB(extra) > 0 Then detalle = detalle & " " & extra
    Call AntitrampaRegistrar(jugador, evento, "log", detalle)
    Exit Sub
AntitrampaRetoRegistrar_Err:
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.AntitrampaRetoRegistrar", Erl)
End Sub

' Vacia el tope si cambio la fecha desde la ultima consulta (el reset de las 00:00 es perezoso).
Private Sub RevisarDiaKills()
    Dim hoy As String
    hoy = Format$(Date, "yyyymmdd")
    If m_KillsPareja Is Nothing Then Set m_KillsPareja = New Dictionary
    If hoy <> m_KillsFecha Then
        Call m_KillsPareja.RemoveAll
        m_KillsFecha = hoy
    End If
End Sub

Private Function ClavePareja(ByVal asesino As Integer, ByVal victima As Integer) As String
    ClavePareja = CStr(UserList(asesino).Id) & ">" & CStr(UserList(victima).Id)
End Function

' La llama ContarMuerte cuando la muerte ya contaria por todas las demas reglas (FactionReKillTime incluido):
' la suma al tope de la pareja y devuelve True si sigue dentro del tope del dia; si no, registra
' kill_tope_diario y devuelve False.
Public Function AntitrampaKillParejaCuenta(ByVal asesino As Integer, ByVal victima As Integer) As Boolean
    On Error GoTo AntitrampaKillParejaCuenta_Err
    Dim clave As String, cantidad As Long
    If KillsMismaParejaPorDia <= 0 Then
        AntitrampaKillParejaCuenta = True
        Exit Function
    End If
    Call RevisarDiaKills
    clave = ClavePareja(asesino, victima)
    If m_KillsPareja.Exists(clave) Then cantidad = m_KillsPareja.Item(clave)
    If cantidad < CONTADOR_TOPE Then cantidad = cantidad + 1
    m_KillsPareja.Item(clave) = cantidad
    If cantidad <= KillsMismaParejaPorDia Then
        AntitrampaKillParejaCuenta = True
    Else
        Call AntitrampaKillRegistrar(asesino, victima, "kill_tope_diario", "kills_hoy=" & cantidad & " tope=" & KillsMismaParejaPorDia)
    End If
    Exit Function
AntitrampaKillParejaCuenta_Err:
    AntitrampaKillParejaCuenta = True
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.AntitrampaKillParejaCuenta", Erl)
End Function

' True si la pareja ya paso el tope del dia (para ActStats: sin experiencia ni UsuariosMatados). No cuenta.
Public Function AntitrampaKillParejaExcedida(ByVal asesino As Integer, ByVal victima As Integer) As Boolean
    On Error GoTo AntitrampaKillParejaExcedida_Err
    Dim clave As String
    If KillsMismaParejaPorDia <= 0 Then Exit Function
    Call RevisarDiaKills
    clave = ClavePareja(asesino, victima)
    If m_KillsPareja.Exists(clave) Then AntitrampaKillParejaExcedida = (m_KillsPareja.Item(clave) > KillsMismaParejaPorDia)
    Exit Function
AntitrampaKillParejaExcedida_Err:
    AntitrampaKillParejaExcedida = False
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.AntitrampaKillParejaExcedida", Erl)
End Function

Private Function SiNo(ByVal valor As Boolean) As String
    If valor Then
        SiNo = "si"
    Else
        SiNo = "no"
    End If
End Function

' "~ANTITRAMPA kills [nuevo_dia | loopback si|no]" (ver HandleAntitrampaChatCommand).
Private Sub ComandoKills(ByVal UserIndex As Integer, ByRef partes() As String)
    On Error GoTo ComandoKills_Err
    Dim opcion As String, valor As String
    If UBound(partes) >= 2 Then opcion = LCase$(Trim$(partes(2)))
    If UBound(partes) >= 3 Then valor = LCase$(Trim$(partes(3)))
    Call RevisarDiaKills
    If LenB(opcion) > 0 Then
        If (UserList(UserIndex).flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin)) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Antitrampa: solo Dios/Admin pueden cambiar la regla de kills.", e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Select Case opcion
            Case "nuevo_dia"
                Call m_KillsPareja.RemoveAll
                Call LogAudit("ANTITRAMPA", "GM " & UserList(UserIndex).name & " vacio el tope diario de kills (como a las 00:00)")
            Case "loopback"
                Select Case valor
                    Case "si"
                        m_KillIncluirLoopback = True
                    Case "no"
                        m_KillIncluirLoopback = False
                    Case Else
                        Call WriteConsoleMsg(UserIndex, "Uso: ~ANTITRAMPA kills [nuevo_dia | loopback si|no]", e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                End Select
                Call LogAudit("ANTITRAMPA", "GM " & UserList(UserIndex).name & " puso kills loopback=" & SiNo(m_KillIncluirLoopback) & " (hasta el proximo reinicio)")
            Case Else
                Call WriteConsoleMsg(UserIndex, "Uso: ~ANTITRAMPA kills [nuevo_dia | loopback si|no]", e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
        End Select
    End If
    Call WriteConsoleMsg(UserIndex, "Antitrampa kills: regla_ip=" & SiNo(IsFeatureEnabled(KILL_TOGGLE_IP)) & " loopback_cuenta=" & SiNo(m_KillIncluirLoopback) & _
            " ips_tunel=" & CantidadIpsTunel() & " tope_por_pareja=" & KillsMismaParejaPorDia & " parejas_hoy=" & m_KillsPareja.Count, _
            e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
    Exit Sub
ComandoKills_Err:
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.ComandoKills", Erl)
End Sub
