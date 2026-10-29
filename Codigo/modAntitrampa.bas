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
'    single-thread (leccion de LogThis, docs/server/anti-macro-estado.md).
' 2. Anti-speedhack: credito en ms tolerante a rafagas (lag). Cada paquete de caminar suma el tiempo
'    real transcurrido (con tope m_CreditoMaxMs) y cada paso aceptado cuesta
'    IntervaloCaminar / velocidad * m_ToleranciaPct / 100. Sin credito el paso se rechaza.
'    Ya NO depende de PacketRatePolicy.ini [Walk] Limit (que vale 0 y lo dejaba muerto).
' 3. Rechazos de IntervaloPermite*: solo se cuentan; con m_IntervaloRechazosEvento en la ventana
'    queda un evento. No cambia que se rechaza.
' 4. Contadores anti-replay de CastSpell/Drop/EquipItem/Work: solo registro, nunca bloqueo.
' 5. Politica (Configuracion.ini [ANTITRAMPA] Accion, o en caliente con ~ANTITRAMPA):
'    log | avisar_gm | expulsar. Nunca banea. Solo el speedhack aplica la politica; 2-4 solo registran.
'    El anti-macro de verifyTimeStamp NO pasa por aca (decision del duenio 2026-08-08).
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

Private Const INTERVALO_CATEGORIA_MAX As Long = 6
Private Const FLUSH_MS As Long = 1000
Private Const FLUSH_LINEAS As Long = 64
Private Const LOG_ARBOL As String = "Antitrampa"

Private m_Accion As e_AntitrampaAccion
Private m_ToleranciaPct As Long
Private m_CreditoMaxMs As Long
Private m_SpeedRechazosEvento As Long
Private m_SpeedVentanaMs As Long
Private m_IntervaloRechazosEvento As Long
Private m_IntervaloVentanaMs As Long
Private m_ContadorVentanaMs As Long
Private m_AvisoGmCadaMs As Long

Private m_Buffer As String
Private m_BufferLineas As Long
Private m_BufferFecha As String
Private m_UltimoFlush As Long

' Lee la configuracion. La llama General.Main despues de cargar Configuracion.ini.
Public Sub InitAntitrampa()
    On Error GoTo InitAntitrampa_Err
    m_Accion = ParseAccion(SvrConfig.GetValue("AntitrampaAccion"))
    m_ToleranciaPct = SvrConfig.GetValue("AntitrampaSpeedToleranciaPct")
    m_CreditoMaxMs = SvrConfig.GetValue("AntitrampaSpeedCreditoMaxMs")
    m_SpeedRechazosEvento = SvrConfig.GetValue("AntitrampaSpeedRechazosEvento")
    m_SpeedVentanaMs = SvrConfig.GetValue("AntitrampaSpeedVentanaMs")
    m_IntervaloRechazosEvento = SvrConfig.GetValue("AntitrampaIntervaloRechazosEvento")
    m_IntervaloVentanaMs = SvrConfig.GetValue("AntitrampaIntervaloVentanaMs")
    m_ContadorVentanaMs = SvrConfig.GetValue("AntitrampaContadorVentanaMs")
    m_AvisoGmCadaMs = SvrConfig.GetValue("AntitrampaAvisoGmCadaMs")
    If m_SpeedRechazosEvento < 1 Then m_SpeedRechazosEvento = 1
    If m_IntervaloRechazosEvento < 1 Then m_IntervaloRechazosEvento = 1
    m_UltimoFlush = GetTickCountRaw()
    Call LogThis(0, "Antitrampa: accion=" & AccionTexto(m_Accion) & " tolerancia=" & m_ToleranciaPct & "% credito_max_ms=" & m_CreditoMaxMs & _
            " speed_rechazos=" & m_SpeedRechazosEvento & "/" & m_SpeedVentanaMs & "ms intervalo_rechazos=" & m_IntervaloRechazosEvento & "/" & _
            m_IntervaloVentanaMs & "ms", vbLogEventTypeInformation)
    Exit Sub
InitAntitrampa_Err:
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.InitAntitrampa", Erl)
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

Public Sub AntitrampaResetUser(ByVal UserIndex As Integer)
    Dim vacio As t_AntitrampaUser
    UserList(UserIndex).Counters.Antitrampa = vacio
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
    Dim raiz As String, carpeta As String
    raiz = App.Path & "\Logs"
    If LenB(Dir$(raiz, vbDirectory)) = 0 Then MkDir raiz
    carpeta = raiz & "\" & LOG_ARBOL
    If LenB(Dir$(carpeta, vbDirectory)) = 0 Then MkDir carpeta
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
    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()
    With UserList(UserIndex).Counters.Antitrampa
        If Not forzar And .UltimoAvisoGm <> 0 Then
            If TicksElapsed(.UltimoAvisoGm, nowRaw) < m_AvisoGmCadaMs Then
                .EventosSinAvisar = .EventosSinAvisar + 1
                Exit Sub
            End If
        End If
        Dim texto As String
        texto = "Antitrampa: " & UserList(UserIndex).name & " - " & evento & " (" & detalle & ")"
        If .EventosSinAvisar > 0 Then texto = texto & " +" & .EventosSinAvisar & " eventos sin avisar"
        texto = texto & ". Accion: " & AccionTexto(m_Accion) & ". Ver Logs\Antitrampa."
        .UltimoAvisoGm = nowRaw
        If .UltimoAvisoGm = 0 Then .UltimoAvisoGm = 1
        .EventosSinAvisar = 0
    End With
    Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg(texto, e_TextChannel.TEXTCHANNEL_SERVER_STAFF, e_FontTypeNames.FONTTYPE_SERVER))
End Sub

' Registra el evento y aplica la politica. Devuelve True si se cerro la conexion: el llamador tiene que
' salir sin volver a tocar el slot.
Private Function AplicarPolitica(ByVal UserIndex As Integer, ByVal evento As String, ByVal detalle As String) As Boolean
    Dim accion As e_AntitrampaAccion
    accion = m_Accion
    Call AntitrampaRegistrar(UserIndex, evento, AccionTexto(accion), detalle)
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
    Dim velocidad As Single, credito As Double, costo As Double
    With UserList(UserIndex)
        velocidad = .Char.speeding
        ' Recien logueado speeding llega en 0 (nadie recalcula la velocidad en el login) y el cliente
        ' camina a IntervaloCaminar: es la velocidad normal.
        If velocidad <= 0 Then velocidad = VelocidadNormal
        costo = CDbl(IntervaloCaminar) / CDbl(velocidad) * CDbl(m_ToleranciaPct) / 100#
        If .Counters.LastStep = 0 Then
            credito = m_CreditoMaxMs
        Else
            credito = CDbl(.Counters.SpeedHackCounter) + TicksElapsed(.Counters.LastStep, nowRaw)
        End If
        If credito > m_CreditoMaxMs Then credito = m_CreditoMaxMs
        .Counters.SpeedHackCounter = credito
        .Counters.LastStep = nowRaw
        If .Counters.LastStep = 0 Then .Counters.LastStep = 1
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
    Dim detalle As String
    With UserList(UserIndex).Counters.Antitrampa
        If .SpeedVentanaInicio = 0 Or TicksElapsed(.SpeedVentanaInicio, nowRaw) >= m_SpeedVentanaMs Then
            .SpeedVentanaInicio = nowRaw
            If .SpeedVentanaInicio = 0 Then .SpeedVentanaInicio = 1
            .SpeedRechazos = 0
        End If
        .SpeedRechazos = .SpeedRechazos + 1
        If .SpeedRechazos < m_SpeedRechazosEvento Then Exit Function
        detalle = "pasos_rechazados=" & .SpeedRechazos & " en_ms=" & JsonEntero(TicksElapsed(.SpeedVentanaInicio, nowRaw)) & _
                " paso_ms=" & JsonEntero(CDbl(IntervaloCaminar) / IIf(UserList(UserIndex).Char.speeding > 0, UserList(UserIndex).Char.speeding, VelocidadNormal)) & _
                " velocidad=" & Replace$(CStr(UserList(UserIndex).Char.speeding), ",", ".")
        .SpeedRechazos = 0
        .SpeedVentanaInicio = nowRaw
        If .SpeedVentanaInicio = 0 Then .SpeedVentanaInicio = 1
    End With
    AntitrampaPasoRechazado = AplicarPolitica(UserIndex, "speedhack", detalle)
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
            .IntervaloVentanaInicio = nowRaw
            If .IntervaloVentanaInicio = 0 Then .IntervaloVentanaInicio = 1
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
        .IntervaloVentanaInicio = nowRaw
        If .IntervaloVentanaInicio = 0 Then .IntervaloVentanaInicio = 1
    End With
    Call AntitrampaRegistrar(UserIndex, "intervalo_rechazado", "log", detalle)
    Exit Sub
AntitrampaRechazoIntervalo_Err:
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.AntitrampaRechazoIntervalo", Erl)
End Sub

' ---------------------------------------------------------------------------------------------
' Contadores anti-replay de los handlers que antes solo los leian (H6): solo registro
' ---------------------------------------------------------------------------------------------

' Si el contador del cliente no crece (paquete grabado y reenviado) queda `contador_repetido`, como mucho
' una linea por jugador por ventana. Nunca bloquea ni cierra (VerifyPacketSequence si cierra: no se usa aca).
Public Sub AntitrampaContador(ByVal UserIndex As Integer, ByVal slot As PacketNames, ByVal recibido As Long, ByVal paquete As String)
    On Error GoTo AntitrampaContador_Err
    Dim nowRaw As Long, anterior As Long, detalle As String
    With UserList(UserIndex)
        anterior = .PacketCounters(slot)
        .PacketCounters(slot) = recibido
        If recibido > anterior Then Exit Sub
        nowRaw = GetTickCountRaw()
        With .Counters.Antitrampa
            If .ContadorVentanaInicio <> 0 Then
                If TicksElapsed(.ContadorVentanaInicio, nowRaw) < m_ContadorVentanaMs Then
                    .ContadorRepetidos = .ContadorRepetidos + 1
                    Exit Sub
                End If
            End If
            detalle = "paquete=" & paquete & " contador=" & recibido & " anterior=" & anterior
            If .ContadorRepetidos > 0 Then detalle = detalle & " repetidos_sin_registrar=" & .ContadorRepetidos
            .ContadorRepetidos = 0
            .ContadorVentanaInicio = nowRaw
            If .ContadorVentanaInicio = 0 Then .ContadorVentanaInicio = 1
        End With
    End With
    Call AntitrampaRegistrar(UserIndex, "contador_repetido", "log", detalle)
    Exit Sub
AntitrampaContador_Err:
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.AntitrampaContador", Erl)
End Sub

' ---------------------------------------------------------------------------------------------
' Comando de GM "~ANTITRAMPA [log|avisar_gm|expulsar]" (viaja como chat, igual que ~LOGLEVEL)
' ---------------------------------------------------------------------------------------------

Public Function HandleAntitrampaChatCommand(ByVal UserIndex As Integer, ByVal chat As String) As Boolean
    On Error GoTo HandleAntitrampaChatCommand_Err
    Dim partes() As String
    partes = Split(Trim$(chat), " ")
    If UCase$(Trim$(partes(0))) <> "~ANTITRAMPA" Then Exit Function
    HandleAntitrampaChatCommand = True
    If UBound(partes) >= 1 Then
        Select Case LCase$(Trim$(partes(1)))
            Case "log", "avisar_gm", "expulsar"
                m_Accion = ParseAccion(partes(1))
                Call LogAudit("ANTITRAMPA", "GM " & UserList(UserIndex).name & " cambio la accion a " & AccionTexto(m_Accion) & " (hasta el proximo reinicio)")
            Case Else
                Call WriteConsoleMsg(UserIndex, "Uso: ~ANTITRAMPA log|avisar_gm|expulsar", e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
        End Select
    End If
    Call WriteConsoleMsg(UserIndex, "Antitrampa: accion = " & AccionTexto(m_Accion), e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
    Exit Function
HandleAntitrampaChatCommand_Err:
    Call TraceError(Err.Number, Err.Description, "modAntitrampa.HandleAntitrampaChatCommand", Erl)
End Function
