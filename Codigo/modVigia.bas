Attribute VB_Name = "modVigia"
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
' modVigia (plan 2026-10-29-seguridad-antitrampas, etapa E2 "vigia")
'
' Suma al log estructurado de modAntitrampa (Logs\Antitrampa\Antitrampa_AAAA-MM-DD.jsonl, mismo buffer y
' flush) lo que necesita tools\antitrampa\analizar.py. SOLO registra: no avisa a los GMs (las anomalias se
' juntan en el analizador; riesgo "ruido a los GMs" del plan) y no sanciona (duenio: detectar y reportar).
'
' 1. Oro e items con su FUENTE. La diferencia la calculan los ganchos de /denunciar (modDenuncias.DenunciaOro
'    y DenunciaItemSlot, desde Protocol_Writes) y la pasan a VigiaDelta. La fuente la marca el codigo que mueve
'    oro/items con VigiaFuente / VigiaFuenteCon(UserIndex, ...) antes del cambio, y vale solo para esa
'    OPERACION: el paquete que se esta procesando, o la vuelta del loop para lo que corre en timers.
'    VigiaInicioPaquete, VigiaFinPaquetes y VigiaTick la vencen (contador m_Epoca: O(1), sin recorrer usuarios).
'    Las diferencias de una operacion se netean por jugador + objeto + fuente (mover un item de slot da -N y +N
'    = nada). Despues:
'    - ganancia SIN fuente: linea inmediata `fuente=sin_fuente` con el paquete en curso (la senial "oro/item
'      sin fuente" del plan; un dupe por un paquete cualquiera aparece asi);
'    - perdida sin fuente (pociones, flechas, pagos) o de consumo (usar_item, trabajo, quest...): no se
'      registra (no es senial y es la mayoria del volumen); ver FuenteRegistraPerdidas;
'    - gm y comercio_usuario: linea inmediata, con el GM o la contraparte en el detalle;
'    - el resto (npc_comercio, piso, trabajo, banco, quest...): se suma por jugador + objeto + fuente + signo y sale UNA
'      linea cada VigiaAgregadoMs (n = operaciones) y al desloguear.
' 2. login / logout con IP y cuenta. Sin id de maquina hasta E4 (H8). Detras de playit la IP es 127.x (H12):
'    el analizador lo detecta y lo dice en el reporte.
' 3. ritmo: por jugador y tipo (ataque, hechizo, usar, click) cantidad, media y desvio del intervalo entre
'    PAQUETES (llegada, no aceptacion: lo aceptado ya lo cuantiza el server). Los huecos mayores a
'    VigiaRitmoPausaMs cuentan como pausa y no entran. Una linea cada VigiaRitmoVentanaMs si hubo al menos
'    VigiaRitmoMinN intervalos, y al desloguear.
' 4. ruta: pasos aceptados comprimidos (N/S/E/O + repeticiones; "@mapa,x,y" en un salto o cambio de mapa).
'    Una linea cada VigiaRutaVentanaMs o RUTA_LARGO_MAX caracteres, y al desloguear, si hubo al menos
'    VigiaRutaMinPasos pasos.
' LIMITE CONOCIDO (W5): la fuente se marca al ENTRAR a la funcion que mueve oro/items (Comercio, PickObjAt,
' UseInvItem, banco...). Un dupe que pase por adentro de una de esas funciones sale con esa fuente (banco, piso,
' npc_comercio, usar_item), NO como sin_fuente: para eso el analizador mira tambien el volumen por fuente.
' Todo falla abierto: un error aca se registra con TraceError y el juego sigue igual.
Option Explicit

Private Const RITMO_ATAQUE As Long = 0
Private Const RITMO_HECHIZO As Long = 1
Private Const RITMO_USAR As Long = 2
Private Const RITMO_CLICK As Long = 3
Private Const RITMO_MAX As Long = 3

Private Const PEND_MAX As Long = 64
Private Const RUTA_LARGO_MAX As Long = 1500
Private Const EPOCA_TOPE As Long = 2000000000
Private Const CONTADOR_TOPE As Long = 1000000
Private Const CANTIDAD_TOPE As Double = 1000000000#
Private Const FUENTE_SIN As String = "sin_fuente"
Private Const FUENTE_GM As String = "gm"
Private Const FUENTE_COMERCIO_USUARIO As String = "comercio_usuario"

Private Type t_VigiaRitmo
    Ultimo As Long
    Inicio As Long
    Cantidad As Long
    Suma As Double
    SumaCuad As Double
    MinMs As Long
    MaxMs As Long
    Pausas As Long
End Type

Private Type t_VigiaUser
    Fuente As String
    FuenteDetalle As String
    FuenteEpoca As Long
    LoginTick As Long
    ClavesAgregadas As String     ' claves de m_Agregados de este jugador, separadas por vbTab
    Ritmo(0 To RITMO_MAX) As t_VigiaRitmo
    Ruta As String
    RutaInicio As String
    RutaDir As String
    RutaRep As Long
    RutaPasos As Long
    RutaTick As Long
    RutaMapa As Integer
    RutaX As Integer
    RutaY As Integer
End Type

Private Type t_VigiaPend
    UserIndex As Integer
    ObjIndex As Long
    Delta As Double
    Total As Long
    Fuente As String
    FuenteDetalle As String
    PaqueteId As Long             ' -1 = ninguno (timer, u otro jugador que el del paquete)
End Type

Private m_U() As t_VigiaUser
Private m_Max As Long
Private m_Epoca As Long
Private m_PaqueteId As Long          ' paquete que se esta procesando (-1 = ninguno)
Private m_PaqueteUser As Integer     ' de quien es ese paquete

Private m_Pend(1 To PEND_MAX) As t_VigiaPend
Private m_PendCant As Long

Private m_Agregados As Dictionary            ' "UserIndex|ObjIndex|fuente" -> Array(delta As Double, n As Long)
Private m_UltimoAgregado As Long

Private m_AgregadoMs As Long
Private m_RitmoVentanaMs As Long
Private m_RitmoPausaMs As Long
Private m_RitmoMinN As Long
Private m_RutaVentanaMs As Long
Private m_RutaMinPasos As Long

' ---------------------------------------------------------------------------------------------
' Arranque y reset por jugador
' ---------------------------------------------------------------------------------------------

Private Sub CargarDefaults()
    m_AgregadoMs = 60000
    m_RitmoVentanaMs = 300000
    m_RitmoPausaMs = 2000
    m_RitmoMinN = 20
    m_RutaVentanaMs = 300000
    m_RutaMinPasos = 20
End Sub

' La llama General.Main despues de InitDenuncias (MaxUsers ya esta cargado).
Public Sub InitVigia()
    On Error GoTo InitVigia_Err
    Call CargarDefaults
    Set m_Agregados = New Dictionary
    m_Epoca = 1
    m_PendCant = 0
    m_PaqueteId = -1
    m_PaqueteUser = 0
    m_UltimoAgregado = GetTickCountRaw()
    m_Max = 0
    If MaxUsers > 0 Then
        ReDim m_U(1 To MaxUsers)
        m_Max = MaxUsers
    End If
    m_AgregadoMs = SvrConfig.GetValue("VigiaAgregadoMs")
    m_RitmoVentanaMs = SvrConfig.GetValue("VigiaRitmoVentanaMs")
    m_RitmoPausaMs = SvrConfig.GetValue("VigiaRitmoPausaMs")
    m_RitmoMinN = SvrConfig.GetValue("VigiaRitmoMinN")
    m_RutaVentanaMs = SvrConfig.GetValue("VigiaRutaVentanaMs")
    m_RutaMinPasos = SvrConfig.GetValue("VigiaRutaMinPasos")
    If m_AgregadoMs < 1000 Then m_AgregadoMs = 1000
    If m_RitmoVentanaMs < 10000 Then m_RitmoVentanaMs = 10000
    If m_RitmoPausaMs < 100 Then m_RitmoPausaMs = 100
    If m_RitmoMinN < 2 Then m_RitmoMinN = 2
    If m_RutaVentanaMs < 10000 Then m_RutaVentanaMs = 10000
    If m_RutaMinPasos < 1 Then m_RutaMinPasos = 1
    Call LogThis(0, "Vigia: agregado_ms=" & m_AgregadoMs & " ritmo_ventana_ms=" & m_RitmoVentanaMs & " ritmo_pausa_ms=" & m_RitmoPausaMs & _
            " ritmo_min_n=" & m_RitmoMinN & " ruta_ventana_ms=" & m_RutaVentanaMs & " ruta_min_pasos=" & m_RutaMinPasos & " slots=" & m_Max, _
            vbLogEventTypeInformation)
    Exit Sub
InitVigia_Err:
    Call CargarDefaults
    Call TraceError(Err.Number, Err.Description, "modVigia.InitVigia (quedan los defaults)", Erl)
End Sub

Private Function SlotValido(ByVal UserIndex As Integer) As Boolean
    SlotValido = (UserIndex >= 1 And UserIndex <= m_Max)
End Function

Private Function TickNoCero(ByVal nowRaw As Long) As Long
    TickNoCero = nowRaw
    If TickNoCero = 0 Then TickNoCero = 1
End Function

' Campo por campo (como el resto de TCP.ResetContadores), sin asignar el UDT entero. Desde AntitrampaResetUser.
Public Sub VigiaResetUser(ByVal UserIndex As Integer)
    On Error GoTo VigiaResetUser_Err
    Dim i As Long
    If Not SlotValido(UserIndex) Then Exit Sub
    With m_U(UserIndex)
        .Fuente = vbNullString
        .FuenteDetalle = vbNullString
        .FuenteEpoca = 0
        .LoginTick = 0
        .ClavesAgregadas = vbNullString
        For i = 0 To RITMO_MAX
            Call ResetRitmo(UserIndex, i, True)
        Next i
        .Ruta = vbNullString
        .RutaInicio = vbNullString
        .RutaDir = vbNullString
        .RutaRep = 0
        .RutaPasos = 0
        .RutaTick = 0
        .RutaMapa = 0
        .RutaX = 0
        .RutaY = 0
    End With
    Exit Sub
VigiaResetUser_Err:
    Call TraceError(Err.Number, Err.Description, "modVigia.VigiaResetUser", Erl)
End Sub

Private Sub ResetRitmo(ByVal UserIndex As Integer, ByVal tipo As Long, ByVal conUltimo As Boolean)
    With m_U(UserIndex).Ritmo(tipo)
        If conUltimo Then .Ultimo = 0
        .Inicio = 0
        .Cantidad = 0
        .Suma = 0
        .SumaCuad = 0
        .MinMs = 0
        .MaxMs = 0
        .Pausas = 0
    End With
End Sub

' ---------------------------------------------------------------------------------------------
' Operacion actual (vence la fuente) y loop principal
' ---------------------------------------------------------------------------------------------

Private Sub AvanzarEpoca()
    If m_PendCant > 0 Then Call Asentar
    If m_Epoca >= EPOCA_TOPE Then
        m_Epoca = 1
    Else
        m_Epoca = m_Epoca + 1
    End If
End Sub

' Desde Protocol.HandleIncomingData, antes de despachar cada paquete. Solo guarda el id (Long): el nombre se
' resuelve al escribir una linea sin fuente, no en cada paquete (caminar incluido).
Public Sub VigiaInicioPaquete(ByVal UserIndex As Integer, ByVal PacketId As Long)
    On Error GoTo VigiaInicioPaquete_Err
    Dim tipo As Long
    Call AvanzarEpoca
    m_PaqueteId = PacketId
    m_PaqueteUser = UserIndex
    If Not SlotValido(UserIndex) Then Exit Sub
    If Not UserList(UserIndex).flags.UserLogged Then Exit Sub
    Select Case PacketId
        Case ClientPacketID.eAttack
            tipo = RITMO_ATAQUE
        Case ClientPacketID.eCastSpell
            tipo = RITMO_HECHIZO
        Case ClientPacketID.eUseItem, ClientPacketID.eUseItemU
            tipo = RITMO_USAR
        Case ClientPacketID.eLeftClick, ClientPacketID.eDoubleClick, ClientPacketID.eWorkLeftClick
            tipo = RITMO_CLICK
        Case Else
            Exit Sub
    End Select
    Call RegistrarRitmo(UserIndex, tipo)
    Exit Sub
VigiaInicioPaquete_Err:
    Call TraceError(Err.Number, Err.Description, "modVigia.VigiaInicioPaquete", Erl)
End Sub

' Desde General.Main despues de procesar la red: lo que siga en esta vuelta (timers) no hereda la fuente.
Public Sub VigiaFinPaquetes()
    On Error GoTo VigiaFinPaquetes_Err
    Call AvanzarEpoca
    m_PaqueteId = -1
    m_PaqueteUser = 0
    Exit Sub
VigiaFinPaquetes_Err:
    Call TraceError(Err.Number, Err.Description, "modVigia.VigiaFinPaquetes", Erl)
End Sub

' Una vez por vuelta del loop principal (despues de los timers).
Public Sub VigiaTick()
    On Error GoTo VigiaTick_Err
    Call AvanzarEpoca
    m_PaqueteId = -1
    m_PaqueteUser = 0
    If TicksElapsed(m_UltimoAgregado, GetTickCountRaw()) >= m_AgregadoMs Then
        m_UltimoAgregado = GetTickCountRaw()
        Call FlushAgregados(0)
    End If
    Exit Sub
VigiaTick_Err:
    Call TraceError(Err.Number, Err.Description, "modVigia.VigiaTick", Erl)
End Sub

' ---------------------------------------------------------------------------------------------
' Fuentes de oro e items
' ---------------------------------------------------------------------------------------------

' Marca de donde sale el proximo cambio de oro/items de UserIndex dentro de esta operacion.
Public Sub VigiaFuente(ByVal UserIndex As Integer, ByVal fuente As String)
    Call VigiaFuenteCon(UserIndex, fuente, vbNullString)
End Sub

' Igual, con detalle (el GM que lo hizo, la contraparte del comercio).
Public Sub VigiaFuenteCon(ByVal UserIndex As Integer, ByVal fuente As String, ByVal detalle As String)
    On Error GoTo VigiaFuenteCon_Err
    If Not SlotValido(UserIndex) Then Exit Sub
    With m_U(UserIndex)
        .Fuente = fuente
        .FuenteDetalle = detalle
        .FuenteEpoca = m_Epoca
    End With
    Exit Sub
VigiaFuenteCon_Err:
    Call TraceError(Err.Number, Err.Description, "modVigia.VigiaFuenteCon", Erl)
End Sub

' "con_pj=.. con_nick=.. con_cuenta=.." de otro jugador (contraparte de un comercio, GM que da algo).
Public Function VigiaContraparte(ByVal prefijo As String, ByVal OtroIndex As Integer) As String
    On Error GoTo VigiaContraparte_Err
    If OtroIndex < 1 Or OtroIndex > m_Max Then Exit Function
    With UserList(OtroIndex)
        VigiaContraparte = prefijo & "_pj=" & .Id & " " & prefijo & "_nick=" & Replace$(.name, " ", "+") & " " & prefijo & "_cuenta=" & .AccountID
    End With
    Exit Function
VigiaContraparte_Err:
    Call TraceError(Err.Number, Err.Description, "modVigia.VigiaContraparte", Erl)
End Function

' Desde modDenuncias (DenunciaOro: ObjIndex = 0; DenunciaItemSlot: el objeto). total = oro despues del cambio.
Public Sub VigiaDelta(ByVal UserIndex As Integer, ByVal ObjIndex As Long, ByVal delta As Double, ByVal total As Long)
    On Error GoTo VigiaDelta_Err
    Dim fuente As String, detalle As String, i As Long
    If Not SlotValido(UserIndex) Then Exit Sub
    If delta = 0 Then Exit Sub
    With m_U(UserIndex)
        If .FuenteEpoca = m_Epoca And LenB(.Fuente) > 0 Then
            fuente = .Fuente
            detalle = .FuenteDetalle
        End If
    End With
    For i = 1 To m_PendCant
        With m_Pend(i)
            If .UserIndex = UserIndex And .ObjIndex = ObjIndex And .Fuente = fuente Then
                .Delta = .Delta + delta
                .Total = total
                Exit Sub
            End If
        End With
    Next i
    If m_PendCant >= PEND_MAX Then Call Asentar
    m_PendCant = m_PendCant + 1
    With m_Pend(m_PendCant)
        .UserIndex = UserIndex
        .ObjIndex = ObjIndex
        .Delta = delta
        .Total = total
        .Fuente = fuente
        .FuenteDetalle = detalle
        ' El paquete solo sirve de evidencia si es del mismo jugador: lo que le pasa a otro en esta operacion
        ' (un connect/disconnect que procesa modNetwork.Tick entre paquetes) no es "por" ese paquete.
        If m_PaqueteUser = UserIndex Then
            .PaqueteId = m_PaqueteId
        Else
            .PaqueteId = -1
        End If
    End With
    Exit Sub
VigiaDelta_Err:
    Call TraceError(Err.Number, Err.Description, "modVigia.VigiaDelta", Erl)
End Sub

Private Function NumeroTexto(ByVal valor As Double) As String
    If valor > CANTIDAD_TOPE Then valor = CANTIDAD_TOPE
    If valor < -CANTIDAD_TOPE Then valor = -CANTIDAD_TOPE
    NumeroTexto = CStr(Fix(valor))
End Function

Private Function CantidadEnInventario(ByVal UserIndex As Integer, ByVal ObjIndex As Long) As Double
    Dim i As Long
    With UserList(UserIndex)
        For i = 1 To .CurrentInventorySlots
            If .invent.Object(i).ObjIndex = ObjIndex Then CantidadEnInventario = CantidadEnInventario + .invent.Object(i).amount
        Next i
    End With
End Function

Private Function DetalleDelta(ByVal objIdx As Long, ByVal fuente As String, ByVal delta As Double, ByVal total As Double, _
                              ByVal operaciones As Long) As String
    Dim texto As String
    texto = "fuente=" & fuente
    If objIdx <> 0 Then texto = texto & " obj=" & objIdx
    texto = texto & " delta=" & NumeroTexto(delta) & " total=" & NumeroTexto(total) & " n=" & operaciones
    DetalleDelta = texto
End Function

' Fuentes donde una perdida es un traspaso (sirve de evidencia: a quien/donde fue la riqueza). En las demas
' (usar_item, trabajo, quest...) perder es consumir y no se registra.
Private Function FuenteRegistraPerdidas(ByVal fuente As String) As Boolean
    Select Case fuente
        Case "piso", "banco", "npc_comercio", "comercio_usuario", "gm", "subasta", "reto", "robo"
            FuenteRegistraPerdidas = True
    End Select
End Function

' Netea y registra las diferencias de la operacion que termino.
Private Sub Asentar()
    On Error GoTo Asentar_Err
    Dim i As Long, cantidad As Long
    cantidad = m_PendCant
    m_PendCant = 0
    For i = 1 To cantidad
        Call EmitirPendiente(i)
        m_Pend(i).Fuente = vbNullString
        m_Pend(i).FuenteDetalle = vbNullString
        m_Pend(i).PaqueteId = -1
    Next i
    Exit Sub
Asentar_Err:
    m_PendCant = 0
    Call TraceError(Err.Number, Err.Description, "modVigia.Asentar", Erl)
End Sub

' Con su propio manejador: una entrada que falla se registra y Asentar sigue con las demas.
Private Sub EmitirPendiente(ByVal i As Long)
    On Error GoTo EmitirPendiente_Err
    Dim evento As String, detalle As String, total As Double, paquete As String
    With m_Pend(i)
        If .Delta = 0 Then Exit Sub
        If Not SlotValido(.UserIndex) Then Exit Sub
        If Not UserList(.UserIndex).flags.UserLogged Then Exit Sub
        If .ObjIndex = 0 Then
            evento = "oro"
            total = .Total
        Else
            evento = "item"
            total = CantidadEnInventario(.UserIndex, .ObjIndex)
        End If
        If LenB(.Fuente) = 0 Then
            ' Perder sin fuente es consumir (pociones, flechas, pagos): no es senial.
            If .Delta < 0 Then Exit Sub
            If .PaqueteId >= 0 Then paquete = GetClientPacketName(.PaqueteId)
            If LenB(paquete) = 0 Then paquete = "ninguno"
            detalle = DetalleDelta(.ObjIndex, FUENTE_SIN, .Delta, total, 1) & " paquete=" & paquete
            Call AntitrampaLinea(.UserIndex, evento, detalle)
        ElseIf .Delta < 0 And Not FuenteRegistraPerdidas(.Fuente) Then
            ' Usar un item, gastar materiales al trabajar, etc.: consumo, no movimiento de riqueza.
            Exit Sub
        ElseIf .Fuente = FUENTE_GM Or .Fuente = FUENTE_COMERCIO_USUARIO Then
            detalle = DetalleDelta(.ObjIndex, .Fuente, .Delta, total, 1)
            If LenB(.FuenteDetalle) > 0 Then detalle = detalle & " " & .FuenteDetalle
            Call AntitrampaLinea(.UserIndex, evento, detalle)
        Else
            Call Agregar(.UserIndex, .ObjIndex, .Fuente, .Delta)
        End If
    End With
    Exit Sub
EmitirPendiente_Err:
    Call TraceError(Err.Number, Err.Description, "modVigia.EmitirPendiente (entrada " & i & ")", Erl)
End Sub

Private Sub Agregar(ByVal UserIndex As Integer, ByVal ObjIndex As Long, ByVal fuente As String, ByVal delta As Double)
    Dim clave As String, valor As Variant
    ' Ganancias y perdidas por separado: depositar y retirar del banco (o tirar y juntar) en la misma ventana no se
    ' anulan entre si; los dos movimientos son evidencia.
    If delta > 0 Then
        clave = CStr(UserIndex) & "|" & CStr(ObjIndex) & "|" & fuente & "|+"
    Else
        clave = CStr(UserIndex) & "|" & CStr(ObjIndex) & "|" & fuente & "|-"
    End If
    If m_Agregados.Exists(clave) Then
        valor = m_Agregados.Item(clave)
        valor(0) = valor(0) + delta
        If valor(1) < CONTADOR_TOPE Then valor(1) = valor(1) + 1
        m_Agregados.Item(clave) = valor
    Else
        m_Agregados.Add clave, Array(delta, 1&)
        m_U(UserIndex).ClavesAgregadas = m_U(UserIndex).ClavesAgregadas & clave & vbTab
    End If
End Sub

' soloUser = 0: todos (cada VigiaAgregadoMs); si no, solo los de ese jugador (al desloguear), por su indice.
Private Sub FlushAgregados(ByVal soloUser As Integer)
    On Error GoTo FlushAgregados_Err
    Dim claves As Variant, i As Long
    If m_Agregados Is Nothing Then Exit Sub
    If m_Agregados.Count = 0 Then Exit Sub
    If soloUser = 0 Then
        claves = m_Agregados.Keys
        For i = 1 To m_Max
            m_U(i).ClavesAgregadas = vbNullString
        Next i
    Else
        If Not SlotValido(soloUser) Then Exit Sub
        If LenB(m_U(soloUser).ClavesAgregadas) = 0 Then Exit Sub
        claves = Split(m_U(soloUser).ClavesAgregadas, vbTab)
        m_U(soloUser).ClavesAgregadas = vbNullString
    End If
    For i = 0 To UBound(claves)
        If LenB(claves(i)) > 0 Then Call EmitirAgregado(CStr(claves(i)))
    Next i
    Exit Sub
FlushAgregados_Err:
    Call TraceError(Err.Number, Err.Description, "modVigia.FlushAgregados", Erl)
End Sub

' Una clave de m_Agregados -> una linea (y se borra). Con su manejador: una que falla no corta a las demas.
Private Sub EmitirAgregado(ByVal clave As String)
    On Error GoTo EmitirAgregado_Err
    Dim partes() As String, UserIndex As Integer, objIdx As Long, valor As Variant, evento As String, total As Double
    If Not m_Agregados.Exists(clave) Then Exit Sub
    valor = m_Agregados.Item(clave)
    Call m_Agregados.Remove(clave)
    partes = Split(clave, "|")
    UserIndex = CInt(partes(0))
    objIdx = CLng(partes(1))
    If valor(0) = 0 Or Not SlotValido(UserIndex) Then Exit Sub
    If Not UserList(UserIndex).flags.UserLogged Then Exit Sub
    If objIdx = 0 Then
        evento = "oro"
        total = UserList(UserIndex).Stats.GLD
    Else
        evento = "item"
        total = CantidadEnInventario(UserIndex, objIdx)
    End If
    Call AntitrampaLinea(UserIndex, evento, DetalleDelta(objIdx, partes(2), valor(0), total, valor(1)))
    Exit Sub
EmitirAgregado_Err:
    Call TraceError(Err.Number, Err.Description, "modVigia.EmitirAgregado", Erl)
End Sub

' ---------------------------------------------------------------------------------------------
' Login / logout
' ---------------------------------------------------------------------------------------------

Private Function SiNo(ByVal valor As Boolean) As String
    If valor Then
        SiNo = "si"
    Else
        SiNo = "no"
    End If
End Function

' Desde ConnectUser_Complete, con el personaje ya logueado.
Public Sub VigiaLogin(ByVal UserIndex As Integer)
    On Error GoTo VigiaLogin_Err
    Dim i As Long
    If Not SlotValido(UserIndex) Then Exit Sub
    m_U(UserIndex).LoginTick = TickNoCero(GetTickCountRaw())
    For i = 0 To RITMO_MAX
        Call ResetRitmo(UserIndex, i, True)
    Next i
    Call ReiniciarRuta(UserIndex, True)
    ' Sin id de maquina hasta E4 (H8): solo IP y cuenta (van en los campos fijos de la linea).
    Call AntitrampaLinea(UserIndex, "login", "gm=" & SiNo(EsGM(UserIndex)) & " oro=" & UserList(UserIndex).Stats.GLD)
    Exit Sub
VigiaLogin_Err:
    Call TraceError(Err.Number, Err.Description, "modVigia.VigiaLogin", Erl)
End Sub

' Desde TCP.CloseUser, ANTES de bajar UserLogged: baja lo pendiente del jugador y registra el logout.
Public Sub VigiaLogout(ByVal UserIndex As Integer)
    On Error GoTo VigiaLogout_Err
    Dim i As Long, sesionMs As Double
    If Not SlotValido(UserIndex) Then Exit Sub
    If Not UserList(UserIndex).flags.UserLogged Then Exit Sub
    If m_PendCant > 0 Then Call Asentar
    Call FlushAgregados(UserIndex)
    For i = 0 To RITMO_MAX
        Call EmitirRitmo(UserIndex, i)
        Call ResetRitmo(UserIndex, i, True)
    Next i
    Call EmitirRuta(UserIndex)
    If m_U(UserIndex).LoginTick <> 0 Then sesionMs = TicksElapsed(m_U(UserIndex).LoginTick, GetTickCountRaw())
    Call AntitrampaLinea(UserIndex, "logout", "sesion_s=" & NumeroTexto(sesionMs / 1000#) & " oro=" & UserList(UserIndex).Stats.GLD)
    m_U(UserIndex).LoginTick = 0
    Exit Sub
VigiaLogout_Err:
    Call TraceError(Err.Number, Err.Description, "modVigia.VigiaLogout", Erl)
End Sub

' ---------------------------------------------------------------------------------------------
' Ritmo de clicks / ataques (CA5 c: intervalo constante +-5 ms por 5 min lo decide el analizador)
' ---------------------------------------------------------------------------------------------

Private Function NombreRitmo(ByVal tipo As Long) As String
    Select Case tipo
        Case RITMO_ATAQUE: NombreRitmo = "ataque"
        Case RITMO_HECHIZO: NombreRitmo = "hechizo"
        Case RITMO_USAR: NombreRitmo = "usar"
        Case Else: NombreRitmo = "click"
    End Select
End Function

Private Function DecimalTexto(ByVal valor As Double) As String
    DecimalTexto = Replace$(Format$(valor, "0.0"), ",", ".")
End Function

Private Sub RegistrarRitmo(ByVal UserIndex As Integer, ByVal tipo As Long)
    Dim nowRaw As Long, dt As Double
    nowRaw = GetTickCountRaw()
    With m_U(UserIndex).Ritmo(tipo)
        If .Inicio = 0 Then .Inicio = TickNoCero(nowRaw)
        If .Ultimo <> 0 Then
            dt = TicksElapsed(.Ultimo, nowRaw)
            If dt > m_RitmoPausaMs Then
                If .Pausas < CONTADOR_TOPE Then .Pausas = .Pausas + 1
            ElseIf .Cantidad < CONTADOR_TOPE Then
                .Cantidad = .Cantidad + 1
                .Suma = .Suma + dt
                .SumaCuad = .SumaCuad + dt * dt
                If .Cantidad = 1 Or dt < .MinMs Then .MinMs = dt
                If dt > .MaxMs Then .MaxMs = dt
            End If
        End If
        .Ultimo = TickNoCero(nowRaw)
        If TicksElapsed(.Inicio, nowRaw) < m_RitmoVentanaMs Then Exit Sub
    End With
    Call EmitirRitmo(UserIndex, tipo)
    Call ResetRitmo(UserIndex, tipo, False)
    m_U(UserIndex).Ritmo(tipo).Inicio = TickNoCero(nowRaw)
End Sub

Private Sub EmitirRitmo(ByVal UserIndex As Integer, ByVal tipo As Long)
    Dim media As Double, varianza As Double, ventana As Double
    With m_U(UserIndex).Ritmo(tipo)
        If .Cantidad < m_RitmoMinN Then Exit Sub
        media = .Suma / .Cantidad
        varianza = .SumaCuad / .Cantidad - media * media
        If varianza < 0 Then varianza = 0
        If .Inicio <> 0 Then ventana = TicksElapsed(.Inicio, GetTickCountRaw())
        Call AntitrampaLinea(UserIndex, "ritmo", "tipo=" & NombreRitmo(tipo) & " n=" & .Cantidad & " media_ms=" & DecimalTexto(media) & _
                " desvio_ms=" & DecimalTexto(Sqr(varianza)) & " min_ms=" & .MinMs & " max_ms=" & .MaxMs & " pausas=" & .Pausas & _
                " ventana_ms=" & NumeroTexto(ventana))
    End With
End Sub

' ---------------------------------------------------------------------------------------------
' Ruta (rutas repetidas de bot)
' ---------------------------------------------------------------------------------------------

Private Sub CerrarTramo(ByVal UserIndex As Integer)
    With m_U(UserIndex)
        If .RutaRep > 0 Then .Ruta = .Ruta & .RutaDir & .RutaRep
        .RutaDir = vbNullString
        .RutaRep = 0
    End With
End Sub

' Empieza una ventana nueva desde la posicion actual (conPosicion: tambien olvida la ultima posicion).
Private Sub ReiniciarRuta(ByVal UserIndex As Integer, ByVal conPosicion As Boolean)
    With m_U(UserIndex)
        .Ruta = vbNullString
        .RutaDir = vbNullString
        .RutaRep = 0
        .RutaPasos = 0
        .RutaTick = TickNoCero(GetTickCountRaw())
        If conPosicion Then
            .RutaMapa = 0
            .RutaInicio = vbNullString
        Else
            .RutaInicio = .RutaMapa & "," & .RutaX & "," & .RutaY
        End If
    End With
End Sub

' Desde Protocol.HandleWalk, con el paso ya aceptado (UserList().pos es la posicion nueva).
Public Sub VigiaPaso(ByVal UserIndex As Integer)
    On Error GoTo VigiaPaso_Err
    Dim mapa As Integer, x As Long, y As Long, dx As Long, dy As Long, direccion As String
    If Not SlotValido(UserIndex) Then Exit Sub
    If Not UserList(UserIndex).flags.UserLogged Then Exit Sub
    mapa = UserList(UserIndex).pos.Map
    x = UserList(UserIndex).pos.x
    y = UserList(UserIndex).pos.y
    With m_U(UserIndex)
        If .RutaMapa = 0 Then
            .RutaMapa = mapa
            .RutaX = x
            .RutaY = y
            .RutaInicio = mapa & "," & x & "," & y
            .RutaTick = TickNoCero(GetTickCountRaw())
            Exit Sub
        End If
        dx = x - .RutaX
        dy = y - .RutaY
        If mapa = .RutaMapa And Abs(dx) + Abs(dy) = 1 Then
            If dy = -1 Then
                direccion = "N"
            ElseIf dy = 1 Then
                direccion = "S"
            ElseIf dx = 1 Then
                direccion = "E"
            Else
                direccion = "O"
            End If
            If direccion = .RutaDir Then
                If .RutaRep < CONTADOR_TOPE Then .RutaRep = .RutaRep + 1
            Else
                Call CerrarTramo(UserIndex)
                .RutaDir = direccion
                .RutaRep = 1
            End If
            If .RutaPasos < CONTADOR_TOPE Then .RutaPasos = .RutaPasos + 1
        ElseIf mapa <> .RutaMapa Or dx <> 0 Or dy <> 0 Then
            Call CerrarTramo(UserIndex)
            .Ruta = .Ruta & "@" & mapa & "," & x & "," & y
        End If
        .RutaMapa = mapa
        .RutaX = x
        .RutaY = y
        If TicksElapsed(.RutaTick, GetTickCountRaw()) < m_RutaVentanaMs And Len(.Ruta) < RUTA_LARGO_MAX Then Exit Sub
    End With
    Call EmitirRuta(UserIndex)
    Exit Sub
VigiaPaso_Err:
    Call TraceError(Err.Number, Err.Description, "modVigia.VigiaPaso", Erl)
End Sub

Private Sub EmitirRuta(ByVal UserIndex As Integer)
    Call CerrarTramo(UserIndex)
    With m_U(UserIndex)
        If .RutaPasos >= m_RutaMinPasos And LenB(.Ruta) > 0 Then
            Call AntitrampaLinea(UserIndex, "ruta", "inicio=" & .RutaInicio & " pasos=" & .RutaPasos & " ms=" & _
                    NumeroTexto(TicksElapsed(.RutaTick, GetTickCountRaw())) & " ruta=" & .Ruta)
        End If
    End With
    If m_U(UserIndex).RutaMapa <> 0 Then
        Call ReiniciarRuta(UserIndex, False)
    Else
        Call ReiniciarRuta(UserIndex, True)
    End If
End Sub
