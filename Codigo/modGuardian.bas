Attribute VB_Name = "modGuardian"
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
' modGuardian (plan 2026-10-29-seguridad-antitrampas, etapa E4b; diseno en ia/work/.../e4-diseno.md)
'
' Lado server del guardian (proceso nativo que corre junto al cliente). Este modulo:
' 1. Entrega el desafio N (16 bytes de RNG_NonceData) a TODO login, despues de WriteLoggedMessage (el login vivo
'    PYMMO=0 no manda version, y un cliente viejo ignora el ServerPacketID desconocido). Un cliente que no contesta
'    queda como `guardian_ausente`.
' 2. Valida cada eGuardianHeartbeat en el orden del diseno (4): rate limit 1 cada GuardianRateLimitMs ANTES de mirar
'    el contenido; logueado + desafio emitido; largo <= 512 y charset [0-9a-z|,._-]; recien ahi kid, HMAC
'    (comparacion XOR acumulada), frescura de `ms`, contador por generacion `g`, flags.
' 3. Registra todo en el log antitrampa (modAntitrampa.AntitrampaLineaAccion). Politica GuardianAccion: log (default) o
'    avisar_gm. Nunca expulsa ni banea en E4b.
' 4. Ausencia: desafio emitido y sin heartbeat valido a los GuardianFactorAusencia x intervalo (contados desde el ultimo
'    valido, o desde el desafio) -> `guardian_ausente`, repetido cada GuardianRepetirAusenteSeg.
' 5. Contadores por sesion hb_validos / hb_invalidos / ultimo_ok en `guardian_resumen` (logout y cada GuardianResumenSeg).
'
' Secretos: Guardian.secret.ini (junto a Configuracion.ini, gitignoreado, nunca en paquetes ni docs):
'   [GUARDIAN]
'   PermitirDev=1          solo para pruebas: sin esto el kid 0 (secreto de prueba PUBLICO) se rechaza
'   SaltLog=<texto>        sal con la que se re-hashea el id de PC antes de loguearlo
'   Kid0=<64 hex>          hasta 2 entradas KidN=<secreto de 32 bytes en hex>
' Archivo ausente o sin la clave del kid = ese kid se rechaza (todo heartbeat invalido) y el server sigue andando.
'
' Frescura (interpretacion de 6.3): `ms` es el reloj monotonico del guardian desde su HELLO. Para una generacion
' ya conocida: |ms - (ms_del_primer_hb_de_g + tiempo_desde_ese_hb)| <= GuardianFrescuraMs. Para el primer hb de una
' generacion no hay ancla propia: solo se exige que el HELLO no sea anterior al desafio, o sea ms <= tiempo desde el
' desafio + GuardianFrescuraMs (el HELLO puede llegar tarde: pipe que reintenta cada 30 s, asi que no hay cota inferior).
' Todo falla abierto: un error en una entrada se registra con TraceError y el juego sigue igual.
Option Explicit

Private Const KS_LABEL As String = "ao20-guard-v1"
Private Const SECRET_FILE As String = "Guardian.secret.ini"
Private Const SECCION_CFG As String = "ANTITRAMPA"
Private Const SECCION_SECRETO As String = "GUARDIAN"
Private Const MAX_LINE As Long = 512
Private Const MAX_GEN As Long = 8
Private Const MAX_KIDS As Long = 2
Private Const TICK_MS As Long = 1000
Private Const CONTADOR_TOPE As Long = 1000000
Private Const AVISO_GM_CADA_MS As Long = 60000

' Indice de t_GuardianUser.EventoTick: eventos con rate limit por jugador.
Private Const EV_FLOOD As Long = 0
Private Const EV_SIN_DESAFIO As Long = 1
Private Const EV_REINICIO As Long = 2
Private Const EV_EXCESIVO As Long = 3
Private Const EV_INVALIDA As Long = 4
Private Const EV_MAX As Long = 4

Private Type t_GuardianGen
    G As String
    UltimoCtr As Double
    AnclaTick As Long
    AnclaMs As Double
End Type

Private Type t_GuardianKid
    Id As Long
    Secreto(0 To 31) As Byte
End Type

Private Type t_GuardianUser
    Activo As Boolean
    Desafiado As Boolean
    CharId As Long
    N(0 To 15) As Byte
    DesafioTick As Long
    UltimoHbTick As Long
    UltimoOkTick As Long
    AusenteUltTick As Long
    ResumenTick As Long
    AvisoTick As Long
    HbValidos As Long
    HbInvalidos As Long
    MidPc As String
    FlagsVistos As Long
    GenCant As Long
    Gen(1 To MAX_GEN) As t_GuardianGen
    EventoTick(0 To EV_MAX) As Long
End Type

Private m_U() As t_GuardianUser
Private m_Max As Long
Private m_Kids(1 To MAX_KIDS) As t_GuardianKid
Private m_KidCant As Long
Private m_PermitirDev As Boolean
Private m_SaltLog As String

Private m_Habilitado As Boolean
Private m_IntervaloSeg As Long
Private m_FactorAusencia As Long
Private m_RepetirAusenteSeg As Long
Private m_ResumenSeg As Long
Private m_RateLimitMs As Long
Private m_FrescuraMs As Long
Private m_EventoCadaMs As Long
Private m_GraciaSeg As Long
Private m_AvisarGm As Boolean
Private m_UltimoTick As Long

' ---------------------------------------------------------------------------------------------
' Arranque
' ---------------------------------------------------------------------------------------------

Private Sub CargarDefaults()
    m_Habilitado = True
    m_IntervaloSeg = 20
    m_FactorAusencia = 3
    m_RepetirAusenteSeg = 300
    m_ResumenSeg = 300
    m_RateLimitMs = 5000
    m_FrescuraMs = 45000
    m_EventoCadaMs = 30000
    m_GraciaSeg = 90
    m_AvisarGm = False
End Sub

' La llama General.Main despues de InitVigia (MaxUsers ya esta cargado).
Public Sub InitGuardian()
    On Error GoTo InitGuardian_Err
    Dim ini As clsIniManager, topeRate As Long
    Call CargarDefaults
    m_UltimoTick = GetTickCountRaw()
    m_Max = 0
    If MaxUsers > 0 Then
        ReDim m_U(1 To MaxUsers)
        m_Max = MaxUsers
    End If
    Set ini = New clsIniManager
    Call ini.Initialize(IniPath & "Configuracion.ini")
    m_Habilitado = (val(ini.GetValue(SECCION_CFG, "GuardianHabilitado", "1")) <> 0)
    m_IntervaloSeg = CLng(val(ini.GetValue(SECCION_CFG, "GuardianIntervaloSeg", "20")))
    m_FactorAusencia = CLng(val(ini.GetValue(SECCION_CFG, "GuardianFactorAusencia", "3")))
    m_RepetirAusenteSeg = CLng(val(ini.GetValue(SECCION_CFG, "GuardianRepetirAusenteSeg", "300")))
    m_ResumenSeg = CLng(val(ini.GetValue(SECCION_CFG, "GuardianResumenSeg", "300")))
    m_RateLimitMs = CLng(val(ini.GetValue(SECCION_CFG, "GuardianRateLimitMs", "5000")))
    m_FrescuraMs = CLng(val(ini.GetValue(SECCION_CFG, "GuardianFrescuraMs", "45000")))
    m_EventoCadaMs = CLng(val(ini.GetValue(SECCION_CFG, "GuardianEventoCadaMs", "30000")))
    m_GraciaSeg = CLng(val(ini.GetValue(SECCION_CFG, "GuardianGraciaSeg", "90")))
    m_AvisarGm = (LCase$(Trim$(ini.GetValue(SECCION_CFG, "GuardianAccion", "log"))) = "avisar_gm")
    ' Piso de cada valor: el guardian clampea el intervalo a 5 s; un 0 pararia la ausencia o el rate limit.
    If m_IntervaloSeg < 5 Then m_IntervaloSeg = 5
    If m_IntervaloSeg > 3600 Then m_IntervaloSeg = 3600
    If m_FactorAusencia < 1 Then m_FactorAusencia = 1
    If m_RepetirAusenteSeg < 5 Then m_RepetirAusenteSeg = 5
    If m_ResumenSeg < 5 Then m_ResumenSeg = 5
    If m_RateLimitMs < 100 Then m_RateLimitMs = 100
    If m_FrescuraMs < 1000 Then m_FrescuraMs = 1000
    If m_EventoCadaMs < 1000 Then m_EventoCadaMs = 1000
    If m_GraciaSeg < 0 Then m_GraciaSeg = 0
    ' W1: el guardian firma una linea cada intervalo/2: el rate limit nunca puede ser mayor que eso (con margen).
    topeRate = ((m_IntervaloSeg * 1000) \ 2) - 500
    If topeRate < 1000 Then topeRate = 1000
    If m_RateLimitMs > topeRate Then m_RateLimitMs = topeRate
    Call CargarSecreto
    ' Nunca se loguea ningun secreto: solo cuantas claves hay y si el kid de prueba esta permitido.
    Call LogThis(0, "Guardian: habilitado=" & m_Habilitado & " intervalo_s=" & m_IntervaloSeg & _
            " factor_ausencia=" & m_FactorAusencia & " repetir_ausente_s=" & m_RepetirAusenteSeg & " resumen_s=" & m_ResumenSeg & _
            " gracia_s=" & m_GraciaSeg & " rate_limit_ms=" & m_RateLimitMs & " frescura_ms=" & m_FrescuraMs & " accion=" & GuardianAccionTexto() & " kids=" & m_KidCant & _
            " permitir_dev=" & m_PermitirDev & " slots=" & m_Max, vbLogEventTypeInformation)
    Exit Sub
InitGuardian_Err:
    Call CargarDefaults
    m_KidCant = 0
    m_PermitirDev = False
    Call TraceError(Err.Number, Err.Description, "modGuardian.InitGuardian (quedan los defaults, sin claves)", Erl)
End Sub

Private Sub CargarSecreto()
    On Error GoTo CargarSecreto_Err
    Dim ruta As String, ini As clsIniManager, k As Long, texto As String, bytes() As Byte, i As Long
    m_KidCant = 0
    m_PermitirDev = False
    m_SaltLog = vbNullString
    ruta = IniPath & SECRET_FILE
    If Not FileExist(ruta) Then
        Call LogThis(0, "Guardian: no existe " & SECRET_FILE & ": todo heartbeat se va a rechazar", vbLogEventTypeWarning)
        Exit Sub
    End If
    Set ini = New clsIniManager
    Call ini.Initialize(ruta)
    m_PermitirDev = (val(ini.GetValue(SECCION_SECRETO, "PermitirDev", "0")) = 1)
    m_SaltLog = Trim$(ini.GetValue(SECCION_SECRETO, "SaltLog", vbNullString))
    For k = 0 To 255
        texto = vbNullString
        If k > 0 Or m_PermitirDev Then texto = LCase$(Trim$(ini.GetValue(SECCION_SECRETO, "Kid" & CStr(k), vbNullString)))
        If LenB(texto) > 0 Then
            If m_KidCant >= MAX_KIDS Then
                Call LogThis(0, "Guardian: se ignora Kid" & k & " (maximo " & MAX_KIDS & " claves)", vbLogEventTypeWarning)
            ElseIf Not EsHex(texto, 64) Then
                Call LogThis(0, "Guardian: Kid" & k & " no son 64 hex: se ignora", vbLogEventTypeWarning)
            Else
                bytes = cnvBytesFromHexStr(texto)
                If cnvBytesLen(bytes) = 32 Then
                    m_KidCant = m_KidCant + 1
                    m_Kids(m_KidCant).Id = k
                    For i = 0 To 31
                        m_Kids(m_KidCant).Secreto(i) = bytes(i)
                    Next i
                End If
            End If
        End If
    Next k
    If LenB(m_SaltLog) = 0 Then Call LogThis(0, "Guardian: SaltLog vacia (el id de PC se loguea sin sal)", vbLogEventTypeWarning)
    Exit Sub
CargarSecreto_Err:
    m_KidCant = 0
    m_PermitirDev = False
    Call TraceError(Err.Number, Err.Description, "modGuardian.CargarSecreto (sin claves)", Erl)
End Sub

Private Function GuardianAccionTexto() As String
    If m_AvisarGm Then
        GuardianAccionTexto = "avisar_gm"
    Else
        GuardianAccionTexto = "log"
    End If
End Function

Private Function SlotValido(ByVal UserIndex As Integer) As Boolean
    SlotValido = (UserIndex >= 1 And UserIndex <= m_Max)
End Function

Private Function TickNoCero(ByVal nowRaw As Long) As Long
    TickNoCero = nowRaw
    If TickNoCero = 0 Then TickNoCero = 1
End Function

' ---------------------------------------------------------------------------------------------
' Helpers de texto y cripto
' ---------------------------------------------------------------------------------------------

Private Function EsDecimal(ByVal texto As String, ByVal largoMax As Long) As Boolean
    Dim i As Long, c As Long
    If Len(texto) < 1 Or Len(texto) > largoMax Then Exit Function
    For i = 1 To Len(texto)
        c = AscW(Mid$(texto, i, 1))
        If c < 48 Or c > 57 Then Exit Function
    Next i
    EsDecimal = True
End Function

Private Function EsHex(ByVal texto As String, ByVal largo As Long) As Boolean
    Dim i As Long, c As Long
    If Len(texto) <> largo Then Exit Function
    For i = 1 To Len(texto)
        c = AscW(Mid$(texto, i, 1))
        Select Case c
            Case 48 To 57, 97 To 102
            Case Else
                Exit Function
        End Select
    Next i
    EsHex = True
End Function

' Solo [0-9a-z|,._-].
Private Function CharsetOk(ByVal linea As String) As Boolean
    Dim i As Long, c As Long
    For i = 1 To Len(linea)
        c = AscW(Mid$(linea, i, 1))
        Select Case c
            Case 48 To 57, 97 To 122, 124, 44, 46, 95, 45
            Case Else
                Exit Function
        End Select
    Next i
    CharsetOk = True
End Function

Private Function ValorHex(ByVal c As String) As Long
    ValorHex = InStr(1, "0123456789abcdef", c, vbBinaryCompare) - 1
End Function

Private Function TextoABytes(ByVal texto As String) As Byte()
    Dim b() As Byte
    If LenB(texto) = 0 Then Exit Function
    b = StrConv(texto, vbFromUnicode)
    TextoABytes = b
End Function

Private Function Hmac256(ByRef clave() As Byte, ByRef mensaje() As Byte) As Byte()
    Dim opcion As Long
    opcion = API_HMAC_SHA256
    Hmac256 = macBytes(mensaje, clave, opcion)
End Function

' KS = HMAC-SHA256(GS, "ao20-guard-v1" || N (16 crudos) || char_id uint32 LE): ancho fijo, 13 + 16 + 4 bytes.
Private Function DerivarKs(ByRef gs() As Byte, ByRef n() As Byte, ByVal charId As Long) As Byte()
    Dim msg() As Byte, i As Long
    ReDim msg(0 To 32)
    For i = 0 To 12
        msg(i) = CByte(Asc(Mid$(KS_LABEL, i + 1, 1)))
    Next i
    For i = 0 To 15
        msg(13 + i) = n(i)
    Next i
    msg(29) = CByte(charId And &HFF&)
    msg(30) = CByte((charId \ &H100&) And &HFF&)
    msg(31) = CByte((charId \ &H10000) And &HFF&)
    msg(32) = CByte((charId \ &H1000000) And &HFF&)
    DerivarKs = Hmac256(gs, msg)
End Function

' Comparacion de los 32 bytes con XOR acumulada: no corta en el primer byte distinto.
Private Function MacIgual(ByRef esperado() As Byte, ByVal macHex As String) As Boolean
    Dim i As Long, diff As Long, alto As Long, bajo As Long
    If cnvBytesLen(esperado) <> 32 Then Exit Function
    diff = 0
    For i = 0 To 31
        alto = ValorHex(Mid$(macHex, (i * 2) + 1, 1))
        bajo = ValorHex(Mid$(macHex, (i * 2) + 2, 1))
        If alto < 0 Or bajo < 0 Then
            diff = diff Or 256
        Else
            diff = diff Or (CLng(esperado(i)) Xor ((alto * 16) + bajo))
        End If
    Next i
    MacIgual = (diff = 0)
End Function

' HMAC(SaltLog, mid)[0..16] en hex minuscula: lo que se loguea en vez del id de PC.
Private Function MidReHasheado(ByVal midHex As String) As String
    Dim clave() As Byte, mensaje() As Byte, r() As Byte, h As String
    If LenB(m_SaltLog) = 0 Then
        MidReHasheado = "sinsal"
        Exit Function
    End If
    clave = TextoABytes(m_SaltLog)
    mensaje = TextoABytes(midHex)
    r = Hmac256(clave, mensaje)
    If cnvBytesLen(r) <> 32 Then
        MidReHasheado = "error"
        Exit Function
    End If
    h = LCase$(cnvHexStrFromBytes(r))
    MidReHasheado = Left$(h, 16)
End Function

Private Function BuscarKid(ByVal kid As Long) As Long
    Dim i As Long
    If kid = 0 And Not m_PermitirDev Then Exit Function
    For i = 1 To m_KidCant
        If m_Kids(i).Id = kid Then
            BuscarKid = i
            Exit Function
        End If
    Next i
End Function

' ---------------------------------------------------------------------------------------------
' Registro y politica
' ---------------------------------------------------------------------------------------------

' Una linea del log antitrampa. `avisar` = con GuardianAccion=avisar_gm tambien se avisa a los GMs (como mucho un aviso por
' jugador cada AVISO_GM_CADA_MS).
Private Sub Emitir(ByVal UserIndex As Integer, ByVal evento As String, ByVal detalle As String, ByVal avisar As Boolean)
    On Error GoTo Emitir_Err
    Dim nowRaw As Long, texto As String
    Call AntitrampaLineaAccion(UserIndex, evento, GuardianAccionTexto(), detalle)
    If avisar And m_AvisarGm Then
        nowRaw = GetTickCountRaw()
        With m_U(UserIndex)
            If .AvisoTick <> 0 Then
                If TicksElapsed(.AvisoTick, nowRaw) < AVISO_GM_CADA_MS Then Exit Sub
            End If
            .AvisoTick = TickNoCero(nowRaw)
        End With
        texto = "Antitrampa: " & UserList(UserIndex).name & " - " & evento & " (" & detalle & "). Accion: avisar_gm. Ver Logs\Antitrampa."
        Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg(texto, e_TextChannel.TEXTCHANNEL_SERVER_STAFF, e_FontTypeNames.FONTTYPE_SERVER))
    End If
    Exit Sub
Emitir_Err:
    Call TraceError(Err.Number, Err.Description, "modGuardian.Emitir", Erl)
End Sub

' True si este tipo de evento puede registrarse ahora (como mucho uno por jugador y tipo cada m_EventoCadaMs).
Private Function EventoPermitido(ByVal UserIndex As Integer, ByVal tipo As Long, ByVal nowRaw As Long) As Boolean
    With m_U(UserIndex)
        If .EventoTick(tipo) <> 0 Then
            If TicksElapsed(.EventoTick(tipo), nowRaw) < m_EventoCadaMs Then Exit Function
        End If
        .EventoTick(tipo) = TickNoCero(nowRaw)
    End With
    EventoPermitido = True
End Function

Private Sub ContarInvalido(ByVal UserIndex As Integer)
    With m_U(UserIndex)
        If .HbInvalidos < CONTADOR_TOPE Then .HbInvalidos = .HbInvalidos + 1
    End With
End Sub

Private Sub ResetEstado(ByVal UserIndex As Integer)
    Dim vacio As t_GuardianUser
    m_U(UserIndex) = vacio
End Sub

' ---------------------------------------------------------------------------------------------
' Login, desafio y logout
' ---------------------------------------------------------------------------------------------

' Desde ConnectUser_Complete, justo despues de WriteLoggedMessage.
Public Sub GuardianLogin(ByVal UserIndex As Integer)
    On Error GoTo GuardianLogin_Err
    Dim nowRaw As Long
    If Not SlotValido(UserIndex) Then Exit Sub
    If Not m_Habilitado Then Exit Sub
    nowRaw = GetTickCountRaw()
    Call ResetEstado(UserIndex)
    With m_U(UserIndex)
        .Activo = True
        .CharId = UserList(UserIndex).Id
        .DesafioTick = TickNoCero(nowRaw)
        .ResumenTick = TickNoCero(nowRaw)
        Call EmitirDesafio(UserIndex)
    End With
    Exit Sub
GuardianLogin_Err:
    Call TraceError(Err.Number, Err.Description, "modGuardian.GuardianLogin", Erl)
End Sub

Private Sub EmitirDesafio(ByVal UserIndex As Integer)
    On Error GoTo EmitirDesafio_Err
    Dim n() As Byte, largo As Long, i As Long, nHex As String
    largo = 16
    n = rngNonce(largo)
    If cnvBytesLen(n) <> 16 Then
        ' Sin desafio el jugador cuenta como ausente: se ve en el log, pero no se rompe el login.
        Call TraceError(0, "RNG_NonceData no devolvio 16 bytes", "modGuardian.EmitirDesafio", Erl)
        Exit Sub
    End If
    With m_U(UserIndex)
        For i = 0 To 15
            .N(i) = n(i)
        Next i
        .DesafioTick = TickNoCero(GetTickCountRaw())
        .Desafiado = True
    End With
    nHex = LCase$(cnvHexStrFromBytes(n))
    Call WriteGuardianChallenge(UserIndex, nHex, m_IntervaloSeg)
    Exit Sub
EmitirDesafio_Err:
    Call TraceError(Err.Number, Err.Description, "modGuardian.EmitirDesafio", Erl)
End Sub

' Desde TCP.CloseUser, ANTES de bajar UserLogged: resumen de la sesion y limpieza del slot.
Public Sub GuardianLogout(ByVal UserIndex As Integer)
    On Error GoTo GuardianLogout_Err
    If Not SlotValido(UserIndex) Then Exit Sub
    If m_U(UserIndex).Activo Then Call EmitirResumen(UserIndex, "logout")
    Call ResetEstado(UserIndex)
    Exit Sub
GuardianLogout_Err:
    Call TraceError(Err.Number, Err.Description, "modGuardian.GuardianLogout", Erl)
End Sub

Private Sub EmitirResumen(ByVal UserIndex As Integer, ByVal motivo As String)
    On Error GoTo EmitirResumen_Err
    Dim detalle As String, haceS As Long
    With m_U(UserIndex)
        haceS = -1
        If .UltimoOkTick <> 0 Then haceS = CLng(Fix(TicksElapsed(.UltimoOkTick, GetTickCountRaw()) / 1000#))
        detalle = "motivo=" & motivo & " hb_validos=" & .HbValidos & " hb_invalidos=" & .HbInvalidos & " ultimo_ok_hace_s=" & haceS & _
                " desafio=" & SiNo(.Desafiado)
    End With
    Call Emitir(UserIndex, "guardian_resumen", detalle, False)
    Exit Sub
EmitirResumen_Err:
    Call TraceError(Err.Number, Err.Description, "modGuardian.EmitirResumen", Erl)
End Sub

' ---------------------------------------------------------------------------------------------
' Id de PC (etapa E4e)
' ---------------------------------------------------------------------------------------------

' True si la sesion tiene un id de PC validado: al menos un heartbeat valido en ESTA sesion (el slot se limpia en
' login y logout). Sin guardian valido el id de PC es "desconocido", nunca "limpio".
Private Function TieneMidPc(ByVal UserIndex As Integer) As Boolean
    If Not SlotValido(UserIndex) Then Exit Function
    If UserIndex > m_Max Then Exit Function
    With m_U(UserIndex)
        TieneMidPc = (.Activo And .HbValidos > 0 And LenB(.MidPc) > 0)
    End With
End Function

' True si los dos jugadores tienen guardian valido y el MISMO id de PC (HMAC de la maquina, 32 hex). Si alguno no lo
' tiene, False: el origen cae a la regla de IP y al tope diario como hasta ahora.
Public Function GuardianMismaPc(ByVal UserA As Integer, ByVal UserB As Integer) As Boolean
    On Error GoTo GuardianMismaPc_Err
    If UserA = UserB Then Exit Function
    If Not TieneMidPc(UserA) Then Exit Function
    If Not TieneMidPc(UserB) Then Exit Function
    GuardianMismaPc = (StrComp(m_U(UserA).MidPc, m_U(UserB).MidPc, vbBinaryCompare) = 0)
    Exit Function
GuardianMismaPc_Err:
    GuardianMismaPc = False
    Call TraceError(Err.Number, Err.Description, "modGuardian.GuardianMismaPc", Erl)
End Function

' Id de PC re-hasheado (el mismo que va en guardian_ok) para el log, o "" si la sesion no tiene guardian valido.
Public Function GuardianPcLog(ByVal UserIndex As Integer) As String
    On Error GoTo GuardianPcLog_Err
    If Not TieneMidPc(UserIndex) Then Exit Function
    GuardianPcLog = MidReHasheado(m_U(UserIndex).MidPc)
    Exit Function
GuardianPcLog_Err:
    GuardianPcLog = vbNullString
    Call TraceError(Err.Number, Err.Description, "modGuardian.GuardianPcLog", Erl)
End Function

Private Function SiNo(ByVal valor As Boolean) As String
    If valor Then
        SiNo = "si"
    Else
        SiNo = "no"
    End If
End Function

' ---------------------------------------------------------------------------------------------
' Heartbeat
' ---------------------------------------------------------------------------------------------

' Desde Protocol.HandleGuardianHeartbeat (el paquete ya se leyo entero). Nunca cierra la conexion.
Public Sub GuardianProcesarHeartbeat(ByVal UserIndex As Integer, ByVal linea As String)
    On Error GoTo GuardianProcesarHeartbeat_Err
    Dim nowRaw As Long
    If Not SlotValido(UserIndex) Then Exit Sub
    If Not m_Habilitado Then Exit Sub
    nowRaw = GetTickCountRaw()
    ' (1) Rate limit ANTES de mirar el contenido. Todo hb que lo pasa gasta la ventana, sea valido o no.
    With m_U(UserIndex)
        If .UltimoHbTick <> 0 Then
            If TicksElapsed(.UltimoHbTick, nowRaw) < m_RateLimitMs Then
                If EventoPermitido(UserIndex, EV_FLOOD, nowRaw) Then
                    Call Emitir(UserIndex, "guardian_flood", "min_ms=" & m_RateLimitMs, False)
                End If
                Exit Sub
            End If
        End If
        .UltimoHbTick = TickNoCero(nowRaw)
    End With
    ' (2) Logueado y desafio emitido: si no, se descarta (la conexion sigue).
    If Not UserList(UserIndex).flags.UserLogged Or Not m_U(UserIndex).Desafiado Then
        If EventoPermitido(UserIndex, EV_SIN_DESAFIO, nowRaw) Then
            Call Emitir(UserIndex, "guardian_sin_desafio", "sin_desafio=1", False)
        End If
        Exit Sub
    End If
    Call ValidarLinea(UserIndex, linea, nowRaw)
    Exit Sub
GuardianProcesarHeartbeat_Err:
    Call TraceError(Err.Number, Err.Description, "modGuardian.GuardianProcesarHeartbeat", Erl)
End Sub

Private Sub LineaInvalida(ByVal UserIndex As Integer, ByVal motivo As String, ByVal nowRaw As Long)
    Call ContarInvalido(UserIndex)
    If EventoPermitido(UserIndex, EV_INVALIDA, nowRaw) Then
        Call Emitir(UserIndex, "guardian_linea_invalida", "motivo=" & motivo, False)
    End If
End Sub

Private Function BuscarGen(ByVal UserIndex As Integer, ByVal g As String) As Long
    Dim i As Long
    With m_U(UserIndex)
        For i = 1 To .GenCant
            If .Gen(i).G = g Then
                BuscarGen = i
                Exit Function
            End If
        Next i
    End With
End Function

Private Function EventoDeFlag(ByVal bit As Long) As String
    Select Case bit
        Case 1: EventoDeFlag = "guardian_exe_alterado"
        Case 2: EventoDeFlag = "guardian_recursos_alterados"
        Case 4: EventoDeFlag = "guardian_proceso_sospechoso"
        Case 8: EventoDeFlag = "guardian_debugger"
        Case 16: EventoDeFlag = "guardian_manifest_no_verificable"
        Case 32: EventoDeFlag = "guardian_dev"
        Case 64: EventoDeFlag = "guardian_sin_conexion"
        Case 128: EventoDeFlag = "guardian_squat"
        Case 256: EventoDeFlag = "guardian_dll_ajena"
    End Select
End Function

Private Sub ValidarLinea(ByVal UserIndex As Integer, ByVal linea As String, ByVal nowRaw As Long)
    On Error GoTo ValidarLinea_Err
    Dim campos() As String, kid As Long, kidIdx As Long, ctr As Double, ms As Double, flagsVal As Long
    Dim gs() As Byte, n() As Byte, ks() As Byte, cuerpo() As Byte, esperado() As Byte
    Dim i As Long, gi As Long, esperadoMs As Double, limiteMs As Double, g As String, bit As Long
    ' (3) Largo y charset, antes de cualquier cripto.
    If Len(linea) > MAX_LINE Then
        Call LineaInvalida(UserIndex, "largo", nowRaw)
        Exit Sub
    End If
    If Not CharsetOk(linea) Then
        Call LineaInvalida(UserIndex, "charset", nowRaw)
        Exit Sub
    End If
    campos = Split(linea, "|")
    If UBound(campos) <> 9 Then
        Call LineaInvalida(UserIndex, "formato_campos", nowRaw)
        Exit Sub
    End If
    If campos(0) <> "v1" Or Not EsDecimal(campos(1), 3) Or Not EsHex(campos(2), 8) Or Not EsDecimal(campos(3), 10) Or _
       Not EsDecimal(campos(4), 13) Or Not EsDecimal(campos(5), 5) Or Not EsHex(campos(6), 16) Or Not EsHex(campos(7), 32) Or _
       Not EsHex(campos(9), 64) Then
        Call LineaInvalida(UserIndex, "formato_valores", nowRaw)
        Exit Sub
    End If
    kid = CLng(campos(1))
    ctr = val(campos(3))
    ms = val(campos(4))
    flagsVal = CLng(campos(5))
    g = campos(2)
    If kid > 255 Or ctr > 4294967295# Then
        Call LineaInvalida(UserIndex, "formato_rango", nowRaw)
        Exit Sub
    End If
    ' (4a) kid aceptado (el kid 0 solo con PermitirDev=1 en el archivo secreto).
    kidIdx = BuscarKid(kid)
    If kidIdx = 0 Then
        Call ContarInvalido(UserIndex)
        Call Emitir(UserIndex, "guardian_firma_invalida", "motivo=kid kid=" & kid, True)
        Exit Sub
    End If
    If m_U(UserIndex).CharId < 0 Then
        Call ContarInvalido(UserIndex)
        Call Emitir(UserIndex, "guardian_firma_invalida", "motivo=char_id_invalido char_id=" & m_U(UserIndex).CharId, True)
        Exit Sub
    End If
    ' (4b) MAC recalculado con la clave de sesion KS y comparado con XOR acumulada.
    ReDim gs(0 To 31)
    For i = 0 To 31
        gs(i) = m_Kids(kidIdx).Secreto(i)
    Next i
    ReDim n(0 To 15)
    For i = 0 To 15
        n(i) = m_U(UserIndex).N(i)
    Next i
    ks = DerivarKs(gs, n, m_U(UserIndex).CharId)
    cuerpo = TextoABytes(Left$(linea, Len(linea) - 64))
    esperado = Hmac256(ks, cuerpo)
    If Not MacIgual(esperado, campos(9)) Then
        Call ContarInvalido(UserIndex)
        Call Emitir(UserIndex, "guardian_firma_invalida", "motivo=mac kid=" & kid & " g=" & g, True)
        Exit Sub
    End If
    ' (4c) Frescura de `ms` (ver la cabecera del modulo).
    gi = BuscarGen(UserIndex, g)
    If gi > 0 Then
        esperadoMs = m_U(UserIndex).Gen(gi).AnclaMs + TicksElapsed(m_U(UserIndex).Gen(gi).AnclaTick, nowRaw)
        If Abs(ms - esperadoMs) > m_FrescuraMs Then
            Call ContarInvalido(UserIndex)
            Call Emitir(UserIndex, "guardian_desfase", "g=" & g & " ctr=" & campos(3) & " ms=" & campos(4) & " esperado_ms=" & Format$(esperadoMs, "0") & _
                    " tolerancia_ms=" & m_FrescuraMs, True)
            Exit Sub
        End If
    Else
        limiteMs = TicksElapsed(m_U(UserIndex).DesafioTick, nowRaw) + m_FrescuraMs
        If ms > limiteMs Then
            Call ContarInvalido(UserIndex)
            Call Emitir(UserIndex, "guardian_desfase", "g=" & g & " ctr=" & campos(3) & " ms=" & campos(4) & " maximo_ms=" & Format$(limiteMs, "0") & _
                    " tolerancia_ms=" & m_FrescuraMs & " motivo=anterior_al_desafio", True)
            Exit Sub
        End If
    End If
    ' (4d) Contador por generacion.
    If gi > 0 Then
        If ctr <= m_U(UserIndex).Gen(gi).UltimoCtr Then
            Call ContarInvalido(UserIndex)
            Call Emitir(UserIndex, "guardian_replay", "g=" & g & " ctr=" & campos(3) & " ultimo=" & Format$(m_U(UserIndex).Gen(gi).UltimoCtr, "0"), True)
            Exit Sub
        End If
    Else
        If ctr < 1 Then
            Call ContarInvalido(UserIndex)
            Call Emitir(UserIndex, "guardian_replay", "g=" & g & " ctr=" & campos(3) & " ultimo=0", True)
            Exit Sub
        End If
        If m_U(UserIndex).GenCant >= MAX_GEN Then
            Call ContarInvalido(UserIndex)
            If EventoPermitido(UserIndex, EV_EXCESIVO, nowRaw) Then
                Call Emitir(UserIndex, "guardian_reinicio_excesivo", "g=" & g & " generaciones=" & m_U(UserIndex).GenCant & " maximo=" & MAX_GEN, True)
            End If
            Exit Sub
        End If
        With m_U(UserIndex)
            .GenCant = .GenCant + 1
            gi = .GenCant
            .Gen(gi).G = g
            .Gen(gi).AnclaTick = TickNoCero(nowRaw)
            .Gen(gi).AnclaMs = ms
            .Gen(gi).UltimoCtr = 0
        End With
        If m_U(UserIndex).GenCant > 1 Then
            If EventoPermitido(UserIndex, EV_REINICIO, nowRaw) Then
                Call Emitir(UserIndex, "guardian_reinicio", "g=" & g & " generaciones=" & m_U(UserIndex).GenCant, False)
            End If
        End If
    End If
    ' Valido.
    With m_U(UserIndex)
        .Gen(gi).UltimoCtr = ctr
        .UltimoOkTick = TickNoCero(nowRaw)
        .AusenteUltTick = 0
        If .HbValidos < CONTADOR_TOPE Then .HbValidos = .HbValidos + 1
        If .HbValidos = 1 Then
            ' E4e: id de PC de la sesion (crudo, 32 hex); solo se compara en memoria, nunca se loguea sin re-hashear.
            .MidPc = LCase$(campos(7))
            Call Emitir(UserIndex, "guardian_ok", "kid=" & kid & " g=" & g & " exe=" & campos(6) & " mid=" & MidReHasheado(campos(7)), False)
        End If
    End With
    ' (5) Un evento por cada bit de flags nuevo en la sesion.
    bit = 1
    For i = 0 To 8
        If (flagsVal And bit) <> 0 And (m_U(UserIndex).FlagsVistos And bit) = 0 Then
            m_U(UserIndex).FlagsVistos = m_U(UserIndex).FlagsVistos Or bit
            Call Emitir(UserIndex, EventoDeFlag(bit), "flags=" & flagsVal & " exe=" & campos(6) & " procs=" & campos(8), bit <> 32)
        End If
        bit = bit * 2
    Next i
    Exit Sub
ValidarLinea_Err:
    Call TraceError(Err.Number, Err.Description, "modGuardian.ValidarLinea", Erl)
End Sub

' ---------------------------------------------------------------------------------------------
' Loop principal: ausencia y resumen periodico
' ---------------------------------------------------------------------------------------------

' Una vez por vuelta del loop principal; trabaja como mucho una vez por TICK_MS.
Public Sub GuardianTick()
    On Error GoTo GuardianTick_Err
    Dim nowRaw As Long, i As Long
    If Not m_Habilitado Then Exit Sub
    nowRaw = GetTickCountRaw()
    If TicksElapsed(m_UltimoTick, nowRaw) < TICK_MS Then Exit Sub
    m_UltimoTick = nowRaw
    For i = 1 To LastUser
        If i <= m_Max Then
            If m_U(i).Activo Then Call TickUsuario(i, nowRaw)
        End If
    Next i
    Exit Sub
GuardianTick_Err:
    Call TraceError(Err.Number, Err.Description, "modGuardian.GuardianTick", Erl)
End Sub

Private Sub TickUsuario(ByVal UserIndex As Integer, ByVal nowRaw As Long)
    On Error GoTo TickUsuario_Err
    Dim refTick As Long, sinHbMs As Double, motivo As String, umbralMs As Double
    If Not UserList(UserIndex).flags.UserLogged Then
        Call ResetEstado(UserIndex)
        Exit Sub
    End If
    With m_U(UserIndex)
        refTick = .UltimoOkTick
        If refTick = 0 Then refTick = .DesafioTick
        sinHbMs = TicksElapsed(refTick, nowRaw)
        umbralMs = CDbl(m_FactorAusencia) * CDbl(m_IntervaloSeg) * 1000#
        ' W4: antes del primer heartbeat valido hay una gracia (arranque lento del guardian / del pipe).
        If .UltimoOkTick = 0 And CDbl(m_GraciaSeg) * 1000# > umbralMs Then umbralMs = CDbl(m_GraciaSeg) * 1000#
        If sinHbMs >= umbralMs Then
            If .AusenteUltTick = 0 Then
                motivo = "ausente"
            ElseIf TicksElapsed(.AusenteUltTick, nowRaw) >= CDbl(m_RepetirAusenteSeg) * 1000# Then
                motivo = "ausente"
            End If
            If LenB(motivo) > 0 Then
                .AusenteUltTick = TickNoCero(nowRaw)
                If Not .Desafiado Then
                    motivo = "sin_desafio"
                Else
                    motivo = "sin_heartbeat"
                End If
                Call Emitir(UserIndex, "guardian_ausente", "motivo=" & motivo & " sin_hb_s=" & CLng(Fix(sinHbMs / 1000#)) & " hb_validos=" & .HbValidos, motivo <> "sin_heartbeat")
            End If
        End If
        If TicksElapsed(.ResumenTick, nowRaw) >= CDbl(m_ResumenSeg) * 1000# Then
            .ResumenTick = TickNoCero(nowRaw)
            Call EmitirResumen(UserIndex, "periodico")
        End If
    End With
    Exit Sub
TickUsuario_Err:
    Call TraceError(Err.Number, Err.Description, "modGuardian.TickUsuario", Erl)
End Sub
