Attribute VB_Name = "ModAccountBridge"
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
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Marquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
' Resultado de AccountBridge_VerifyLogin (plan 29.001).
' Distinguir RECHAZADO de ERROR es lo que permite decirle al jugador "el
' servicio de cuentas no esta disponible" en vez de "clave incorrecta" cuando
' la web esta caida.
Public Const VERIFY_LOGIN_OK As Long = 0
Public Const VERIFY_LOGIN_REJECTED As Long = 1
Public Const VERIFY_LOGIN_ERROR As Long = 2

Option Explicit

' ---------------------------------------------------------------------------
' Account bridge: aplica sobre la SQLite local las operaciones de cuenta
' (alta, cambio de password, borrado/restauracion) que la web encola en
' Postgres. El server es el UNICO escritor de su SQLite (Arquitectura B,
' plan 29.001); este modulo solo consume el canal de operaciones (pull) y
' confirma (ack) las que aplico, nunca al reves.
' Contrato del endpoint web (ao20-patreon-commerce, plan 08.001 Fase 1):
'   GET  /api/bridge/account-ops   X-Service-Token: <token>
'        -> 200 text/plain: linea 1 "OPS;<n>", luego una op por linea:
'           create_account: <op_id>;create_account;<email>;<salt>;<password_hash>;<validate_code>
'           set_password:   <op_id>;set_password;<email>;<salt>;<password_hash>
'           set_deleted:    <op_id>;set_deleted;<email>
'           unset_deleted:  <op_id>;unset_deleted;<email>
'   POST /api/bridge/account-ops/ack   X-Service-Token, Content-Type: text/plain
'        body: ids=1,2,3   -> 200 text/plain: OK;<acked_count>
' Fail-silent en toda la cadena: un error de red, de parseo o de SQL nunca
' debe romper el login ni el timer que dispara el poll.
' ---------------------------------------------------------------------------

' Module-level config cache (loaded once on first call)
Private m_ConfigLoaded            As Boolean
Private m_Enabled                 As Boolean
Private m_BaseUrl                 As String
Private m_ServiceToken            As String
Private m_RegisterUrl             As String
Private m_AllowIngameRegistration As Boolean

' Guard de intervalo minimo entre polls (evita llamadas espurias cuando
' Minuto_Timer y el fallback de pull-on-login-attempt coinciden).
Private Const ACCOUNT_BRIDGE_POLL_MIN_INTERVAL_MS As Double = 5000
Private m_LastPollTick  As Long
Private m_HasPolledOnce As Boolean

' ---- MAO bridge (mercado de personajes, plan 16.002): toggle e intervalo propios ----
Private m_MaoEnabled       As Boolean
Private Const MAO_BRIDGE_POLL_MIN_INTERVAL_MS As Double = 5000
Private m_LastMaoPollTick  As Long
Private m_HasPolledMaoOnce As Boolean

' ---- Char-ops bridge (borrar personaje, plan 2026-11-26): toggle borrar_personaje e intervalo propios ----
Private Const CHAR_OPS_POLL_MIN_INTERVAL_MS As Double = 5000
Private m_LastCharOpsPollTick  As Long
Private m_HasPolledCharOpsOnce As Boolean

' ---------------------------------------------------------------------------
' LoadAccountBridgeConfig
' Reads [AccountBridge] section from Server.ini into module-level cache.
' Defaults: bridge OFF, alta in-game OFF (el registro se hace desde la web).
' ---------------------------------------------------------------------------
Private Sub LoadAccountBridgeConfig()
    If m_ConfigLoaded Then Exit Sub
    On Error GoTo LoadAccountBridgeConfig_Err

    Dim ini As clsIniManager
    Set ini = New clsIniManager
    Call ini.Initialize(IniPath & "Server.ini")

    Dim url As String
    url = ini.GetValue("AccountBridge", "BaseUrl")
    If Len(url) = 0 Then url = "http://localhost:4000"
    m_BaseUrl = url

    m_ServiceToken = ini.GetValue("AccountBridge", "ServiceToken")

    Dim registerUrl As String
    registerUrl = ini.GetValue("AccountBridge", "RegisterUrl")
    If Len(registerUrl) = 0 Then registerUrl = "https://ao.muraliarevestimientos.com/registro"
    m_RegisterUrl = registerUrl

    ' Enabled por defecto OFF: el poll nunca corre solo, hay que prenderlo
    ' explicito en Server.ini una vez que la web este configurada.
    m_Enabled = (Trim$(ini.GetValue("AccountBridge", "Enabled")) = "1")

    ' AllowIngameRegistration por defecto OFF: el alta se hace desde la web.
    ' Se puede volver a habilitar como fallback si la web cae.
    m_AllowIngameRegistration = (Trim$(ini.GetValue("AccountBridge", "AllowIngameRegistration")) = "1")

    ' MAO (mercado de personajes, plan 16.002): OFF por defecto. Se prende en
    ' Server.ini una vez que la web tenga /api/bridge/mao-ops vivo.
    ' Plan 30.001: Server.ini esta gitignoreado y no hay plantilla versionada, asi
    ' que un deploy nuevo NO trae esta clave y el canal quedaba apagado en
    ' silencio mientras la web cobraba igual. Ahora se avisa por log (una sola vez
    ' por ejecucion, porque m_ConfigLoaded corta la relectura) y ademas el poll
    ' publica el valor en el header X-Game-Mao-Enabled para que la web no cobre.
    Dim maoRaw As String
    maoRaw = Trim$(ini.GetValue("AccountBridge", "MaoEnabled"))
    m_MaoEnabled = (maoRaw = "1")
    If LenB(maoRaw) = 0 Then
        Call LogError("AccountBridge: falta la clave MaoEnabled en la seccion [AccountBridge] de Server.ini. El canal del Mercado de Personajes queda APAGADO.")
    End If

    m_ConfigLoaded = True
    Exit Sub

LoadAccountBridgeConfig_Err:
    ' Si falla la lectura del ini, usamos defaults seguros: bridge apagado y
    ' alta in-game deshabilitada (el login nunca se rompe por esto).
    m_BaseUrl = "http://localhost:4000"
    m_ServiceToken = ""
    m_RegisterUrl = "https://ao.muraliarevestimientos.com/registro"
    m_Enabled = False
    m_AllowIngameRegistration = False
    m_MaoEnabled = False
    m_ConfigLoaded = True
End Sub

' Cupo de personajes de un tier, topeado igual que en ApplyMaoTransfer para que
' la web vea exactamente el numero que el server va a aplicar.
Private Function CapForTier(ByVal tier As e_TipoUsuario) As Long
    CapForTier = MaxCharacterForTier(tier)
    If CapForTier > MAX_PERSONAJES Then CapForTier = MAX_PERSONAJES
End Function

' Publica la tabla de cupos por tier para que la web NO la hardcodee (plan
' 30.001, ola 0). Viaja como header del GET del poll que ya se hace, asi que no
' agrega ni una llamada ni un endpoint. Formato:
'   "normal=4,aventurero=6,heroe=8,leyenda=10,noble=10,emperador=10"
' El server es la unica autoridad del cupo: si estos numeros cambian, la web se
' entera en el proximo poll sin tocar codigo web.
Private Function BuildCharCapsHeader() As String
    BuildCharCapsHeader = "normal=" & CapForTier(e_TipoUsuario.tNormal) & _
                          ",aventurero=" & CapForTier(e_TipoUsuario.tAventurero) & _
                          ",heroe=" & CapForTier(e_TipoUsuario.tHeroe) & _
                          ",leyenda=" & CapForTier(e_TipoUsuario.tLeyenda) & _
                          ",noble=" & CapForTier(e_TipoUsuario.tNoble) & _
                          ",emperador=" & CapForTier(e_TipoUsuario.tEmperador) & _
                          ",max=" & MAX_PERSONAJES
End Function

' Testigo de compilacion (plan 30.001). Devuelve con que constantes salio
' este binario. No cambia ningun comportamiento: solo arma un string.
' Existe para poder AUDITAR el .exe sin adivinar - los literales de VB6
' quedan en el binario en UTF-16 y tools/audit-server-binary.py los lee.
' Motivo: se cambio el cupo de personajes por tier, el binario no lo
' aplicaba, y no habia forma de saber que constantes se habian usado.
Public Function BuildStamp() As String
    Dim s As String
    s = "AO20STAMP_INICIO;"
    #If DEBUGGING Then
        s = s & "STAMP_DEBUGGING=1;"
    #Else
        s = s & "STAMP_DEBUGGING=0;"
    #End If
    #If PYMMO = 1 Then
        s = s & "STAMP_PYMMO=1;"
    #Else
        s = s & "STAMP_PYMMO=0;"
    #End If
    #If UsarQueSocket Then
        s = s & "STAMP_UsarQueSocket=1;"
    #Else
        s = s & "STAMP_UsarQueSocket=0;"
    #End If
    #If ConUpTime Then
        s = s & "STAMP_ConUpTime=1;"
    #Else
        s = s & "STAMP_ConUpTime=0;"
    #End If
    #If Lac Then
        s = s & "STAMP_Lac=1;"
    #Else
        s = s & "STAMP_Lac=0;"
    #End If
    #If AntiExternos Then
        s = s & "STAMP_AntiExternos=1;"
    #Else
        s = s & "STAMP_AntiExternos=0;"
    #End If
    #If UNLOCK_CPU Then
        s = s & "STAMP_UNLOCK_CPU=1;"
    #Else
        s = s & "STAMP_UNLOCK_CPU=0;"
    #End If
    #If DIRECT_PLAY Then
        s = s & "STAMP_DIRECT_PLAY=1;"
    #Else
        s = s & "STAMP_DIRECT_PLAY=0;"
    #End If
    #If BATTLESERVER = 1 Then
        s = s & "STAMP_BATTLESERVER=1;"
    #Else
        s = s & "STAMP_BATTLESERVER=0;"
    #End If
    s = s & "AO20STAMP_FIN"
    BuildStamp = s
End Function

' ---------------------------------------------------------------------------
' AccountBridge_AllowIngameRegistration  (PUBLIC)
' Usado por Protocol.HandleCreateAccount para decidir si el alta in-game
' sigue habilitada como fallback o si debe redirigir a la web.
' ---------------------------------------------------------------------------
Public Function AccountBridge_AllowIngameRegistration() As Boolean
    Call LoadAccountBridgeConfig
    AccountBridge_AllowIngameRegistration = m_AllowIngameRegistration
End Function

' ---------------------------------------------------------------------------
' AccountBridge_Enabled  (PUBLIC)
'
' Si el bridge esta prendido, la web es la FUENTE de los privilegios de GM y
' Server.ini no se consulta (plan 31.001, ola 5). Si esta apagado -dev local-,
' manda Server.ini como siempre. Es una sola pregunta y por eso vive aca: que
' Admin.bas tenga que saber como se lee la config del bridge seria repetir el
' criterio en dos lugares, y el dia que uno cambie el otro miente.
' ---------------------------------------------------------------------------
Public Function AccountBridge_Enabled() As Boolean
    Call LoadAccountBridgeConfig
    AccountBridge_Enabled = m_Enabled
End Function

' ---------------------------------------------------------------------------
' AccountBridge_PrivilegiosDePersonaje  (PUBLIC)
'
' Traduce el mapa que mando la web en el login de la cuenta al valor de
' e_PlayerType que le corresponde a UN personaje (plan 31.001, ola 5, D6).
'
' gmChars: "NOMBRE=rango[+rm],NOMBRE2=rango" (o vacio)
' rangos  : consejero | semidios | dios | admin
' flag    : +rm / +rolemaster, ADITIVO (e_PlayerType.RoleMaster es un bit
'           aparte, no un escalon: un RoleMaster es ademas alguna otra cosa)
'
' SIN ENTRADA PARA ESE NOMBRE, DEVUELVE User. Ese es todo el modelo: el mapa
' es de la cuenta que acaba de autenticarse, asi que un nombre que no esta no
' es GM y no hay ninguna lista global donde equivocarse. Si alguien escribio
' mal el nombre en la config de la web, el mapa igual le llega a SU cuenta y
' nunca matchea: el GM entra como jugador raso. La falla de un typo NUNCA es
' que un jugador entre como GM.
'
' Cualquier error de parseo cae en User y se logea: un mapa raro no puede ni
' elevar ni romper el login.
' ---------------------------------------------------------------------------
' True si el mapa de GMs de la cuenta trae AL MENOS un personaje admin o dios.
' Lo usa el cupo de personajes (esos dos rangos quedan exentos, decision del dueno
' 2026-08-07). SemiDios/Consejero/RM NO: siguen con el cupo de su tier.
' Mismo criterio de seguridad que PrivilegiosDePersonaje: lo que no se entiende no eleva.
Public Function AccountBridge_MapaTieneAdminODios(ByVal gmChars As String) As Boolean
    On Error GoTo MapaTieneAdminODios_Err
    Dim entradas() As String
    Dim partes()   As String
    Dim i          As Long
    Dim rango      As e_PlayerType
    If LenB(Trim$(gmChars)) = 0 Then Exit Function
    entradas = Split(gmChars, ",")
    For i = 0 To UBound(entradas)
        partes = Split(entradas(i), "=")
        If UBound(partes) = 1 Then
            rango = RangoDesdeTexto(partes(1))
            If (rango And e_PlayerType.Admin) <> 0 Or (rango And e_PlayerType.Dios) <> 0 Then
                AccountBridge_MapaTieneAdminODios = True
                Exit Function
            End If
        End If
    Next i
    Exit Function
MapaTieneAdminODios_Err:
    Call LogError("AccountBridge_MapaTieneAdminODios error: " & Err.Description)
End Function

Public Function AccountBridge_PrivilegiosDePersonaje(ByVal gmChars As String, _
                                                     ByVal charName As String) As e_PlayerType
    On Error GoTo PrivilegiosDePersonaje_Err

    Dim buscado  As String
    Dim entradas() As String
    Dim partes() As String
    Dim i        As Long

    AccountBridge_PrivilegiosDePersonaje = e_PlayerType.User

    If LenB(Trim$(gmChars)) = 0 Then Exit Function
    buscado = UCase$(Trim$(charName))
    If LenB(buscado) = 0 Then Exit Function

    entradas = Split(gmChars, ",")
    For i = 0 To UBound(entradas)
        partes = Split(entradas(i), "=")
        If UBound(partes) = 1 Then
            If UCase$(Trim$(partes(0))) = buscado Then
                AccountBridge_PrivilegiosDePersonaje = RangoDesdeTexto(partes(1))
                Exit Function
            End If
        End If
    Next i
    Exit Function

PrivilegiosDePersonaje_Err:
    Call LogError("AccountBridge_PrivilegiosDePersonaje error: " & Err.Description)
    AccountBridge_PrivilegiosDePersonaje = e_PlayerType.User
End Function

' ---------------------------------------------------------------------------
' RangoDesdeTexto  (PRIVATE)
'
' "dios" -> Dios ; "consejero+rm" -> Consejero Or RoleMaster.
'
' Un rango que no se reconoce devuelve User, NO el escalon mas bajo de GM:
' la web ya descarta lo que no entiende (staffGmChars.ts), asi que si algo
' raro llego hasta aca es un desajuste de version entre las dos puntas, y ahi
' la respuesta correcta es no dar privilegios.
' ---------------------------------------------------------------------------
Private Function RangoDesdeTexto(ByVal token As String) As e_PlayerType
    Dim partes() As String
    Dim rango    As e_PlayerType
    Dim i        As Long

    RangoDesdeTexto = e_PlayerType.User
    partes = Split(LCase$(Trim$(token)), "+")

    Select Case Trim$(partes(0))
        Case "admin"
            rango = e_PlayerType.Admin
        Case "dios"
            rango = e_PlayerType.Dios
        Case "semidios"
            rango = e_PlayerType.SemiDios
        Case "consejero"
            rango = e_PlayerType.Consejero
        Case Else
            Exit Function
    End Select

    For i = 1 To UBound(partes)
        If Trim$(partes(i)) = "rm" Or Trim$(partes(i)) = "rolemaster" Then
            rango = rango Or e_PlayerType.RoleMaster
        End If
    Next i

    RangoDesdeTexto = rango
End Function

' ---------------------------------------------------------------------------
' AccountBridge_RegisterUrl  (PUBLIC)
' URL de registro que se muestra cuando el alta in-game esta deshabilitada.
' ---------------------------------------------------------------------------
Public Function AccountBridge_RegisterUrl() As String
    Call LoadAccountBridgeConfig
    AccountBridge_RegisterUrl = m_RegisterUrl
End Function

' ---------------------------------------------------------------------------
' AccountBridge_Poll  (PUBLIC entry point)
' Llamado desde Minuto_Timer (cada 1 minuto) y como fallback desde
' Protocol.HandleLoginAccount (pull-on-login-attempt). Guard interno: no
' vuelve a pollear si ya lo hizo hace menos de 5 segundos.
' GET /api/bridge/account-ops, aplica cada operacion sobre la SQLite local
' y hace ack de las aplicadas. On ANY error: logea y sale en silencio, nunca
' rompe el login ni el timer.
' ---------------------------------------------------------------------------
Public Sub AccountBridge_Poll()
    On Error GoTo AccountBridge_Poll_Err

    Call LoadAccountBridgeConfig
    If Not m_Enabled Then Exit Sub

    Dim nowTick As Long
    nowTick = GetTickCountRaw()
    If m_HasPolledOnce Then
        If TicksElapsed(m_LastPollTick, nowTick) < ACCOUNT_BRIDGE_POLL_MIN_INTERVAL_MS Then Exit Sub
    End If
    m_LastPollTick = nowTick
    m_HasPolledOnce = True

    Dim http As MSXML2.ServerXMLHTTP60
    Set http = New MSXML2.ServerXMLHTTP60
    http.setTimeouts 2000, 2000, 3000, 3000

    http.Open "GET", m_BaseUrl & "/api/bridge/account-ops", False
    http.setRequestHeader "X-Service-Token", m_ServiceToken
    http.setRequestHeader "X-Game-Char-Caps", BuildCharCapsHeader()
    http.setRequestHeader "X-Game-Mao-Enabled", IIf(m_MaoEnabled, "1", "0")
    http.send

    If http.Status <> 200 Then
        Call LogError("AccountBridge poll GET failed. Status: " & http.Status)
        Exit Sub
    End If

    Dim rawLines() As String
    rawLines = Split(http.responseText, vbLf)
    If UBound(rawLines) < 0 Then Exit Sub

    Dim headerLine As String
    headerLine = Trim$(rawLines(0))
    If Len(headerLine) > 0 Then
        If Right$(headerLine, 1) = vbCr Then headerLine = Left$(headerLine, Len(headerLine) - 1)
    End If

    Dim headerParts() As String
    headerParts = Split(headerLine, ";")
    If UBound(headerParts) < 1 Then Exit Sub
    If UCase$(Trim$(headerParts(0))) <> "OPS" Then Exit Sub

    Dim opCount As Long
    opCount = CLng(Val(headerParts(1)))
    If opCount <= 0 Then Exit Sub

    Dim appliedIds As String
    appliedIds = ""

    Dim i As Long
    For i = 1 To UBound(rawLines)
        Dim oneLine As String
        oneLine = Trim$(rawLines(i))
        If Len(oneLine) > 0 Then
            If Right$(oneLine, 1) = vbCr Then oneLine = Left$(oneLine, Len(oneLine) - 1)
        End If
        If Len(oneLine) > 0 Then
            Dim appliedId As String
            appliedId = ApplyAccountOp(oneLine)
            If Len(appliedId) > 0 Then
                If Len(appliedIds) > 0 Then appliedIds = appliedIds & ","
                appliedIds = appliedIds & appliedId
            End If
        End If
    Next i

    If Len(appliedIds) > 0 Then Call AckAccountOps(appliedIds)

    Exit Sub
AccountBridge_Poll_Err:
    Call LogError("AccountBridge_Poll error: " & Err.Description)
End Sub

' ---------------------------------------------------------------------------
' ApplyAccountOp  (PRIVATE helper)
' Parsea una linea "op_id;tipo;campos..." y aplica la operacion sobre la
' SQLite local. Devuelve el op_id si se aplico (o si es idempotente-aplicada,
' ej. cuenta que ya existe / email que no matchea ninguna fila), o cadena
' vacia si la linea esta malformada (no se ackea, queda en la web para
' diagnostico).
' ---------------------------------------------------------------------------
Private Function ApplyAccountOp(ByVal OpLine As String) As String
    On Error GoTo ApplyAccountOp_Err
    ApplyAccountOp = ""

    Dim parts() As String
    parts = Split(OpLine, ";")
    If UBound(parts) < 2 Then
        ' NO loguear OpLine: en create_account/set_password lleva el salt y el hash.
        Call LogError("AccountBridge op malformada (faltan campos): " & (UBound(parts) + 1) & " campos")
        Exit Function
    End If

    Dim opId   As String
    Dim opType As String
    opId = Trim$(parts(0))
    opType = LCase$(Trim$(parts(1)))

    Select Case opType
        Case "create_account"
            If UBound(parts) < 5 Then
                ' NO loguear OpLine: lleva el salt y el hash de la password.
                Call LogError("AccountBridge op create_account malformada. opId=" & opId & " campos=" & (UBound(parts) + 1))
                Exit Function
            End If
            Call ApplyCreateAccountOp(opId, Trim$(parts(2)), parts(3), parts(4), parts(5))
            ApplyAccountOp = opId

        Case "set_password"
            If UBound(parts) < 4 Then
                ' NO loguear OpLine: lleva el salt y el hash de la password.
                Call LogError("AccountBridge op set_password malformada. opId=" & opId & " campos=" & (UBound(parts) + 1))
                Exit Function
            End If
            Call ApplySetPasswordOp(opId, Trim$(parts(2)), parts(3), parts(4))
            ApplyAccountOp = opId

        Case "set_deleted"
            Call ApplySetDeletedOp(opId, Trim$(parts(2)), 1)
            ApplyAccountOp = opId

        Case "unset_deleted"
            Call ApplySetDeletedOp(opId, Trim$(parts(2)), 0)
            ApplyAccountOp = opId

        Case Else
            Call LogError("AccountBridge op con tipo desconocido. opId=" & opId & " tipo=" & opType)
    End Select

    Exit Function
ApplyAccountOp_Err:
    Call LogError("ApplyAccountOp error: " & Err.Description & " - opId=" & opId)
    ApplyAccountOp = ""
End Function

' ---------------------------------------------------------------------------
' ApplyCreateAccountOp  (PRIVATE helper)
' Idempotente: si ya existe una cuenta con ese email, la operacion se
' considera aplicada (se ackea igual) y solo se logea un warning.
' ---------------------------------------------------------------------------
Private Sub ApplyCreateAccountOp(ByVal opId As String, ByVal email As String, ByVal salt As String, ByVal pwHash As String, ByVal validateCode As String)
    On Error GoTo ApplyCreateAccountOp_Err

    Dim existing As ADODB.Recordset
    Set existing = Query("SELECT id FROM account WHERE UPPER(email) = UPPER(?);", email)
    If Not (existing Is Nothing) Then
        If Not existing.EOF Then
            Call LogError("AccountBridge create_account op " & opId & ": la cuenta " & email & " ya existe, se trata como aplicada (idempotente).")
            Exit Sub
        End If
    End If

    Call Query("INSERT INTO account (email, password, salt, validate_code, validated) VALUES (?,?,?,?,1);", email, pwHash, salt, validateCode)
    Exit Sub
ApplyCreateAccountOp_Err:
    Call LogError("ApplyCreateAccountOp error: " & Err.Description & " - op " & opId & " email " & email)
End Sub

' ---------------------------------------------------------------------------
' ApplySetPasswordOp  (PRIVATE helper)
' El email se compara case-insensitive, igual que el login (plan 29.001): antes
' se comparaba exacto y una diferencia de mayusculas hacia que la op se
' descartara EN SILENCIO tratandola como aplicada.
' Si aun asi no existe la cuenta, se ackea (re-encolar no lo resolveria) pero se
' logea como ERROR, no como warning: es una password que el jugador cambio y
' que nunca llego.
' ---------------------------------------------------------------------------
Private Sub ApplySetPasswordOp(ByVal opId As String, ByVal email As String, ByVal salt As String, ByVal pwHash As String)
    On Error GoTo ApplySetPasswordOp_Err

    Dim existing As ADODB.Recordset
    Set existing = Query("SELECT id FROM account WHERE UPPER(email) = UPPER(?);", email)
    If existing Is Nothing Then
        Call LogError("ApplySetPasswordOp: SELECT fallo para email " & email & " - op " & opId)
        Exit Sub
    End If
    If existing.EOF Then
        ' NO es "aplicada": es una password que el jugador cambio y que nunca
        ' llego a su cuenta. Se logea como ERROR (no como warning) para que
        ' quede visible; el ack se hace igual porque re-encolarla tampoco la
        ' resolveria sola, pero el operador tiene que enterarse.
        Call LogError("ERROR AccountBridge set_password op " & opId & ": NO existe cuenta con email " & email & ". La password NO se aplico y la operacion se descarta.")
        Exit Sub
    End If

    Call Query("UPDATE account SET password = ?, salt = ? WHERE UPPER(email) = UPPER(?);", pwHash, salt, email)
    Exit Sub
ApplySetPasswordOp_Err:
    Call LogError("ApplySetPasswordOp error: " & Err.Description & " - op " & opId & " email " & email)
End Sub

' ---------------------------------------------------------------------------
' ApplySetDeletedOp  (PRIVATE helper)
' deletedValue: 1 = set_deleted, 0 = unset_deleted. Idempotente: si no existe
' cuenta con ese email, se considera aplicada y solo se logea un warning.
' ---------------------------------------------------------------------------
Private Sub ApplySetDeletedOp(ByVal opId As String, ByVal email As String, ByVal deletedValue As Integer)
    On Error GoTo ApplySetDeletedOp_Err

    Dim existing As ADODB.Recordset
    Set existing = Query("SELECT id FROM account WHERE UPPER(email) = UPPER(?);", email)
    If existing Is Nothing Then
        Call LogError("ApplySetDeletedOp: SELECT fallo para email " & email & " - op " & opId)
        Exit Sub
    End If
    If existing.EOF Then
        Call LogError("AccountBridge set_deleted/unset_deleted op " & opId & ": no existe cuenta con email " & email & ", se trata como aplicada.")
        Exit Sub
    End If

    Call Query("UPDATE account SET deleted = ? WHERE UPPER(email) = UPPER(?);", deletedValue, email)
    Exit Sub
ApplySetDeletedOp_Err:
    Call LogError("ApplySetDeletedOp error: " & Err.Description & " - op " & opId & " email " & email)
End Sub

' ---------------------------------------------------------------------------
' AckAccountOps  (PRIVATE helper)
' Confirma en la web los op_id aplicados. Content-Type explicito "text/plain":
' sin ese header express.text no parsea el body (hallazgo del review de
' Fase 1). Fail-silent: si el ack falla, la proxima pasada reaplica
' idempotente (no hay riesgo de duplicar datos).
' ---------------------------------------------------------------------------
Private Sub AckAccountOps(ByVal Ids As String)
    On Error GoTo AckAccountOps_Err

    Dim http As MSXML2.ServerXMLHTTP60
    Set http = New MSXML2.ServerXMLHTTP60
    http.setTimeouts 2000, 2000, 3000, 3000

    http.Open "POST", m_BaseUrl & "/api/bridge/account-ops/ack", False
    http.setRequestHeader "X-Service-Token", m_ServiceToken
    http.setRequestHeader "Content-Type", "text/plain"
    http.send "ids=" & Ids

    If http.Status <> 200 Then
        Call LogError("AccountBridge ack POST failed. Status: " & http.Status & " ids: " & Ids)
    End If

    Exit Sub
AckAccountOps_Err:
    Call LogError("AckAccountOps error: " & Err.Description & " - ids: " & Ids)
End Sub

' ===========================================================================
' MAO bridge (Mercado de Personajes, plan 16.002)
' Segundo canal de ops sobre el mismo [AccountBridge] (BaseUrl/ServiceToken),
' con su propio toggle m_MaoEnabled. El server sigue siendo el UNICO escritor
' de su SQLite (Arquitectura B); este canal solo aplica (pull) y confirma (ack).
' Contrato del endpoint web (16.002 F2):
'   GET  /api/bridge/mao-ops   X-Service-Token
'        -> 200 text/plain: linea 1 "MAOPS;<n>", luego una op por linea:
'           <op_id>;lock_character;<seller_account_id>;<char_id>
'           <op_id>;unlock_character;<seller_account_id>;<char_id>
'           <op_id>;transfer_character;<from_account_id>;<to_account_id>;<char_id>;<sale_transaction_id>
'   POST /api/bridge/mao-ops/ack   X-Service-Token, Content-Type: text/plain
'        body: una linea por op -> "<op_id>;ok" | "<op_id>;error;<detail_code>"
'        -> 200 text/plain: OK;<acked_count>
' Verificacion de escritura por READ-BACK de estado (no changes()/RecordsAffected):
' el UPDATE es un compare-and-swap en el WHERE y luego se relee la fila por PK y
' se interpreta su estado final. Verifica la post-condicion real y resuelve la
' idempotencia gratis. El precio (ARS) vive SOLO en la web; el server no escribe
' price_in_mao (decision 16.002 F0).
' ===========================================================================

' ---------------------------------------------------------------------------
' MaoBridge_Poll  (PUBLIC entry point). Llamado desde Minuto_Timer. Guard de
' intervalo propio. Fail-silent en toda la cadena.
' ---------------------------------------------------------------------------
Public Sub MaoBridge_Poll()
    On Error GoTo MaoBridge_Poll_Err

    Call LoadAccountBridgeConfig
    If Not m_MaoEnabled Then Exit Sub

    Dim nowTick As Long
    nowTick = GetTickCountRaw()
    If m_HasPolledMaoOnce Then
        If TicksElapsed(m_LastMaoPollTick, nowTick) < MAO_BRIDGE_POLL_MIN_INTERVAL_MS Then Exit Sub
    End If
    m_LastMaoPollTick = nowTick
    m_HasPolledMaoOnce = True

    Dim http As MSXML2.ServerXMLHTTP60
    Set http = New MSXML2.ServerXMLHTTP60
    http.setTimeouts 2000, 2000, 3000, 3000

    http.Open "GET", m_BaseUrl & "/api/bridge/mao-ops", False
    http.setRequestHeader "X-Service-Token", m_ServiceToken
    http.send

    If http.Status <> 200 Then
        Call LogError("MaoBridge poll GET failed. Status: " & http.Status)
        Exit Sub
    End If

    Dim rawLines() As String
    rawLines = Split(http.responseText, vbLf)
    If UBound(rawLines) < 0 Then Exit Sub

    Dim headerLine As String
    headerLine = Trim$(rawLines(0))
    If Len(headerLine) > 0 Then
        If Right$(headerLine, 1) = vbCr Then headerLine = Left$(headerLine, Len(headerLine) - 1)
    End If

    Dim headerParts() As String
    headerParts = Split(headerLine, ";")
    If UBound(headerParts) < 1 Then Exit Sub
    If UCase$(Trim$(headerParts(0))) <> "MAOPS" Then Exit Sub

    Dim opCount As Long
    opCount = CLng(Val(headerParts(1)))
    If opCount <= 0 Then Exit Sub

    Dim ackBody As String
    ackBody = ""

    Dim i As Long
    For i = 1 To UBound(rawLines)
        Dim oneLine As String
        oneLine = Trim$(rawLines(i))
        If Len(oneLine) > 0 Then
            If Right$(oneLine, 1) = vbCr Then oneLine = Left$(oneLine, Len(oneLine) - 1)
        End If
        If Len(oneLine) > 0 Then
            Dim ackLine As String
            ackLine = ApplyMaoOp(oneLine)
            If Len(ackLine) > 0 Then
                If Len(ackBody) > 0 Then ackBody = ackBody & vbLf
                ackBody = ackBody & ackLine
            End If
        End If
    Next i

    If Len(ackBody) > 0 Then Call AckMaoOps(ackBody)

    Exit Sub
MaoBridge_Poll_Err:
    Call LogError("MaoBridge_Poll error: " & Err.Description)
End Sub

' ---------------------------------------------------------------------------
' ApplyMaoOp  (PRIVATE). Parsea "op_id;tipo;campos..." y despacha. Devuelve la
' linea de ack ("<op_id>;ok" | "<op_id>;error;<detail>") o "" si la linea esta
' malformada (no se ackea). OJO: el contrato manda <seller/from>;<char/to>; las
' funciones reciben charId primero, por eso el mapeo posicional explicito.
' ---------------------------------------------------------------------------
Private Function ApplyMaoOp(ByVal OpLine As String) As String
    On Error GoTo ApplyMaoOp_Err
    ApplyMaoOp = ""

    Dim opId As String
    opId = ""

    Dim parts() As String
    parts = Split(OpLine, ";")
    If UBound(parts) < 2 Then
        Call LogError("MaoBridge op malformada (faltan campos): " & OpLine)
        Exit Function
    End If

    Dim opType As String
    opId = Trim$(parts(0))
    opType = LCase$(Trim$(parts(1)))

    Select Case opType
        Case "lock_character"
            If UBound(parts) < 3 Then
                Call LogError("MaoBridge op lock_character malformada: " & OpLine)
                Exit Function
            End If
            ApplyMaoOp = opId & ";" & ApplyMaoLock(CLng(parts(3)), CLng(parts(2)))

        Case "unlock_character"
            If UBound(parts) < 3 Then
                Call LogError("MaoBridge op unlock_character malformada: " & OpLine)
                Exit Function
            End If
            ApplyMaoOp = opId & ";" & ApplyMaoUnlock(CLng(parts(3)), CLng(parts(2)))

        Case "transfer_character"
            If UBound(parts) < 5 Then
                Call LogError("MaoBridge op transfer_character malformada: " & OpLine)
                Exit Function
            End If
            ApplyMaoOp = opId & ";" & ApplyMaoTransfer(CLng(parts(4)), CLng(parts(2)), CLng(parts(3)), Trim$(parts(5)))

        Case Else
            Call LogError("MaoBridge op con tipo desconocido: " & OpLine)
    End Select

    Exit Function
ApplyMaoOp_Err:
    Call LogError("ApplyMaoOp error: " & Err.Description & " - Line: " & OpLine)
    If Len(opId) > 0 Then
        ApplyMaoOp = opId & ";error;unknown_error"
    Else
        ApplyMaoOp = ""
    End If
End Function

' ---------------------------------------------------------------------------
' ApplyMaoLock  (PRIVATE)  -> "ok" | "error;<detail>". CAS: bloquea solo si el
' pj es del vendedor y no estaba bloqueado; relee por PK y verifica el estado
' final (idempotente). Si el pj esta online, lo kickea para que no pueda vaciar
' el inventario mientras esta publicado.
' ---------------------------------------------------------------------------
Private Function ApplyMaoLock(ByVal charId As Long, ByVal sellerAcc As Long) As String
    On Error GoTo ApplyMaoLock_Err

    ' GM gap (16.002 addendum): los personajes de GMs no se publican. El rol
    ' NO esta persistido en la tabla user (se resuelve por NOMBRE contra las
    ' listas de Administradores via EsGmChar), asi que se chequea aca antes
    ' del lock. La web, al recibir este detail code en el ack, cancela el
    ' listing automaticamente (auto-cancel del driver del bridge).
    Dim rsPre As ADODB.Recordset
    Set rsPre = Query("SELECT name, deleted, delete_requested_at FROM user WHERE id = ?;", charId)
    If rsPre Is Nothing Then
        ApplyMaoLock = "error;unknown_error"
        Exit Function
    End If
    If rsPre.EOF Then
        ApplyMaoLock = "error;char_not_found"
        Exit Function
    End If
    ' Plan borrar-personaje: un personaje borrado no existe para el mercado, y uno con pedido de
    ' borrado pendiente no se puede publicar (hay que cancelar el borrado entrando al juego).
    If CBool(SanitizeNullValue(rsPre!deleted, 0)) Then
        ApplyMaoLock = "error;char_not_found"
        Exit Function
    End If
    If Not IsNull(rsPre!delete_requested_at) Then
        ApplyMaoLock = "error;delete_pending"
        Exit Function
    End If
    If EsGmChar(CStr(rsPre!name)) Then
        ApplyMaoLock = "error;gm_character_blocked"
        Exit Function
    End If

    Call Query("UPDATE user SET is_locked_in_mao = 1 WHERE id = ? AND account_id = ? AND is_locked_in_mao = 0;", charId, sellerAcc)

    Dim rs As ADODB.Recordset
    Set rs = Query("SELECT name, account_id, is_locked_in_mao FROM user WHERE id = ?;", charId)
    If rs Is Nothing Then
        ApplyMaoLock = "error;unknown_error"
        Exit Function
    End If
    If rs.EOF Then
        ApplyMaoLock = "error;char_not_found"
        Exit Function
    End If
    If CLng(rs!account_id) <> sellerAcc Then
        ApplyMaoLock = "error;owner_mismatch"
        Exit Function
    End If
    If Not CBool(rs!is_locked_in_mao) Then
        ApplyMaoLock = "error;unknown_error"
        Exit Function
    End If

    Dim charName As String
    charName = CStr(rs!name)
    Dim uref As t_UserReference
    uref = NameIndex(charName)
    If IsValidUserRef(uref) Then
        Call modNetwork.Kick(UserList(uref.ArrayIndex).ConnectionDetails.ConnID, "Tu personaje fue publicado en el mercado.")
    End If

    ApplyMaoLock = "ok"
    Exit Function
ApplyMaoLock_Err:
    Call LogError("ApplyMaoLock error: " & Err.Description & " - char " & charId & " seller " & sellerAcc)
    ApplyMaoLock = "error;unknown_error"
End Function

' ---------------------------------------------------------------------------
' ApplyMaoUnlock  (PRIVATE)  -> "ok". Idempotente (0 filas = ya desbloqueado o
' no es del seller = ok).
' ---------------------------------------------------------------------------
Private Function ApplyMaoUnlock(ByVal charId As Long, ByVal sellerAcc As Long) As String
    On Error GoTo ApplyMaoUnlock_Err

    Call Query("UPDATE user SET is_locked_in_mao = 0, price_in_mao = 0 WHERE id = ? AND account_id = ?;", charId, sellerAcc)

    ApplyMaoUnlock = "ok"
    Exit Function
ApplyMaoUnlock_Err:
    Call LogError("ApplyMaoUnlock error: " & Err.Description & " - char " & charId & " seller " & sellerAcc)
    ApplyMaoUnlock = "error;unknown_error"
End Function

' ---------------------------------------------------------------------------
' ApplyMaoTransfer  (PRIVATE)  -> "ok" | "error;<detail>". Reparenta el pj al
' comprador de forma atomica y verificada. Valida en orden: existe / ya
' transferido (idempotente) / sigue bloqueado / es del vendedor / no online /
' cuenta destino existe-no borrada-no baneada / cupo / colision de nombre.
' NO toca clan (guild_index) ni matrimonio (spouse). saleTxId solo trazabilidad.
' ---------------------------------------------------------------------------
Private Function ApplyMaoTransfer(ByVal charId As Long, ByVal fromAcc As Long, ByVal toAcc As Long, ByVal saleTxId As String) As String
    On Error GoTo ApplyMaoTransfer_Err

    Dim rs As ADODB.Recordset
    Set rs = Query("SELECT name, account_id, is_locked_in_mao FROM user WHERE id = ?;", charId)
    If rs Is Nothing Then
        ApplyMaoTransfer = "error;unknown_error"
        Exit Function
    End If
    If rs.EOF Then
        ApplyMaoTransfer = "error;char_not_found"
        Exit Function
    End If

    If CLng(rs!account_id) = toAcc And Not CBool(rs!is_locked_in_mao) Then
        ApplyMaoTransfer = "ok"
        Exit Function
    End If
    If Not CBool(rs!is_locked_in_mao) Then
        ApplyMaoTransfer = "error;not_locked"
        Exit Function
    End If
    If CLng(rs!account_id) <> fromAcc Then
        ApplyMaoTransfer = "error;owner_mismatch"
        Exit Function
    End If

    Dim charName As String
    charName = CStr(rs!name)

    Dim uref As t_UserReference
    uref = NameIndex(charName)
    If IsValidUserRef(uref) Then
        ApplyMaoTransfer = "error;char_online"
        Exit Function
    End If

    Dim rsAcc As ADODB.Recordset
    Set rsAcc = Query("SELECT deleted, is_banned FROM account WHERE id = ?;", toAcc)
    If rsAcc Is Nothing Then
        ApplyMaoTransfer = "error;unknown_error"
        Exit Function
    End If
    If rsAcc.EOF Then
        ApplyMaoTransfer = "error;target_deleted"
        Exit Function
    End If
    If CBool(rsAcc!deleted) Then
        ApplyMaoTransfer = "error;target_deleted"
        Exit Function
    End If
    If CBool(rsAcc!is_banned) Then
        ApplyMaoTransfer = "error;target_banned"
        Exit Function
    End If

    ' Plan 30.001: una compra del mercado NO la limita el cupo del tier, solo el
    ' techo absoluto. El cupo del tier limita crear (ver HandleLoginNewChar en
    ' Protocol.bas); esta transferencia ya fue pagada, asi que un comprador de
    ' cuenta gratuita no puede rebotar por su tier despues de haber pagado.
    If GetPersonajesCountByIDDatabase(toAcc) >= MAX_PERSONAJES Then
        ApplyMaoTransfer = "error;target_over_cap"
        Exit Function
    End If

    Dim rsName As ADODB.Recordset
    Set rsName = Query("SELECT 1 FROM user WHERE account_id = ? AND name = ? COLLATE NOCASE AND id <> ? AND deleted = 0;", toAcc, charName, charId)
    If rsName Is Nothing Then
        ApplyMaoTransfer = "error;unknown_error"
        Exit Function
    End If
    If Not rsName.EOF Then
        ApplyMaoTransfer = "error;name_conflict"
        Exit Function
    End If

    Call Query("UPDATE user SET account_id = ?, is_locked_in_mao = 0, price_in_mao = 0 WHERE id = ? AND is_locked_in_mao = 1 AND account_id = ?;", toAcc, charId, fromAcc)

    Dim rsAfter As ADODB.Recordset
    Set rsAfter = Query("SELECT account_id, is_locked_in_mao FROM user WHERE id = ?;", charId)
    If rsAfter Is Nothing Then
        ApplyMaoTransfer = "error;unknown_error"
        Exit Function
    End If
    If rsAfter.EOF Then
        ApplyMaoTransfer = "error;unknown_error"
        Exit Function
    End If
    If CLng(rsAfter!account_id) = toAcc And Not CBool(rsAfter!is_locked_in_mao) Then
        ApplyMaoTransfer = "ok"
    Else
        ApplyMaoTransfer = "error;unknown_error"
    End If
    Exit Function
ApplyMaoTransfer_Err:
    Call LogError("ApplyMaoTransfer error: " & Err.Description & " - char " & charId & " from " & fromAcc & " to " & toAcc & " tx " & saleTxId)
    ApplyMaoTransfer = "error;unknown_error"
End Function

' ---------------------------------------------------------------------------
' AckMaoOps  (PRIVATE). Confirma en la web las ops procesadas. Body = una linea
' por op ("<op_id>;ok" | "<op_id>;error;<detail>") separadas por vbLf.
' Content-Type text/plain explicito (sin el, express.text no parsea). Fail-silent.
' ---------------------------------------------------------------------------
Private Sub AckMaoOps(ByVal Body As String)
    On Error GoTo AckMaoOps_Err

    Dim http As MSXML2.ServerXMLHTTP60
    Set http = New MSXML2.ServerXMLHTTP60
    http.setTimeouts 2000, 2000, 3000, 3000

    http.Open "POST", m_BaseUrl & "/api/bridge/mao-ops/ack", False
    http.setRequestHeader "X-Service-Token", m_ServiceToken
    http.setRequestHeader "Content-Type", "text/plain"
    http.send Body

    If http.Status <> 200 Then
        Call LogError("MaoBridge ack POST failed. Status: " & http.Status)
    End If

    Exit Sub
AckMaoOps_Err:
    Call LogError("AckMaoOps error: " & Err.Description)
End Sub

' ===========================================================================
' Char-ops bridge (borrar personaje, plan 2026-11-26-borrar-personaje)
' Tercer canal de ops sobre el mismo [AccountBridge] (BaseUrl/ServiceToken), con su propio toggle
' (feature_toggle.ini, borrar_personaje; apagado = no se consulta). Mismo molde que mao-ops: el server
' es el UNICO escritor de su SQLite; este canal solo aplica (pull) y confirma (ack).
' Contrato del endpoint web:
'   GET  /api/bridge/char-ops   X-Service-Token
'        -> 200 text/plain: linea 1 "OPS;<n>", luego una op por linea (hasta 50):
'           <op_id>;<tipo>;<char_id>;<email>
'           tipos: mark_char_delete | set_char_deleted
'   POST /api/bridge/char-ops/ack   X-Service-Token, Content-Type: text/plain
'        body: una linea por op -> "<op_id>;ok" | "<op_id>;error;<detalle>"
' Toda op con id legible recibe una linea de resultado (tipo desconocido -> error;unknown_type), asi
' una op rara nunca traba la cola. Una consulta que falla devuelve error, nunca ok.
' ===========================================================================

' ---------------------------------------------------------------------------
' CharOpsBridge_Poll  (PUBLIC entry point). Llamado desde Minuto_Timer. Guard de intervalo propio.
' Fail-silent en toda la cadena.
' ---------------------------------------------------------------------------
Public Sub CharOpsBridge_Poll()
    On Error GoTo CharOpsBridge_Poll_Err

    If Not IsFeatureEnabled("borrar_personaje") Then Exit Sub
    Call LoadAccountBridgeConfig

    Dim nowTick As Long
    nowTick = GetTickCountRaw()
    If m_HasPolledCharOpsOnce Then
        If TicksElapsed(m_LastCharOpsPollTick, nowTick) < CHAR_OPS_POLL_MIN_INTERVAL_MS Then Exit Sub
    End If
    m_LastCharOpsPollTick = nowTick
    m_HasPolledCharOpsOnce = True

    Dim http As MSXML2.ServerXMLHTTP60
    Set http = New MSXML2.ServerXMLHTTP60
    http.setTimeouts 2000, 2000, 3000, 3000

    http.Open "GET", m_BaseUrl & "/api/bridge/char-ops", False
    http.setRequestHeader "X-Service-Token", m_ServiceToken
    http.send

    If http.Status <> 200 Then
        Call LogError("CharOpsBridge poll GET failed. Status: " & http.Status)
        Exit Sub
    End If

    Dim rawLines() As String
    rawLines = Split(http.responseText, vbLf)
    If UBound(rawLines) < 0 Then Exit Sub

    Dim headerLine As String
    headerLine = Trim$(rawLines(0))
    If Len(headerLine) > 0 Then
        If Right$(headerLine, 1) = vbCr Then headerLine = Left$(headerLine, Len(headerLine) - 1)
    End If

    Dim headerParts() As String
    headerParts = Split(headerLine, ";")
    If UBound(headerParts) < 1 Then Exit Sub
    If UCase$(Trim$(headerParts(0))) <> "OPS" Then Exit Sub

    Dim opCount As Long
    opCount = CLng(Val(headerParts(1)))
    If opCount <= 0 Then Exit Sub

    Dim ackBody As String
    ackBody = vbNullString

    Dim i As Long
    For i = 1 To UBound(rawLines)
        Dim oneLine As String
        oneLine = Trim$(rawLines(i))
        If Len(oneLine) > 0 Then
            If Right$(oneLine, 1) = vbCr Then oneLine = Left$(oneLine, Len(oneLine) - 1)
        End If
        If Len(oneLine) > 0 Then
            Dim ackLine As String
            ackLine = ApplyCharOp(oneLine)
            If Len(ackLine) > 0 Then
                If Len(ackBody) > 0 Then ackBody = ackBody & vbLf
                ackBody = ackBody & ackLine
            End If
        End If
    Next i

    If Len(ackBody) > 0 Then Call AckCharOps(ackBody)

    Exit Sub
CharOpsBridge_Poll_Err:
    Call LogError("CharOpsBridge_Poll error: " & Err.Description)
End Sub

' ---------------------------------------------------------------------------
' CharOps_ParseLine  (PUBLIC, pura: la usan los tests). "<op_id>;<tipo>;<char_id>;<email>".
' opId se devuelve siempre que la linea tenga un primer campo legible, aunque el resto este mal, para
' poder confirmar la op como error en vez de dejarla trabando la cola. El email es todo lo que sigue
' al tercer ';' (la web ya le saca los ';', pero asi una linea rara no se corta en silencio).
' Devuelve False si la linea esta malformada.
' ---------------------------------------------------------------------------
Public Function CharOps_ParseLine(ByVal OpLine As String, ByRef opId As String, ByRef opType As String, ByRef charId As Long, ByRef email As String) As Boolean
    CharOps_ParseLine = False
    opId = vbNullString
    opType = vbNullString
    charId = 0
    email = vbNullString

    Dim parts() As String
    parts = Split(OpLine, ";", 4)
    If UBound(parts) >= 0 Then opId = Trim$(parts(0))
    If UBound(parts) < 3 Then Exit Function
    If LenB(opId) = 0 Then Exit Function

    opType = LCase$(Trim$(parts(1)))

    Dim charIdText As String
    charIdText = Trim$(parts(2))
    If LenB(charIdText) = 0 Or Len(charIdText) > 9 Then Exit Function
    Dim i As Long
    For i = 1 To Len(charIdText)
        If Mid$(charIdText, i, 1) < "0" Or Mid$(charIdText, i, 1) > "9" Then Exit Function
    Next i
    charId = CLng(charIdText)
    If charId <= 0 Then Exit Function

    email = Trim$(parts(3))
    If LenB(email) = 0 Then Exit Function

    CharOps_ParseLine = True
End Function

' Quita lo que rompe el formato de linea del ack (';', CR, LF) de un codigo de detalle.
Public Function CharOps_SanitizeDetail(ByVal Detail As String) As String
    Detail = Replace(Detail, ";", "_")
    Detail = Replace(Detail, vbCr, vbNullString)
    Detail = Replace(Detail, vbLf, vbNullString)
    Detail = Trim$(Detail)
    If LenB(Detail) = 0 Then Detail = "unknown_error"
    CharOps_SanitizeDetail = Detail
End Function

' Linea de ack: Result es "ok" o "error;<detalle>" (lo que devuelven las Apply*). Pura.
Public Function CharOps_FormatAck(ByVal opId As String, ByVal Result As String) As String
    If StrComp(Result, "ok", vbTextCompare) = 0 Then
        CharOps_FormatAck = opId & ";ok"
    ElseIf StrComp(Left$(Result, 6), "error;", vbTextCompare) = 0 Then
        CharOps_FormatAck = opId & ";error;" & CharOps_SanitizeDetail(Mid$(Result, 7))
    Else
        CharOps_FormatAck = opId & ";error;" & CharOps_SanitizeDetail(Result)
    End If
End Function

' ---------------------------------------------------------------------------
' ApplyCharOp  (PRIVATE). Devuelve la linea de ack, o "" si la linea no tiene ni op_id legible (no se puede
' confirmar: queda en la web para diagnostico).
' ---------------------------------------------------------------------------
Private Function ApplyCharOp(ByVal OpLine As String) As String
    On Error GoTo ApplyCharOp_Err
    ApplyCharOp = vbNullString

    Dim opId As String
    Dim opType As String
    Dim charId As Long
    Dim email As String
    If Not CharOps_ParseLine(OpLine, opId, opType, charId, email) Then
        ' NO se loguea la linea completa: lleva el email.
        Call LogError("CharOpsBridge op malformada. opId=" & opId)
        If LenB(opId) > 0 Then ApplyCharOp = CharOps_FormatAck(opId, "error;malformed")
        Exit Function
    End If

    Select Case opType
        Case "mark_char_delete"
            ApplyCharOp = CharOps_FormatAck(opId, ApplyMarkCharDelete(charId, email))
        Case "set_char_deleted"
            ApplyCharOp = CharOps_FormatAck(opId, ApplySetCharDeleted(charId, email))
        Case Else
            Call LogError("CharOpsBridge op con tipo desconocido. opId=" & opId & " tipo=" & opType)
            ApplyCharOp = CharOps_FormatAck(opId, "error;unknown_type")
    End Select
    Exit Function
ApplyCharOp_Err:
    Call LogError("ApplyCharOp error: " & Err.Description & " - opId=" & opId)
    If LenB(opId) > 0 Then ApplyCharOp = CharOps_FormatAck(opId, "error;unknown_error")
End Function

' El personaje esta conectado ahora mismo.
Private Function CharOps_IsOnline(ByVal charName As String) As Boolean
    Dim uref As t_UserReference
    uref = NameIndex(charName)
    CharOps_IsOnline = IsValidUserRef(uref)
End Function

' vbNullString si el personaje (cuenta dueña ownerAccountId) es de la cuenta de ese email; si no, el resultado de error.
Private Function CharOps_CheckOwner(ByVal ownerAccountId As Long, ByVal email As String) As String
    Dim rsAcc As ADODB.Recordset
    Set rsAcc = Query("SELECT id FROM account WHERE UPPER(email) = UPPER(?);", email)
    If rsAcc Is Nothing Then
        CharOps_CheckOwner = "error;unknown_error"
        Exit Function
    End If
    If rsAcc.EOF Then
        CharOps_CheckOwner = "error;not_owner"
        Exit Function
    End If
    If CLng(rsAcc!id) <> ownerAccountId Then
        CharOps_CheckOwner = "error;not_owner"
        Exit Function
    End If
    CharOps_CheckOwner = vbNullString
End Function

' ---------------------------------------------------------------------------
' ApplyMarkCharDelete  (PRIVATE)  -> "ok" | "error;<detail>". Pone user.delete_requested_at. Idempotente (ya
' marcado = ok, sin pisar la fecha original). Escritura verificada por relectura.
' ---------------------------------------------------------------------------
Private Function ApplyMarkCharDelete(ByVal charId As Long, ByVal email As String) As String
    On Error GoTo ApplyMarkCharDelete_Err

    Dim rs As ADODB.Recordset
    Set rs = Query("SELECT name, account_id, deleted, delete_requested_at, is_locked_in_mao FROM user WHERE id = ?;", charId)
    If rs Is Nothing Then
        ApplyMarkCharDelete = "error;unknown_error"
        Exit Function
    End If
    If rs.EOF Then
        ApplyMarkCharDelete = "error;not_found"
        Exit Function
    End If

    Dim ownerResult As String
    ownerResult = CharOps_CheckOwner(CLng(rs!account_id), email)
    If LenB(ownerResult) > 0 Then
        ApplyMarkCharDelete = ownerResult
        Exit Function
    End If
    If CBool(SanitizeNullValue(rs!deleted, 0)) Then
        ApplyMarkCharDelete = "error;already_deleted"
        Exit Function
    End If
    If CharOps_IsOnline(CStr(rs!name)) Then
        ApplyMarkCharDelete = "error;char_online"
        Exit Function
    End If
    If CBool(SanitizeNullValue(rs!is_locked_in_mao, 0)) Then
        ApplyMarkCharDelete = "error;mao_locked"
        Exit Function
    End If
    If Not IsNull(rs!delete_requested_at) Then
        ApplyMarkCharDelete = "ok"
        Exit Function
    End If

    Set rs = Query("UPDATE user SET delete_requested_at = CAST(strftime('%s','now') AS INTEGER) WHERE id = ? AND deleted = 0 AND delete_requested_at IS NULL;", charId)
    If rs Is Nothing Then
        ApplyMarkCharDelete = "error;unknown_error"
        Exit Function
    End If

    Set rs = Query("SELECT deleted, delete_requested_at FROM user WHERE id = ?;", charId)
    If rs Is Nothing Then
        ApplyMarkCharDelete = "error;unknown_error"
        Exit Function
    End If
    If rs.EOF Then
        ApplyMarkCharDelete = "error;unknown_error"
        Exit Function
    End If
    If IsNull(rs!delete_requested_at) Or CBool(SanitizeNullValue(rs!deleted, 0)) Then
        ApplyMarkCharDelete = "error;unknown_error"
        Exit Function
    End If

    Call LogAudit("BORRADO_PJ", "pedido de borrado marcado: char_id=" & charId)
    ApplyMarkCharDelete = "ok"
    Exit Function
ApplyMarkCharDelete_Err:
    Call LogError("ApplyMarkCharDelete error: " & Err.Description & " - char " & charId)
    ApplyMarkCharDelete = "error;unknown_error"
End Function

' ---------------------------------------------------------------------------
' ApplySetCharDeleted  (PRIVATE)  -> "ok" | "error;<detail>". Aplica el borrado logico (deleted = 1) cuando
' vence el plazo. Solo si la marca sigue puesta (si el dueno entro al juego, ya no esta: cancelled) y el
' personaje no esta conectado. Idempotente (ya borrado = ok).
'
' Toda la parte de base va en UNA transaccion sobre la conexion sincronica (Query). OJO: ahi no se pueden
' usar las rutinas que escriben con Execute (ExpulsarMiembro, SetLeader, ...): Execute rota por el pool
' asincrono de conexiones y chocaria con la transaccion abierta. Por eso el SQL va directo. La memoria
' (clan, castillo, pareja conectada) se sincroniza DESPUES del COMMIT.
' Clan: lider -> pasa al miembro mas antiguo (menor guild_members.id); si estaba solo el clan queda VACIO
' (no se borra la fila de guilds: user.guild_index es la posicion del clan en el array, borrar una fila
' corre los indices de todos los clanes siguientes).
' ---------------------------------------------------------------------------
Private Function ApplySetCharDeleted(ByVal charId As Long, ByVal email As String) As String
    On Error GoTo ApplySetCharDeleted_Err

    Dim rs As ADODB.Recordset
    Set rs = Query("SELECT name, account_id, deleted, delete_requested_at, is_locked_in_mao, guild_index, spouse FROM user WHERE id = ?;", charId)
    If rs Is Nothing Then
        ApplySetCharDeleted = "error;unknown_error"
        Exit Function
    End If
    If rs.EOF Then
        ApplySetCharDeleted = "error;not_found"
        Exit Function
    End If

    Dim ownerResult As String
    ownerResult = CharOps_CheckOwner(CLng(rs!account_id), email)
    If LenB(ownerResult) > 0 Then
        ApplySetCharDeleted = ownerResult
        Exit Function
    End If
    If CBool(SanitizeNullValue(rs!deleted, 0)) Then
        ApplySetCharDeleted = "ok"
        Exit Function
    End If
    If IsNull(rs!delete_requested_at) Then
        ApplySetCharDeleted = "error;cancelled"
        Exit Function
    End If

    Dim charName As String
    charName = CStr(rs!name)
    If CharOps_IsOnline(charName) Then
        ApplySetCharDeleted = "error;char_online"
        Exit Function
    End If
    If CBool(SanitizeNullValue(rs!is_locked_in_mao, 0)) Then
        ApplySetCharDeleted = "error;mao_locked"
        Exit Function
    End If

    Dim accountId As Long
    Dim guildNumber As Long
    Dim spouseId As Long
    accountId = CLng(rs!account_id)
    guildNumber = CLng(SanitizeNullValue(rs!guild_index, 0))
    spouseId = CLng(SanitizeNullValue(rs!spouse, 0))

    ' Pareja: solo cuenta si existe y esta viva. Si esta conectada, su copia en memoria se corrige despues
    ' del COMMIT (y se le avisa en vivo); si no, el aviso va a su message_info dentro de la transaccion.
    Dim spouseName As String
    Dim spouseOnline As Boolean
    If spouseId > 0 Then
        Dim rsSpouse As ADODB.Recordset
        Set rsSpouse = Query("SELECT name FROM user WHERE id = ? AND deleted = 0;", spouseId)
        If rsSpouse Is Nothing Then
            ApplySetCharDeleted = "error;unknown_error"
            Exit Function
        End If
        If rsSpouse.EOF Then
            spouseId = 0
        Else
            spouseName = CStr(rsSpouse!name)
            spouseOnline = CharOps_IsOnline(spouseName)
        End If
    End If

    Dim newLeaderId As Long
    Dim wasInGuild As Boolean
    Dim txOk As Boolean
    Dim rsTx As ADODB.Recordset

    DBError = vbNullString
    Set rsTx = Query("BEGIN TRANSACTION")
    If rsTx Is Nothing Then
        Call LogError("ApplySetCharDeleted: no se pudo abrir la transaccion. char " & charId & ": " & DBError)
        ApplySetCharDeleted = "error;unknown_error"
        Exit Function
    End If

    txOk = CharOps_DeleteInTransaction(charId, charName, spouseId, spouseOnline, newLeaderId, wasInGuild)
    If txOk Then
        DBError = vbNullString
        Set rsTx = Query("COMMIT")
        If rsTx Is Nothing Then
            Call LogError("ApplySetCharDeleted: COMMIT fallo. char " & charId & ": " & DBError)
            txOk = False
        End If
    End If
    If Not txOk Then
        Call Query("ROLLBACK")
        ApplySetCharDeleted = "error;unknown_error"
        Exit Function
    End If

    ' Verificacion por relectura: la post-condicion real es deleted = 1.
    Set rs = Query("SELECT deleted FROM user WHERE id = ?;", charId)
    If rs Is Nothing Then
        ApplySetCharDeleted = "error;unknown_error"
        Exit Function
    End If
    If rs.EOF Then
        ApplySetCharDeleted = "error;unknown_error"
        Exit Function
    End If
    If Not CBool(SanitizeNullValue(rs!deleted, 0)) Then
        ApplySetCharDeleted = "error;unknown_error"
        Exit Function
    End If

    ' ---- memoria, ya con la base confirmada ----
    If wasInGuild Then Call m_SincronizarClanTrasBorrado(guildNumber, newLeaderId)
    Call CastleForgetCharacter(charId, charName)
    If spouseOnline Then
        Dim spouseRef As t_UserReference
        spouseRef = NameIndex(spouseName)
        If IsValidUserRef(spouseRef) Then
            UserList(spouseRef.ArrayIndex).flags.Casado = 0
            UserList(spouseRef.ArrayIndex).flags.SpouseId = 0
            Call WriteLocaleMsg(spouseRef.ArrayIndex, MSG_DIVORCE_NOTIFICATION, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO, charName)
        End If
    End If

    Call LogAudit("BORRADO_PJ", "personaje borrado: char_id=" & charId & " char=" & charName & " account_id=" & accountId & _
            " clan=" & guildNumber & " nuevo_lider_id=" & newLeaderId & " pareja_id=" & spouseId)
    ApplySetCharDeleted = "ok"
    Exit Function
ApplySetCharDeleted_Err:
    Call LogError("ApplySetCharDeleted error: " & Err.Description & " - char " & charId)
    ApplySetCharDeleted = "error;unknown_error"
End Function

' ---------------------------------------------------------------------------
' CharOps_DeleteInTransaction  (PRIVATE). Cuerpo de la transaccion de ApplySetCharDeleted: devuelve False ante
' cualquier consulta que falle (el llamador hace ROLLBACK) y True si todo salio bien (el llamador hace COMMIT).
' ---------------------------------------------------------------------------
Private Function CharOps_DeleteInTransaction(ByVal charId As Long, ByVal charName As String, ByVal spouseId As Long, ByVal spouseOnline As Boolean, ByRef newLeaderId As Long, ByRef wasInGuild As Boolean) As Boolean
    On Error GoTo CharOps_DeleteInTransaction_Err
    CharOps_DeleteInTransaction = False
    newLeaderId = 0
    wasInGuild = False

    Dim rs As ADODB.Recordset
    Dim guildId As Long

    ' 1. Clan.
    DBError = vbNullString
    Set rs = Query("SELECT guild_id FROM guild_members WHERE user_id = ? LIMIT 1;", charId)
    If rs Is Nothing Then GoTo QueryFailed
    If Not rs.EOF Then
        guildId = CLng(rs!guild_id)
        wasInGuild = True
        Set rs = Query("SELECT leader_id FROM guilds WHERE id = ?;", guildId)
        If rs Is Nothing Then GoTo QueryFailed
        If Not rs.EOF Then
            If CLng(rs!leader_id) = charId Then
                Set rs = Query("SELECT user_id FROM guild_members WHERE guild_id = ? AND user_id <> ? ORDER BY id ASC LIMIT 1;", guildId, charId)
                If rs Is Nothing Then GoTo QueryFailed
                If Not rs.EOF Then
                    newLeaderId = CLng(rs!user_id)
                    Set rs = Query("UPDATE guilds SET leader_id = ? WHERE id = ?;", newLeaderId, guildId)
                    If rs Is Nothing Then GoTo QueryFailed
                End If
            End If
        End If
        Set rs = Query("DELETE FROM guild_members WHERE user_id = ? AND guild_id = ?;", charId, guildId)
        If rs Is Nothing Then GoTo QueryFailed
    End If
    ' Pedidos de ingreso a clanes que el personaje tenia pendientes.
    Set rs = Query("DELETE FROM guild_request WHERE user_id = ?;", charId)
    If rs Is Nothing Then GoTo QueryFailed

    ' 2. Castillo: queda sin dueno.
    Set rs = Query("UPDATE castle SET owner_character_id = NULL, owner_account_id = NULL WHERE owner_character_id = ?;", charId)
    If rs Is Nothing Then GoTo QueryFailed

    ' 3. castle_whitelist (por nombre).
    Set rs = Query("DELETE FROM castle_whitelist WHERE character_name = ? COLLATE NOCASE;", charName)
    If rs Is Nothing Then GoTo QueryFailed

    ' 4. Pareja: se limpia de los dos lados; si la pareja esta desconectada, mensaje offline.
    If spouseId > 0 Then
        Set rs = Query("UPDATE user SET spouse = 0 WHERE id = ? AND spouse = ?;", spouseId, charId)
        If rs Is Nothing Then GoTo QueryFailed
        If Not spouseOnline Then
            Set rs = Query("UPDATE user SET message_info = COALESCE(message_info, '') || ? WHERE id = ?;", _
                    charName & " se ha divorciado de ti porque su personaje fue borrado." & vbNewLine, spouseId)
            If rs Is Nothing Then GoTo QueryFailed
        End If
    End If

    ' 5. La fila: borrado logico.
    Set rs = Query("UPDATE user SET deleted = 1, deleted_at = CAST(strftime('%s','now') AS INTEGER), delete_requested_at = NULL, spouse = 0, guild_index = 0, guild_aspirant_index = 0 WHERE id = ? AND deleted = 0;", charId)
    If rs Is Nothing Then GoTo QueryFailed

    CharOps_DeleteInTransaction = True
    Exit Function

QueryFailed:
    Call LogError("CharOps_DeleteInTransaction: una consulta fallo. char " & charId & ": " & DBError)
    Exit Function
CharOps_DeleteInTransaction_Err:
    Call LogError("CharOps_DeleteInTransaction error: " & Err.Description & " - char " & charId)
End Function

' ---------------------------------------------------------------------------
' AckCharOps  (PRIVATE). Confirma en la web las ops procesadas. Body = una linea por op
' ("<op_id>;ok" | "<op_id>;error;<detalle>") separadas por vbLf. Content-Type text/plain explicito
' (sin el, express.text no parsea). Fail-silent: si el ack falla, la web reenvia la op en el proximo poll
' y las dos ops son idempotentes.
' ---------------------------------------------------------------------------
Private Sub AckCharOps(ByVal Body As String)
    On Error GoTo AckCharOps_Err

    Dim http As MSXML2.ServerXMLHTTP60
    Set http = New MSXML2.ServerXMLHTTP60
    http.setTimeouts 2000, 2000, 3000, 3000

    http.Open "POST", m_BaseUrl & "/api/bridge/char-ops/ack", False
    http.setRequestHeader "X-Service-Token", m_ServiceToken
    http.setRequestHeader "Content-Type", "text/plain"
    http.send Body

    If http.Status <> 200 Then
        Call LogError("CharOpsBridge ack POST failed. Status: " & http.Status)
    End If

    Exit Sub
AckCharOps_Err:
    Call LogError("AckCharOps error: " & Err.Description)
End Sub

' ---------------------------------------------------------------------------
' AccountBridge_VerifyLogin  (plan 29.001)
' Verificacion de credenciales DELEGADA a la web.
'
' El server ya no hashea ni compara passwords. Desde que el contrato pasa a
' argon2id hay una sola implementacion (Node) y este modulo solo pregunta;
' antes habia dos implementaciones espejo que tenian que coincidir byte a byte.
'
' Request  : POST <BaseUrl>/api/bridge/verify-login
'            X-Service-Token, Content-Type: text/plain
'            body = email & vbLf & password   (dos lineas, sin escapes)
' Respuesta: "OK;<account_id>;<validated>;<banned>;<deleted>;<gmmap>"
'            | "FAIL" | "ERR;<causa>"
'
' EL SEXTO CAMPO (plan 31.001, ola 5): el mapa de personajes GM de la cuenta,
' con formato "NOMBRE=rango[+rm],NOMBRE2=rango" o vacio. Es el MAPA COMPLETO
' y no el rango del que entra porque esto pasa cuando se autentica la CUENTA:
' el jugador todavia no eligio personaje. El rango se resuelve despues, al
' conectar (ver Admin.UserDarPrivilegioLevel).
'
' El nombre de un personaje solo puede tener letras A-Z y espacios
' (TCP.ValidarNombre, TCP.bas:1545), asi que ni ';' ni ',' ni '=' pueden
' aparecer adentro: el campo no necesita escapes y no puede inyectar campos
' falsos en la respuesta.
'
' Un campo AUSENTE (web vieja) se lee como mapa VACIO, nunca como error: sin
' mapa no hay privilegios, que es la direccion segura.
'
' Devuelve VERIFY_LOGIN_OK / VERIFY_LOGIN_REJECTED / VERIFY_LOGIN_ERROR.
' Cualquier problema de red, timeout o respuesta inesperada es ERROR, NUNCA
' REJECTED: un servicio caido no debe parecer una clave equivocada.
'
' La password NO se loguea en ningun camino, ni siquiera en el de error.
' ---------------------------------------------------------------------------
Public Function AccountBridge_VerifyLogin(ByVal email As String, _
                                          ByVal Password As String, _
                                          ByRef outAccountId As Long, _
                                          ByRef outValidated As Boolean, _
                                          ByRef outBanned As Boolean, _
                                          ByRef outDeleted As Boolean, _
                                          ByRef outGmChars As String) As Long
    On Error GoTo VerifyLogin_Err

    ' La config del bridge es lazy y esta guardada por m_ConfigLoaded, asi que
    ' TODO punto de entrada publico tiene que cargarla. Sin esto, el primer
    ' login despues de arrancar el server encuentra m_Enabled = False y falla
    ' con "servicio no disponible" hasta que el timer del minuto dispare el
    ' primer AccountBridge_Poll.
    Call LoadAccountBridgeConfig

    AccountBridge_VerifyLogin = VERIFY_LOGIN_ERROR
    outAccountId = 0
    outValidated = False
    outBanned = False
    outDeleted = False
    outGmChars = vbNullString

    If Not m_Enabled Then
        Call LogError("AccountBridge_VerifyLogin: el bridge esta deshabilitado, no se puede validar el login.")
        Exit Function
    End If

    Dim http As MSXML2.ServerXMLHTTP60
    Set http = New MSXML2.ServerXMLHTTP60
    http.setTimeouts 2000, 2000, 3000, 3000

    http.Open "POST", m_BaseUrl & "/api/bridge/verify-login", False
    http.setRequestHeader "X-Service-Token", m_ServiceToken
    http.setRequestHeader "Content-Type", "text/plain"
    http.send email & vbLf & Password

    If http.Status <> 200 Then
        Call LogError("AccountBridge verify-login HTTP status " & http.Status)
        Exit Function
    End If

    ' Trim$ saca espacios pero NO CR/LF, y de eso depende que "FAIL" matchee y
    ' que el ultimo campo ("deleted") se lea bien. Si un proxy agregara un
    ' salto, un rechazo se leeria como ERROR y una cuenta baneada entraria.
    Dim body As String
    Dim lastCh As String
    body = Trim$(http.responseText)
    Do While Len(body) > 0
        lastCh = Right$(body, 1)
        If lastCh = vbCr Or lastCh = vbLf Then
            body = Left$(body, Len(body) - 1)
        Else
            Exit Do
        End If
    Loop

    If body = "FAIL" Then
        AccountBridge_VerifyLogin = VERIFY_LOGIN_REJECTED
        Exit Function
    End If

    Dim parts() As String
    parts = Split(body, ";")
    If UBound(parts) < 4 Then
        Call LogError("AccountBridge verify-login respuesta inesperada: " & body)
        Exit Function
    End If
    If UCase$(Trim$(parts(0))) <> "OK" Then
        Call LogError("AccountBridge verify-login respuesta inesperada: " & body)
        Exit Function
    End If

    outAccountId = CLng(Val(parts(1)))
    If outAccountId <= 0 Then
        Call LogError("AccountBridge verify-login devolvio un account id invalido: " & body)
        outAccountId = 0
        Exit Function
    End If

    outValidated = (Trim$(parts(2)) = "1")
    outBanned = (Trim$(parts(3)) = "1")
    outDeleted = (Trim$(parts(4)) = "1")

    ' Mapa de personajes GM (plan 31.001, ola 5). OPCIONAL A PROPOSITO: si la
    ' web es vieja y no lo manda, UBound se queda en 4 y el mapa queda vacio.
    ' Eso NO es un error: significa que esta cuenta no tiene privilegios de GM,
    ' que es la respuesta segura cuando no se sabe. Un login jamas puede fallar
    ' por este campo.
    If UBound(parts) >= 5 Then
        outGmChars = Trim$(parts(5))
    End If

    AccountBridge_VerifyLogin = VERIFY_LOGIN_OK
    Exit Function

VerifyLogin_Err:
    Call LogError("AccountBridge_VerifyLogin error: " & Err.Description)
    AccountBridge_VerifyLogin = VERIFY_LOGIN_ERROR
End Function
