Attribute VB_Name = "Logging"
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
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
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit
Private Declare Function ReportEvent _
                Lib "advapi32.dll" _
                Alias "ReportEventA" (ByVal hEventLog As Long, _
                                      ByVal wType As Integer, _
                                      ByVal wCategory As Integer, _
                                      ByVal dwEventID As Long, _
                                      ByVal lpUserSid As Long, _
                                      ByVal wNumStrings As Integer, _
                                      ByVal dwDataSize As Long, _
                                      plpStrings As String, _
                                      lpRawData As Long) As Long

Private Enum eType_Log
    EventoDeSubasta = 0
    Ban = 1
    CreditosPatreon = 2
    ShopTransactions = 3
    ShopErrors = 4
    EdicionPaquete = 5
    MacroServidor = 6
    MacroCliente = 7
    VentaCasa = 8
    CriticEvent = 9
    EjercitoReal = 10
    EjercitoCaos = 11
    Error = 12
    Performance = 13
    Consulta = 14
    Clanes = 15
    GM = 16
    Premios = 17
    DatabaseError = 18
    Security = 19
    BankTransfer = 20
    SafeCommerceTransfer = 21
    SkinConsumption = 22
    QuestEvent = 23
    CardConsumption = 24
End Enum

Private Type t_CircularBuffer
    currentIndex As Integer
    Messages() As String
    size As Integer
End Type

Public CircularLogBuffer As t_CircularBuffer
Private Declare Function RegisterEventSource Lib "advapi32.dll" Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
Private Declare Function DeregisterEventSource Lib "advapi32.dll" (ByVal hEventLog As Long) As Long

' Handle al Event Log de Windows. Se abre UNA vez (la primera que se loguea) y se
' reusa mientras siga sirviendo. Antes se abria en CADA llamada a LogThis y no se
' liberaba nunca: como TraceError pasa por ahi, cada error del server filtraba un
' handle del sistema operativo, asi que cuanto mas fallaba mas se degradaba.
Private m_hEventLog As Long

' --- Niveles de severidad -------------------------------------------------
' Windows solo distingue Informacion/Advertencia/Error; DEBUG es un nivel propio
' nuestro, que existe para poder subir el detalle cuando se caza un problema.
Public Enum e_LogLevel
    LOG_DEBUG = 0
    LOG_INFO = 1
    LOG_WARN = 2
    LOG_ERROR = 3
End Enum

' Nivel minimo que se escribe. Arranca en INFO y lo pisa Configuracion.ini
' ([CONFIGURACIONES] LogLevel). Un GM lo cambia en caliente con ~LOGLEVEL.
Public g_LogLevel As e_LogLevel

' Arboles de archivos. Se separan porque tienen RETENCION distinta: lo operativo se
' purga a los 30 dias y la auditoria se guarda un anio, porque es la prueba ante un
' reclamo de pago o de sancion. La purga NUNCA debe tocar Auditoria.
Private Const LOG_TREE_OPERATIVO As String = "Operativo"
Private Const LOG_TREE_AUDITORIA As String = "Auditoria"

Public Sub InitializeCircularLogBuffer(Optional ByVal size As Integer = 30)
    CircularLogBuffer.size = size
    CircularLogBuffer.currentIndex = 0
    ReDim CircularLogBuffer.Messages(0 To size)
End Sub

Public Sub AddLogToCircularBuffer(Message As String)
    CircularLogBuffer.currentIndex = CircularLogBuffer.currentIndex + 1
    CircularLogBuffer.currentIndex = (CircularLogBuffer.currentIndex Mod CircularLogBuffer.size)
    CircularLogBuffer.Messages(CircularLogBuffer.currentIndex) = Message
End Sub

Public Function GetLastMessages() As String()
    Dim errorList() As String
    ReDim errorList(CircularLogBuffer.size)
    Dim i             As Integer
    Dim circularIndex As Integer
    For i = 1 To CircularLogBuffer.size
        circularIndex = ((CircularLogBuffer.currentIndex + i) Mod CircularLogBuffer.size)
        errorList(i) = CircularLogBuffer.Messages(circularIndex)
    Next i
    GetLastMessages = errorList
End Function

' Se llama al arrancar, antes del primer log.
Public Sub InitLogging()
    g_LogLevel = LOG_INFO
    Call InitializeCircularLogBuffer
End Sub

Private Function EventTypeToLevel(ByVal eventType As LogEventTypeConstants) As e_LogLevel
    Select Case eventType
        Case vbLogEventTypeError
            EventTypeToLevel = LOG_ERROR
        Case vbLogEventTypeWarning
            EventTypeToLevel = LOG_WARN
        Case Else
            EventTypeToLevel = LOG_INFO
    End Select
End Function

Public Function LevelTag(ByVal nivel As e_LogLevel) As String
    Select Case nivel
        Case LOG_ERROR
            LevelTag = "ERROR"
        Case LOG_WARN
            LevelTag = "WARN "
        Case LOG_DEBUG
            LevelTag = "DEBUG"
        Case Else
            LevelTag = "INFO "
    End Select
End Function

Public Function ParseLogLevel(ByVal texto As String) As e_LogLevel
    Select Case UCase$(Trim$(texto))
        Case "DEBUG"
            ParseLogLevel = LOG_DEBUG
        Case "WARN", "WARNING"
            ParseLogLevel = LOG_WARN
        Case "ERROR"
            ParseLogLevel = LOG_ERROR
        Case Else
            ParseLogLevel = LOG_INFO
    End Select
End Function

Public Sub SetLogLevel(ByVal nivel As e_LogLevel)
    g_LogLevel = nivel
    Call WriteToLogFile(LOG_TREE_OPERATIVO, TimeStampLog() & " INFO  LogLevel = " & LevelTag(nivel))
End Sub

Private Function TimeStampLog() As String
    TimeStampLog = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Function

' Escribe una linea en Logs\<arbol>\<arbol>_YYYY-MM-DD.log: un archivo por dia, el mismo
' patron que ya usaban los logs de venenos y elemental.
' OJO: el manejador de error NO puede llamar TraceError. TraceError termina en LogThis,
' que vuelve aca, y se arma una recursion infinita. Un log que falla se traga y listo.
Private Sub WriteToLogFile(ByVal arbol As String, ByVal linea As String)
    Dim fn As Integer
    On Error GoTo ErrHandler
    ' MkDir crea un solo nivel, asi que la raiz va primero: en una instalacion nueva
    ' la carpeta Logs no viene con el repo (esta gitignoreada).
    Dim raiz As String, carpeta As String
    raiz = App.Path & "\Logs"
    If LenB(Dir$(raiz, vbDirectory)) = 0 Then
        MkDir raiz
    End If
    carpeta = raiz & "\" & arbol
    If LenB(Dir$(carpeta, vbDirectory)) = 0 Then
        MkDir carpeta
    End If
    fn = FreeFile
    Open carpeta & "\" & arbol & "_" & Format$(Now, "yyyy-mm-dd") & ".log" For Append As #fn
    ' Un evento = UNA linea. Varios mensajes (los de TraceError, sobre todo) traen saltos
    ' de linea adentro, y si se escriben tal cual el archivo deja de poder leerse ni
    ' filtrarse por linea.
    Print #fn, Replace$(Replace$(linea, vbCrLf, " | "), vbLf, " | ")
    Close #fn
    Exit Sub
ErrHandler:
    On Error Resume Next
    If fn <> 0 Then Close #fn
End Sub

' Auditoria: plata real, sanciones y acciones de GM. Va a su propio arbol porque se
' guarda un anio y la purga de 30 dias no lo toca. NO pasa por el filtro de nivel:
' si algo es auditable, se registra siempre, aunque el server este en modo silencioso.
Public Sub LogAudit(ByVal categoria As String, ByVal texto As String)
    Call WriteToLogFile(LOG_TREE_AUDITORIA, TimeStampLog() & " " & categoria & " " & texto)
End Sub

' Arranque y apagado: al archivo y TAMBIEN al Visor de Eventos, que es donde un
' administrador espera ver que el servicio arranco o se cerro.
Public Sub LogServerLifecycle(ByVal texto As String)
    Call WriteToLogFile(LOG_TREE_OPERATIVO, TimeStampLog() & " INFO  " & texto)
    Call ReportToEventLog(0, texto, vbLogEventTypeInformation)
End Sub

Private Sub ReportToEventLog(ByVal nErrNo As Long, ByRef sLogMsg As String, ByVal eventType As LogEventTypeConstants)
    If m_hEventLog = 0 Then
        m_hEventLog = RegisterEventSource("", "Argentum20")
    End If
    ' dwEventID lleva el codigo real -categoria del wrapper o numero de error de VB- en vez
    ' del 20 fijo que tenia antes, para poder filtrar en el Visor de Eventos. Se descartan
    ' los negativos, que salen de los Err.Raise vbObjectError + N y quedarian ilegibles.
    Dim idEvento As Long
    idEvento = nErrNo
    If idEvento < 0 Then idEvento = 0
    ' Si ReportEvent falla -tipicamente porque se reinicio el servicio Event Log y el handle
    ' cacheado quedo muerto- se resetea para reabrirlo en la proxima llamada. Sin esto, un
    ' reinicio del servicio dejaba al server sin logs para siempre, en silencio.
    If ReportEvent(m_hEventLog, eventType, 0, idEvento, 0, 1, Len(sLogMsg), nErrNo & " - " & sLogMsg, 0) = 0 Then
        m_hEventLog = 0
    End If
End Sub

Public Sub LogThis(nErrNo As Long, sLogMsg As String, eventType As LogEventTypeConstants)
    Dim nivel As e_LogLevel
    nivel = EventTypeToLevel(eventType)
    If nivel < g_LogLevel Then Exit Sub
    If eventType = vbLogEventTypeWarning Or eventType = vbLogEventTypeError Then
        Call AddLogToCircularBuffer(sLogMsg)
    End If
    Call WriteToLogFile(LOG_TREE_OPERATIVO, TimeStampLog() & " " & LevelTag(nivel) & " " & nErrNo & " - " & sLogMsg)
    ' Al Visor de Eventos van SOLO los errores. El resto vive en archivo: ese log es un
    ' buzon compartido de tamanio fijo que Windows recorta solo cuando se llena, y el
    ' server llego a ocupar el 74% de todo lo que la maquina anotaba ahi.
    If eventType = vbLogEventTypeError Then
        Call ReportToEventLog(nErrNo, sLogMsg, eventType)
    End If
End Sub

' Comando de GM "~LOGLEVEL <DEBUG|INFO|WARN|ERROR>". Viaja como texto de chat, igual que
' los comandos de profesiones: asi se cambia el detalle de los logs sin reiniciar el
' server y sin necesidad de un paquete nuevo ni de recompilar el cliente.
Public Function HandleLogLevelChatCommand(ByVal UserIndex As Integer, ByVal chat As String) As Boolean
    On Error GoTo ErrHandler
    HandleLogLevelChatCommand = False
    Dim partes() As String
    partes = Split(Trim$(chat), " ")
    If UCase$(Trim$(partes(0))) <> "~LOGLEVEL" Then Exit Function
    HandleLogLevelChatCommand = True
    If UBound(partes) < 1 Then
        Call WriteConsoleMsg(UserIndex, "LogLevel actual: " & Trim$(LevelTag(g_LogLevel)) & ". Uso: ~LOGLEVEL DEBUG|INFO|WARN|ERROR", e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    Call SetLogLevel(ParseLogLevel(partes(1)))
    Call WriteConsoleMsg(UserIndex, "LogLevel = " & Trim$(LevelTag(g_LogLevel)), e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Logging.HandleLogLevelChatCommand", Erl)
End Function

' Libera el handle del Event Log. La llama General.Main al terminar.
Public Sub ShutdownLogging()
    If m_hEventLog <> 0 Then
        Call DeregisterEventSource(m_hEventLog)
        m_hEventLog = 0
    End If
End Sub

Public Sub LogearEventoDeSubasta(s As String)
    On Error GoTo ErrHandler
    Call LogAudit("SUBASTA", s)
    Exit Sub
ErrHandler:
End Sub


Public Sub LogBanByName(ByVal BannedName As String, ByVal Baneador As String, ByVal Motivo As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.Ban, "[Bans] " & BannedName & " BannedBy " & Baneador & " Reason " & Motivo, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal Motivo As String)
    On Error GoTo ErrHandler
    Dim s As String
    s = UserList(BannedIndex).name & " BannedBy " & UserList(UserIndex).name & " Reason " & Motivo
    Call LogAudit("BAN", s)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogCreditosPatreon(Desc As String)
    On Error GoTo ErrHandler
    Call LogAudit("CREDITOS", Desc)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogShopTransactions(Desc As String)
    On Error GoTo ErrHandler
    Call LogAudit("TIENDA", Desc)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogShopErrors(Desc As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.ShopErrors, "[MonetizationShopErrors.log] " & Desc, vbLogEventTypeError)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogEdicionPaquete(texto As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.EdicionPaquete, "[EdicionPaquete.log] " & texto, vbLogEventTypeWarning)
    Exit Sub
ErrHandler:
End Sub
Public Sub LogInfoServidor(texto As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.MacroServidor, "[Info] " & texto, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogMacroServidor(texto As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.MacroServidor, "[MacroServidor] " & texto, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogMacroCliente(texto As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.MacroCliente, "[MacroCliente] " & texto, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub logVentaCasa(ByVal texto As String)
    On Error GoTo ErrHandler
    Call LogAudit("PROPIEDAD", texto)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogCriticEvent(Desc As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.CriticEvent, "[Eventos.log] " & Desc, vbLogEventTypeWarning)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogEjercitoReal(Desc As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.EjercitoReal, "[EjercitoReal.log] " & Desc, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogEjercitoCaos(Desc As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.EjercitoCaos, "[EjercitoCaos.log] " & Desc, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogError(Desc As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.Error, "[Errores.log] " & Desc, vbLogEventTypeError)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogPerformance(Desc As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.Performance, "[Performance.log] " & Desc, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogConsulta(Desc As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.Consulta, "[obtenemos.log] " & Desc, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogClanes(ByVal str As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.Clanes, "[Clans.log] " & str, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogGM(name As String, Desc As String)
    On Error GoTo ErrHandler
    Call LogAudit("GM", "[" & name & "] " & Desc)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogPremios(GM As String, username As String, ByVal ObjIndex As Integer, ByVal Cantidad As Integer, Motivo As String)
    On Error GoTo ErrHandler
    Dim s As String
    s = "Item: " & ObjData(ObjIndex).name & " (" & ObjIndex & ") Cantidad: " & Cantidad & vbNewLine & "Motivo: " & Motivo & vbNewLine & vbNewLine
    Call LogAudit("PREMIO", "[" & GM & " -> " & username & "] " & s)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogDatabaseError(Desc As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.DatabaseError, "[Database.log] " & Desc, vbLogEventTypeError)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogSecurity(str As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.Security, "[Cheating.log] " & str, vbLogEventTypeWarning)
    Exit Sub
ErrHandler:
End Sub
Public Sub LogBankTransfer(ByVal originUser As String, ByVal targetUser As String, ByVal amount As Long, Optional ByVal receiverOnline As Boolean = False)
    On Error GoTo ErrHandler
    Dim transferContext As String
    If receiverOnline Then
        transferContext = "destinatario en línea"
    Else
        transferContext = "destinatario fuera de línea"
    End If
    Call LogAudit("BANCO", originUser & " transfirió " & amount & " monedas a " & targetUser & " (" & transferContext & ")")
    Exit Sub
ErrHandler:
End Sub

Public Sub LogSafeCommerceTransfer(ByVal originUser As String, ByVal targetUser As String, ByVal itemIndex As Integer, ByVal amount As Long, ByVal elementalTags As Long)
    On Error GoTo ErrHandler
    If itemIndex <= 0 Or amount <= 0 Then Exit Sub
    Call LogThis(eType_Log.SafeCommerceTransfer, "[SafeCommerceTransfers.log] Sender: " & originUser & " | Receiver: " & targetUser & " | Item: " & ObjData(itemIndex).name & " (" & itemIndex & ") | Amount: " & amount & " | ElementalTags: " & elementalTags, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogSkinConsumption(ByVal characterId As Long, ByVal characterName As String, ByVal itemIndex As Integer, ByVal amount As Long)
    On Error GoTo ErrHandler
    If itemIndex <= 0 Or amount <= 0 Then Exit Sub
    Call LogThis(eType_Log.SkinConsumption, "[SkinConsumptions.log] Character ID: " & characterId & " | Nick: " & characterName & " | Item: " & ObjData(itemIndex).name & " (" & itemIndex & ") | Amount: " & amount, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogCardConsumption(ByVal characterId As Long, ByVal characterName As String, ByVal accountId As Long, ByVal itemIndex As Integer, ByVal amount As Long)
    On Error GoTo ErrHandler
    If itemIndex <= 0 Or amount <= 0 Then Exit Sub
    Call LogThis(eType_Log.CardConsumption, "[CardConsumptions.log] Character ID: " & characterId & " | Nick: " & characterName & _
        " | Account ID: " & accountId & " | Item: " & ObjData(itemIndex).name & " (" & itemIndex & ") | Amount: " & amount & _
        " | Card Index: " & ObjData(itemIndex).CollectibleCardIndex & " | Rarity: " & ObjData(itemIndex).CollectibleCardRarity, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogQuestEvent(ByVal characterId As Long, ByVal characterName As String, ByVal questIndex As Integer, ByVal questName As String, ByVal action As String)
    On Error GoTo ErrHandler
    If questIndex <= 0 Then Exit Sub
    Call LogThis(eType_Log.QuestEvent, "[Quests.log] Character ID: " & characterId & " | Nick: " & characterName & " | Action: " & action & " | Quest: " & questName & " (" & questIndex & ")", vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub



Public Sub TraceError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
    #If DEBUGGING = 1 Then
        Debug.Print "TraceError: " & Descripcion & " " & Componente
    #End If
    Call LogThis(Numero, "Error number: " & Numero & " | Description: " & Descripcion & vbNewLine & "Component: " & Componente & " | Line number: " & Linea, vbLogEventTypeError)
End Sub
