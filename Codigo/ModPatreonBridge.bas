Attribute VB_Name = "ModPatreonBridge"
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
Option Explicit

' ---------------------------------------------------------------------------
' Monetization bridge: syncs Mercado Pago tier before character load and
' pending web credits after login.
' Web endpoint contract (ao20-patreon-commerce):
'   GET  /api/bridge/sync?game_account_id=<id>  X-Service-Token: <token>
'        -> 200 plain text: TIER;<tierCode>;PAID_UNTIL;<iso>, then GRANT lines
'   GET  /api/bridge/pending?email=<email>      X-Service-Token: <token>
'        -> 200 plain text, one line per grant: grnt_id;credits
'   POST /api/bridge/claim                      X-Service-Token, Content-Type: application/json
'        body: {"claim_ids":["id1"],"game_account_id":42}
'        -> 200 plain text: OK;total_credits;claimed_count
' ---------------------------------------------------------------------------

' Module-level config cache (loaded once on first call)
Private m_ConfigLoaded   As Boolean
Private m_Enabled        As Boolean
Private m_BaseUrl        As String
Private m_ServiceToken   As String

' ---------------------------------------------------------------------------
' LoadBridgeConfig
' Reads [PatreonBridge] section from Server.ini into module-level cache.
' Defaults: URL=http://localhost:4001, token=empty string.
' ---------------------------------------------------------------------------
Private Sub LoadBridgeConfig()
    If m_ConfigLoaded Then Exit Sub
    On Error GoTo LoadBridgeConfig_Err

    Dim ini As clsIniManager
    Set ini = New clsIniManager
    Call ini.Initialize(IniPath & "Server.ini")

    Dim url As String
    url = ini.GetValue("PatreonBridge", "BaseUrl")
    If Len(url) = 0 Then url = "http://localhost:4001"
    m_BaseUrl = url

    m_ServiceToken = ini.GetValue("PatreonBridge", "ServiceToken")

    ' Enabled defaults OFF. The game server already reads the patron tier from
    ' account.is_active_patron (GetPatronTierFromAccountID); this synchronous HTTP
    ' pull to the commerce web is an opt-in refresh. Anything other than 1 = OFF,
    ' so local testing never blocks login on an unreachable web service.
    m_Enabled = (Trim$(ini.GetValue("PatreonBridge", "Enabled")) = "1")

    m_ConfigLoaded = True
    Exit Sub

LoadBridgeConfig_Err:
    ' If ini read fails, use safe defaults so login is never broken.
    m_BaseUrl = "http://localhost:4001"
    m_ServiceToken = ""
    m_Enabled = False
    m_ConfigLoaded = True
End Sub

' ---------------------------------------------------------------------------
' UrlEncodeSimple
' Minimal URL encoding sufficient for email addresses.
' Covers: space, @, +, %, &, =, #
' ---------------------------------------------------------------------------
Private Function UrlEncodeSimple(ByVal s As String) As String
    Dim result As String
    Dim i      As Integer
    Dim c      As String
    result = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        Select Case c
            Case " ":  result = result & "%20"
            Case "@":  result = result & "%40"
            Case "+":  result = result & "%2B"
            Case "%":  result = result & "%25"
            Case "&":  result = result & "%26"
            Case "=":  result = result & "%3D"
            Case "#":  result = result & "%23"
            Case Else: result = result & c
        End Select
    Next i
    UrlEncodeSimple = result
End Function

' ---------------------------------------------------------------------------
' SyncPatreonBridgeTier  (PUBLIC entry point)
' Called after account login and BEFORE LoadCharacterFromDB reads
' account.is_active_patron into UserList().Stats.tipoUsuario.
' Pulls the authoritative game tier code from /api/bridge/sync and writes it
' to the account table. On ANY error: logs and returns silently. Never breaks login.
' ---------------------------------------------------------------------------
Public Sub SyncPatreonBridgeTier(ByVal UserIndex As Integer)
    On Error GoTo SyncTier_Err

    Call LoadBridgeConfig
    If Not m_Enabled Then Exit Sub

    Dim accountID As Long
    accountID = UserList(UserIndex).AccountID
    If accountID <= 0 Then Exit Sub

    Dim http As MSXML2.ServerXMLHTTP60
    Set http = New MSXML2.ServerXMLHTTP60
    http.setTimeouts 2000, 2000, 3000, 3000

    http.Open "GET", m_BaseUrl & "/api/bridge/sync?game_account_id=" & CStr(accountID), False
    http.setRequestHeader "X-Service-Token", m_ServiceToken
    http.send

    If http.Status <> 200 Then
        Call LogCreditosPatreon("Bridge GET sync failed. Status: " & http.Status & " accountID: " & accountID)
        Exit Sub
    End If

    Dim rawLines() As String
    rawLines = Split(http.responseText, vbLf)
    If UBound(rawLines) < 0 Then Exit Sub

    Dim firstLine As String
    firstLine = Trim$(rawLines(0))
    If Len(firstLine) > 0 Then
        If Right$(firstLine, 1) = vbCr Then firstLine = Left$(firstLine, Len(firstLine) - 1)
    End If
    If Len(firstLine) = 0 Then Exit Sub

    Dim parts() As String
    parts = Split(firstLine, ";")
    If UBound(parts) < 1 Then Exit Sub
    If UCase$(Trim$(parts(0))) <> "TIER" Then Exit Sub

    Dim tierCode As Long
    tierCode = CLng(Val(parts(1)))

    Select Case tierCode
        Case 0, patron_tier_aventurero, patron_tier_heroe, patron_tier_leyenda, patron_tier_noble, patron_tier_emperador
            ' valid tier code
        Case Else
            Call LogCreditosPatreon("Bridge sync returned invalid tier code " & tierCode & " accountID: " & accountID)
            Exit Sub
    End Select

    Dim rsUpd As ADODB.Recordset
    Set rsUpd = Query("update account set is_active_patron = ? where id = ?;", tierCode, accountID)
    Set rsUpd = Nothing

    Exit Sub

SyncTier_Err:
    Call LogError("SyncPatreonBridgeTier error: " & Err.Description & " - UserIndex: " & UserIndex)
End Sub

' ---------------------------------------------------------------------------
' AppendPendingClaims  (PRIVATE helper)
' GETs a /api/bridge/pending URL and appends each grant id to claimIds (JSON-quoted,
' comma-separated), skipping ids already queued (dedup). Lets us claim grants from
' more than one identity (email + game_account_id) without double-counting.
' On ANY error: logs and returns silently. Never breaks login.
' ---------------------------------------------------------------------------
Private Sub AppendPendingClaims(ByVal pendingUrl As String, ByRef claimIds As String, ByRef claimCount As Integer)
    On Error GoTo AppendPendingClaims_Err

    Dim http As MSXML2.ServerXMLHTTP60
    Set http = New MSXML2.ServerXMLHTTP60
    http.setTimeouts 2000, 2000, 3000, 3000

    http.Open "GET", pendingUrl, False
    http.setRequestHeader "X-Service-Token", m_ServiceToken
    http.send

    If http.Status <> 200 Then
        Call LogCreditosPatreon("Bridge GET pending failed. Status: " & http.Status & " url: " & pendingUrl)
        Exit Sub
    End If

    Dim rawLines() As String
    rawLines = Split(http.responseText, vbLf)

    Dim i As Integer
    For i = 0 To UBound(rawLines)
        Dim oneLine As String
        oneLine = Trim$(rawLines(i))
        If Len(oneLine) > 0 Then
            If Right$(oneLine, 1) = vbCr Then oneLine = Left$(oneLine, Len(oneLine) - 1)
        End If
        If Len(oneLine) > 0 Then
            Dim parts() As String
            parts = Split(oneLine, ";")
            If UBound(parts) >= 1 Then
                Dim grantId As String
                grantId = Trim$(parts(0))
                If Len(grantId) > 0 Then
                    If InStr(claimIds, Chr$(34) & grantId & Chr$(34)) = 0 Then
                        If Len(claimIds) > 0 Then claimIds = claimIds & ","
                        claimIds = claimIds & Chr$(34) & grantId & Chr$(34)
                        claimCount = claimCount + 1
                    End If
                End If
            End If
        End If
    Next i

    Exit Sub

AppendPendingClaims_Err:
    Call LogError("AppendPendingClaims error: " & Err.Description & " - url: " & pendingUrl)
End Sub

' ---------------------------------------------------------------------------
' SyncPatreonBridgeCredits  (PUBLIC entry point)
' Called on character login from ConnectUser_Complete, after UserLogged=True. Applies pending web credit grants to the DB so later reads (shop open / EST) see the fresh balance.
' On ANY error: logs and returns silently. Never breaks login.
' ---------------------------------------------------------------------------
Public Sub SyncPatreonBridgeCredits(ByVal UserIndex As Integer)
    On Error GoTo SyncBridge_Err

    Call LoadBridgeConfig
    If Not m_Enabled Then Exit Sub

    Dim accountID As Long
    accountID = UserList(UserIndex).AccountID
    If accountID <= 0 Then Exit Sub

    ' Identity: game_account_id is authoritative (Mercado Pago grants carry it and
    ' may have an empty payer email). Email is the legacy fallback for Stripe/Patreon
    ' grants (created with an email and no game_account_id).
    Dim email As String
    Dim rsEmail As ADODB.Recordset
    Set rsEmail = Query("select email from account where id = ?;", accountID)
    If Not rsEmail Is Nothing Then
        If Not (rsEmail.EOF And rsEmail.BOF) Then email = Trim$(rsEmail!Email & "")
    End If
    Set rsEmail = Nothing

    Dim claimIds   As String
    Dim claimCount As Integer
    claimIds = ""
    claimCount = 0

    ' Source 1: grants tied to this account's email (Stripe/Patreon legacy).
    If Len(email) > 0 Then
        Call AppendPendingClaims(m_BaseUrl & "/api/bridge/pending?email=" & UrlEncodeSimple(email), claimIds, claimCount)
    End If

    ' Source 2: grants tied to this game_account_id (Mercado Pago).
    Call AppendPendingClaims(m_BaseUrl & "/api/bridge/pending?game_account_id=" & CStr(accountID), claimIds, claimCount)

    If claimCount = 0 Then Exit Sub

    ' --- POST /api/bridge/claim ---
    Dim body As String
    body = "{" & Chr$(34) & "claim_ids" & Chr$(34) & ":[" & claimIds & "]," & _
           Chr$(34) & "game_account_id" & Chr$(34) & ":" & CStr(accountID) & "}"

    Dim http As MSXML2.ServerXMLHTTP60
    Set http = New MSXML2.ServerXMLHTTP60
    http.setTimeouts 2000, 2000, 3000, 3000

    http.Open "POST", m_BaseUrl & "/api/bridge/claim", False
    http.setRequestHeader "X-Service-Token", m_ServiceToken
    http.setRequestHeader "Content-Type", "application/json"
    http.send body

    If http.Status <> 200 Then
        Call LogCreditosPatreon("Bridge POST claim failed. Status: " & http.Status & " accountID: " & accountID)
        Exit Sub
    End If

    Dim claimResp() As String
    claimResp = Split(http.responseText, ";")
    If UBound(claimResp) < 1 Then Exit Sub

    Dim totalCredits As Long
    totalCredits = CLng(Val(claimResp(1)))
    If totalCredits <= 0 Then Exit Sub

    Dim rsUpd As ADODB.Recordset
    Set rsUpd = Query("update account set offline_patron_credits = offline_patron_credits + ? where id = ?;", _
                    totalCredits, accountID)
    Set rsUpd = Nothing
    ' Releer el saldo real de SQL y propagarlo a TODAS las sesiones en linea de la cuenta (no queda memoria vieja).
    Dim creditsNow As Long
    If ReadAccountCreditsFromDb(accountID, creditsNow) Then Call SyncAccountCreditsInMemory(accountID, creditsNow)

    ' Sin email: el accountID ya identifica la cuenta para auditoria y no es un dato personal.
    Call LogCreditosPatreon("Bridge applied " & totalCredits & " credits to accountID " & accountID & ". Claims: " & claimCount)

    Exit Sub

SyncBridge_Err:
    Call LogError("SyncPatreonBridgeCredits error: " & Err.Description & " - UserIndex: " & UserIndex)
End Sub
