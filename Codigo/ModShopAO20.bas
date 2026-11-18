Attribute VB_Name = "ModShopAO20"
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

Public Enum e_DebitResult
    DebitOk = 0
    DebitRefused = 1     ' seguro: no se debito nada
    DebitUnknown = 2     ' el UPDATE pudo ejecutarse y no se pudo confirmar el saldo
End Enum

Public Sub init_transaction(ByVal obj_num As Long, ByVal UserIndex As Integer)
    On Error GoTo init_transaction_Err
    Call VigiaFuente(UserIndex, "tienda")  ' E2 vigia (modVigia)
    Dim obj As t_ObjData
    Dim debitedAccount As Long      ' cuenta debitada (0 = nada debitado)
    Dim debitedAmount As Long
    Dim delivered As Boolean
    obj.ObjNum = obj_num
    With UserList(UserIndex)
        'Me fijo si es un item de shop
        If Not is_purchaseable_item(obj) Then
            'Msg1087= Error al realizar la transacción
            Call WriteLocaleMsg(UserIndex, MSG_ERROR_REALIZAR_TRANSACCION, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_New_Naranja)
            Call LogShopErrors("El usuario " & .name & " intentó comprar un objeto que no es de shop (REVISAR) | " & obj.name)
            Exit Sub
        End If
        Call LoadPatronCreditsFromDB(UserIndex)
        If obj.Valor > .Stats.Creditos Then
            'Msg1088= Error al realizar la transacción.
            Call WriteLocaleMsg(UserIndex, MSG_ERROR_REALIZAR_TRANSACCION_1088, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_New_Naranja)
            Call LogShopErrors("El usuario " & .name & " intentó editar el valor del objeto (REVISAR) | " & obj.name)
            Exit Sub
        End If
        'Me fijo si tiene espacio en el inventario
        Dim objInventario As t_Obj
        objInventario.amount = 1
        objInventario.ObjIndex = obj.ObjNum
        If GetSlotForItemInInventory(UserIndex, objInventario) <= 0 Then
            'Msg1089= Asegurate de tener espacio suficiente en tu inventario.
            Call WriteLocaleMsg(UserIndex, MSG_ASEGURATE_TENER_ESPACIO_SUFICIENTE_INVENTARIO, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_SUBASTA)
            Exit Sub
        End If
        ' Flujo: espacio -> debito atomico en SQL -> entrega -> (si la entrega falla: reembolso) -> sincronizar -> avisar.
        Dim creditsLeft As Long
        Select Case DebitAccountCredits(.AccountID, obj.Valor, creditsLeft)
            Case e_DebitResult.DebitRefused
                Call WriteLocaleMsg(UserIndex, MSG_ERROR_REALIZAR_TRANSACCION_1088, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_New_Naranja)
                Call LogShopErrors("Debito de creditos rechazado por saldo en SQL | " & .name & " | " & ObjData(obj.ObjNum).name)
                Exit Sub
            Case e_DebitResult.DebitUnknown
                ' No se sabe si se cobro: se BLOQUEA la compra (no se reembolsa a ciegas: podria regalar creditos). Queda en Auditoria.
                Call WriteLocaleMsg(UserIndex, MSG_ERROR_REALIZAR_TRANSACCION, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_New_Naranja)
                Call LogShopErrors("Debito de creditos con resultado DESCONOCIDO | " & .name & " | " & ObjData(obj.ObjNum).name)
                Exit Sub
        End Select
        debitedAccount = .AccountID
        debitedAmount = obj.Valor
        Call SyncAccountCreditsInMemory(.AccountID, creditsLeft)
        delivered = MeterItemEnInventario(UserIndex, objInventario)
        If Not delivered Then
            Call RefundAfterFailedDelivery(UserIndex, debitedAccount, debitedAmount, ObjData(obj.ObjNum).name)
            debitedAccount = 0
            Call WriteLocaleMsg(UserIndex, MSG_ERROR_REALIZAR_TRANSACCION, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_New_Naranja)
            Exit Sub
        End If
        'Genero un log de los creditos que gasto y cuantos le quedan luego de la transaccion.
        Call LogShopTransactions(.name & " | Compro -> " & ObjData(obj.ObjNum).name & " | Valor -> " & obj.Valor)
        Call writeUpdateShopClienteCredits(UserIndex)
        Call RegisterTransaction(.AccountID, .Id, obj.ObjNum, obj.Valor, .Stats.Creditos)
        Call LogCreditosPatreon(.name & " | Compro objeto " & ObjData(obj.ObjNum).name & " | Creditos usados: " & obj.Valor & " | Creditos restantes: " & .Stats.Creditos)
    End With
    Exit Sub
init_transaction_Err:
    Call TraceError(Err.Number, Err.Description, "ShopAo20.init_transaction", Erl)
    ' Error tras un debito exitoso y antes de entregar: reembolso relativo (credits + ?) y resincronizar.
    If debitedAccount > 0 And Not delivered Then Call RefundAfterFailedDelivery(UserIndex, debitedAccount, debitedAmount, "error en init_transaction")
End Sub

' Lee el saldo de creditos de la cuenta directo de SQL (fuente de verdad). False si la lectura falla.
Public Function ReadAccountCreditsFromDb(ByVal AccountID As Long, ByRef Credits As Long) As Boolean
    On Error GoTo ReadAccountCreditsFromDb_Err
    Dim RS As ADODB.Recordset
    Set RS = Query("select offline_patron_credits from account where id = ?;", AccountID)
    If RS Is Nothing Then Exit Function
    If RS.EOF Then
        RS.Close
        Exit Function
    End If
    Credits = CLng(RS!offline_patron_credits)
    RS.Close
    Set RS = Nothing
    ReadAccountCreditsFromDb = True
    Exit Function
ReadAccountCreditsFromDb_Err:
    Call TraceError(Err.Number, Err.Description, "ShopAo20.ReadAccountCreditsFromDb", Erl)
    Set RS = Nothing
End Function

' Copia el saldo de SQL a la memoria de TODOS los personajes en linea de la cuenta.
Public Sub SyncAccountCreditsInMemory(ByVal AccountID As Long, ByVal Credits As Long)
    Dim i As Long
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged And UserList(i).AccountID = AccountID Then
            UserList(i).Stats.Creditos = Credits
        End If
    Next i
End Sub

' Debita creditos de la cuenta en SQL: el UPDATE resta (no escribe un total) y lleva la condicion de saldo en el WHERE.
' Verifica por relectura (convencion del proyecto, sin RecordsAffected): final = previo - monto -> DebitOk; final = previo
' -> DebitRefused; cualquier otra cosa o relectura fallida (reintenta una vez) -> DebitUnknown + ERROR en Auditoria con
' cuenta y monto para conciliar. Con DebitUnknown NO se reembolsa a ciegas (podria crear creditos): se bloquea la compra.
' CreditsLeft queda con el saldo real de SQL cuando se pudo leer.
Public Function DebitAccountCredits(ByVal AccountID As Long, ByVal Amount As Long, ByRef CreditsLeft As Long) As e_DebitResult
    On Error GoTo DebitAccountCredits_Err
    Dim creditsBefore As Long
    DebitAccountCredits = e_DebitResult.DebitRefused
    If Amount <= 0 Then Exit Function
    If Not ReadAccountCreditsFromDb(AccountID, creditsBefore) Then Exit Function
    CreditsLeft = creditsBefore
    If creditsBefore < Amount Then Exit Function
    Call Query("update account set offline_patron_credits = offline_patron_credits - ? where id = ? and offline_patron_credits >= ?;", Amount, AccountID, Amount)
    Dim readOk As Boolean
    readOk = ReadAccountCreditsFromDb(AccountID, CreditsLeft)
    If Not readOk Then readOk = ReadAccountCreditsFromDb(AccountID, CreditsLeft)
    If readOk Then
        If CreditsLeft = creditsBefore - Amount Then
            DebitAccountCredits = e_DebitResult.DebitOk
            Exit Function
        End If
        If CreditsLeft = creditsBefore Then Exit Function
    End If
    DebitAccountCredits = e_DebitResult.DebitUnknown
    Call LogCreditosPatreon("ERROR debito de creditos DESCONOCIDO | accountID: " & AccountID & " | monto: " & Amount & " | saldo previo: " & creditsBefore & " | releido: " & IIf(readOk, CStr(CreditsLeft), "sin lectura") & " | conciliar")
    Exit Function
DebitAccountCredits_Err:
    Call TraceError(Err.Number, Err.Description, "ShopAo20.DebitAccountCredits", Erl)
    DebitAccountCredits = e_DebitResult.DebitUnknown
    Call LogCreditosPatreon("ERROR debito de creditos con excepcion | accountID: " & AccountID & " | monto: " & Amount & " | conciliar")
End Function

' Reembolso relativo (credits + ?) de un debito ya hecho cuya entrega fallo; resincroniza todas las sesiones y deja rastro en Auditoria.
Public Sub RefundAfterFailedDelivery(ByVal UserIndex As Integer, ByVal AccountID As Long, ByVal Amount As Long, ByVal Reason As String)
    On Error GoTo RefundAfterFailedDelivery_Err
    Dim creditsNow As Long
    Call Query("update account set offline_patron_credits = offline_patron_credits + ? where id = ?;", Amount, AccountID)
    If ReadAccountCreditsFromDb(AccountID, creditsNow) Then
        Call SyncAccountCreditsInMemory(AccountID, creditsNow)
        If UserIndex > 0 Then Call writeUpdateShopClienteCredits(UserIndex)
    End If
    Call LogCreditosPatreon("REEMBOLSO de " & Amount & " creditos | accountID: " & AccountID & " | motivo: entrega fallida (" & Reason & ") | saldo: " & creditsNow)
    Exit Sub
RefundAfterFailedDelivery_Err:
    Call TraceError(Err.Number, Err.Description, "ShopAo20.RefundAfterFailedDelivery", Erl)
    Call LogCreditosPatreon("ERROR reembolso fallido | accountID: " & AccountID & " | monto: " & Amount & " | conciliar")
End Sub

' Avisa el saldo nuevo a TODAS las sesiones en linea de la cuenta (para que una tienda abierta se refresque).
Public Sub NotifyAccountCreditsToSessions(ByVal AccountID As Long)
    Dim i As Long
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged And UserList(i).AccountID = AccountID Then
            Call writeUpdateShopClienteCredits(i)
        End If
    Next i
End Sub

Private Function is_purchaseable_item(ByRef obj As t_ObjData) As Boolean
    Dim i As Long
    For i = 1 To UBound(ObjShop)
        If ObjShop(i).ObjNum = obj.ObjNum Then
            'Si es un item de shop, aparte le agrego el valor (por ref)
            obj.Valor = ObjShop(i).Valor
            is_purchaseable_item = True
            Exit Function
        End If
    Next i
    is_purchaseable_item = False
End Function

Private Sub RegisterTransaction(ByVal AccId As Long, ByVal CharId As Long, ByVal ItemId As Long, ByVal price As Long, ByVal CreditLeft As Long)
    On Error GoTo RegisterTransaction_Err
    Call Query("insert into patreon_shop_audit (acc_id, char_id, item_id, price, credit_left, time) VALUES (?,?,?,?,?, STRFTIME('%s'));", AccId, CharId, ItemId, price, CreditLeft)
    Exit Sub
RegisterTransaction_Err:
    Call TraceError(Err.Number, Err.Description, "ShopAo20.RegisterTransaction", Erl)
End Sub
