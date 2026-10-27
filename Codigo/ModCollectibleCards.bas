Attribute VB_Name = "ModCollectibleCards"
Option Explicit

'Cards And Npcs should be saved as the following
'If i have a scorpion npc and his card is the objIndex 32
'both the scorpion npc and the objindex32 should have matching properties
'CollectibleCardIndex (1-1024)
'CollectibleCardIndex is the direct index into the AccountCollectibleCardQuantities array
'If the card visually says that the scorpion is #234, then the quantity is stored at
'AccountCollectibleCardQuantities(234)
'Then the game compares the npc value with the value that the player has in his account
'whenever the business logic sees fit (eg: combat, finance, etc)

Private Const MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE = 1024
Private Const MAX_COLLECTIBLE_CARDS_BITARRAY_SIZE = 128

Public Enum e_CardRarity
    Bronce = 1
    Silver = 2
    gold = 3
End Enum

' Maximum copies of a single card an account can hold.
Private Const MAX_COLLECTIBLE_CARD_COPIES As Long = 10

' One row per card (plan 24.001, D2). Parameters: account_id, card_index, cap, inc, cap, inc.
' The sum and the cap happen inside SQLite, so two characters of the same account adding cards at
' the same time cannot overwrite each other (the old blob was rewritten whole on logout/autosave).
Private Const ADD_ACCOUNT_COLLECTIBLE_CARD As String = _
    "INSERT INTO account_collectible_card (account_id, card_index, quantity) " & _
    "VALUES (?, ?, MIN(?, ?)) " & _
    "ON CONFLICT(account_id, card_index) " & _
    "DO UPDATE SET quantity = MIN(?, quantity + ?);"

Private Const GET_ACCOUNT_COLLECTIBLE_CARDS As String = "SELECT card_index, quantity FROM account_collectible_card WHERE account_id = ?"
Private Const GET_ACCOUNT_COLLECTIBLE_CARD As String = "SELECT quantity FROM account_collectible_card WHERE account_id = ? AND card_index = ?"

Public Sub ResetUserAccountCollectibleCardCollection(ByRef User As t_User)
    ' Clears the in-memory card collection held by a user slot.
    '
    ' WHY THIS EXISTS: user slots are reused across sessions. When a player
    ' disconnects, ResetUserSlot wipes most fields but NOT the card arrays, so
    ' the freed slot keeps the previous account's card data in memory. If the
    ' next account to occupy the slot has no rows in the
    ' account_collectible_card table (the SELECT hits EOF), that stale data
    ' would be left untouched and the new account would inherit the previous
    ' account's card collection and all its gameplay bonuses (EXP, damage,
    ' damage reduction and drop rate).
    '
    ' This routine is called both when a slot is freed (TCP.ResetUserSlot) and
    ' right before loading a collection from the database, so the arrays are
    ' always consistent with the account that is currently logged in.
    Dim i As Integer

    For i = 1 To MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE
        User.AccountCollectibleCardQuantities(i) = 0
    Next i

    For i = 1 To MAX_COLLECTIBLE_CARDS_BITARRAY_SIZE
        User.AccountCollectibleCardBitArray(i) = 0
    Next i
End Sub

Public Function SetupUserAccountAccountCollectibleCardBitArray(ByRef User As t_User)
    On Error GoTo GetUserCollectibleCards_Err
    Dim RS As ADODB.Recordset
    Dim CardIndex As Long
    Dim Quantity As Long

    ' Wipe any card data left in memory by a previous account that occupied
    ' this slot before loading the real collection for the current account.
    ' Without this, an account with no saved cards (EOF on the SELECT below)
    ' would silently keep the previous account's cards.
    Call ResetUserAccountCollectibleCardCollection(User)

    ' One row per owned card: card_index is the same 1-based index the array uses.
    DBError = vbNullString
    Set RS = Query(GET_ACCOUNT_COLLECTIBLE_CARDS, User.AccountID)
    If RS Is Nothing Then
        Call TraceError(0, "Could not load collectible cards: " & DBError, "ModCollectibleCards.SetupUserAccountAccountCollectibleCardBitArray", Erl)
        Exit Function
    End If

    Do While Not RS.EOF
        CardIndex = CLng(RS!card_index)
        Quantity = CLng(RS!quantity)
        If CardIndex >= 1 And CardIndex <= MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
            If Quantity < 0 Then Quantity = 0
            If Quantity > 255 Then Quantity = 255
            User.AccountCollectibleCardQuantities(CardIndex) = CByte(Quantity)
        End If
        RS.MoveNext
    Loop
    RS.Close

    Set RS = Nothing
    Exit Function

GetUserCollectibleCards_Err:
    Call TraceError(Err.Number, Err.Description, "ModCollectibleCards.SetupUserAccountAccountCollectibleCardBitArray", Erl)
    Set RS = Nothing
End Function

' Reads one card quantity straight from the database (the source of truth: another character of the
' same account may have changed it after this one logged in). Returns False on a database failure.
Private Function ReadCardQuantityFromDb(ByVal AccountID As Long, ByVal CardIndex As Long, ByRef Quantity As Long) As Boolean
    On Error GoTo ReadCardQuantityFromDb_Err
    Dim RS As ADODB.Recordset
    DBError = vbNullString
    Set RS = Query(GET_ACCOUNT_COLLECTIBLE_CARD, AccountID, CardIndex)
    If RS Is Nothing Then Exit Function
    Quantity = 0
    If Not RS.EOF Then Quantity = CLng(RS!quantity)
    RS.Close
    Set RS = Nothing
    ReadCardQuantityFromDb = True
    Exit Function
ReadCardQuantityFromDb_Err:
    Call TraceError(Err.Number, Err.Description, "ModCollectibleCards.ReadCardQuantityFromDb", Erl)
    Set RS = Nothing
End Function

' Copies a card quantity into the in-memory collection of EVERY online character of the account,
' so the bonuses (GetCard*BonusForNpc) of the other connected characters are not stale.
Private Sub SyncCardQuantityInMemory(ByVal AccountID As Long, ByVal CardIndex As Long, ByVal Quantity As Long)
    Dim i As Long
    If Quantity < 0 Then Quantity = 0
    If Quantity > 255 Then Quantity = 255
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged And UserList(i).AccountID = AccountID Then
            UserList(i).AccountCollectibleCardQuantities(CardIndex) = CByte(Quantity)
        End If
    Next i
End Sub

' Consumes one card item into the account collection. Returns True when the item must be consumed.
'  - Returns False (item NOT consumed) when the account already holds MAX_COLLECTIBLE_CARD_COPIES of
'    the card, or when the database could not be read/written (nothing is granted in memory either).
'  - The increment depends on the rarity (CardBronce/Silver/GoldQuantity in the server config); with an
'    increment of 0 the item is still consumed and nothing is added (same as before plan 24.001).
'  - The quantity is capped at MAX_COLLECTIBLE_CARD_COPIES inside SQL and written immediately.
Public Function AddCollectibleCardToUser(ByVal UserIndex As Integer, ByRef ObjCard As t_Obj) As Boolean
    On Error GoTo AddCollectibleCardToUser_Err
    If ObjCard.ObjIndex = 0 Then Exit Function

    Dim CardIndex As Integer
    CardIndex = ObjData(ObjCard.ObjIndex).CollectibleCardIndex

    ' Validate card index
    If CardIndex < 1 Or CardIndex > MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then Exit Function

    Dim AccountID As Long
    Dim CurrentQty As Long
    Dim NewQty As Long
    Dim Inc As Long
    Dim RS As ADODB.Recordset
    AccountID = UserList(UserIndex).AccountID

    ' Cap check against the database, not against this character's possibly stale copy.
    If Not ReadCardQuantityFromDb(AccountID, CardIndex, CurrentQty) Then Exit Function
    Call SyncCardQuantityInMemory(AccountID, CardIndex, CurrentQty)
    If CurrentQty >= MAX_COLLECTIBLE_CARD_COPIES Then
        AddCollectibleCardToUser = False
        Exit Function
    End If

    If ObjData(ObjCard.ObjIndex).CollectibleCardRarity = e_CardRarity.Bronce Then
        Inc = CLng(SvrConfig.GetValue("CardBronceQuantity"))
    ElseIf ObjData(ObjCard.ObjIndex).CollectibleCardRarity = e_CardRarity.Silver Then
        Inc = CLng(SvrConfig.GetValue("CardSilverQuantity"))
    ElseIf ObjData(ObjCard.ObjIndex).CollectibleCardRarity = e_CardRarity.gold Then
        Inc = CLng(SvrConfig.GetValue("CardGoldQuantity"))
    Else
        Inc = 0
    End If

    If Inc > 0 Then
        DBError = vbNullString
        Set RS = Query(ADD_ACCOUNT_COLLECTIBLE_CARD, AccountID, CLng(CardIndex), MAX_COLLECTIBLE_CARD_COPIES, Inc, MAX_COLLECTIBLE_CARD_COPIES, Inc)
        If RS Is Nothing Then
            ' Not persisted: do not grant the card in memory and do not consume the item.
            Call TraceError(0, "Could not save collectible card " & CardIndex & " for account " & AccountID & ": " & DBError, "ModCollectibleCards.AddCollectibleCardToUser", Erl)
            Exit Function
        End If
        Set RS = Nothing

        ' Take the final value back from the database (it also reflects concurrent additions).
        If Not ReadCardQuantityFromDb(AccountID, CardIndex, NewQty) Then
            ' The write succeeded but the read-back failed: mirror what SQL just computed.
            NewQty = CurrentQty + Inc
            If NewQty > MAX_COLLECTIBLE_CARD_COPIES Then NewQty = MAX_COLLECTIBLE_CARD_COPIES
        End If
        Call SyncCardQuantityInMemory(AccountID, CardIndex, NewQty)
    End If

    AddCollectibleCardToUser = True
    Exit Function

AddCollectibleCardToUser_Err:
    Call TraceError(Err.Number, Err.Description, "ModCollectibleCards.AddCollectibleCardToUser", Erl)
    Set RS = Nothing
    AddCollectibleCardToUser = False
End Function

Public Function HasUserCollectedNpcCard(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
    Dim CardIndex As Integer
    CardIndex = NpcList(NpcIndex).CollectibleCardIndex
    
    If CardIndex < 1 Or CardIndex > MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
        HasUserCollectedNpcCard = False
        Exit Function
    End If
    
    ' Check if quantity is greater than 0
    HasUserCollectedNpcCard = (UserList(UserIndex).AccountCollectibleCardQuantities(CardIndex) > 0)
End Function

Public Function HasUserCollectedCard(ByVal UserIndex As Integer, ByVal CardIndex As Integer) As Boolean
    If CardIndex < 1 Or CardIndex > MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
        HasUserCollectedCard = False
        Exit Function
    End If
    
    HasUserCollectedCard = (UserList(UserIndex).AccountCollectibleCardQuantities(CardIndex) > 0)
End Function

Public Function GetUserCardQuantity(ByVal UserIndex As Integer, ByVal CardIndex As Integer) As Byte
    ' Returns the quantity of a specific card (0-255)
    If CardIndex < 1 Or CardIndex > MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
        GetUserCardQuantity = 0
    Else
        GetUserCardQuantity = UserList(UserIndex).AccountCollectibleCardQuantities(CardIndex)
    End If
End Function

Public Function GetCardExpBonusForNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Single
    ' Returns the EXP multiplier based on the card level for THIS specific NPC
    ' Returns 1.0 if no bonus applies
    Dim CardIndex As Integer
    Dim CardQuantity As Byte
    
    ' Get the card index for this NPC
    CardIndex = NpcList(NpcIndex).CollectibleCardIndex
    
    ' Validate card index
    If CardIndex < 1 Or CardIndex > MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
        GetCardExpBonusForNpc = 1#
        Exit Function
    End If
    
    ' Get the quantity of this specific card the user has
    CardQuantity = GetUserCardQuantity(UserIndex, CardIndex)
    
    ' Apply EXP bonus based on card level (quantity)
    ' Accumulative bonuses as the player levels up their card
    GetCardExpBonusForNpc = CSng(SvrConfig.GetValue("CardExpBonusBase"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardExpBonusLvl1")) Then GetCardExpBonusForNpc = CSng(SvrConfig.GetValue("CardExpBonus1"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardExpBonusLvl2")) Then GetCardExpBonusForNpc = CSng(SvrConfig.GetValue("CardExpBonus2"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardExpBonusLvl3")) Then GetCardExpBonusForNpc = CSng(SvrConfig.GetValue("CardExpBonus3"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardExpBonusLvl4")) Then GetCardExpBonusForNpc = CSng(SvrConfig.GetValue("CardExpBonus4"))
End Function

Public Function GetCardDamageBonusForNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Single
    ' Returns the damage multiplier based on the card level for THIS specific NPC
    ' Returns 1.0 if no bonus applies
    Dim CardIndex As Integer
    Dim CardQuantity As Byte
    
    ' Get the card index for this NPC
    CardIndex = NpcList(NpcIndex).CollectibleCardIndex
    
    ' Validate card index
    If CardIndex < 1 Or CardIndex > MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
        GetCardDamageBonusForNpc = 1#
        Exit Function
    End If
    
    ' Get the quantity of this specific card the user has
    CardQuantity = GetUserCardQuantity(UserIndex, CardIndex)
    
    ' Apply damage bonus based on card level (quantity)
    ' Accumulative bonuses as the player levels up their card
    GetCardDamageBonusForNpc = CSng(SvrConfig.GetValue("CardDamageBonusBase"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardDamageBonusLvl1")) Then GetCardDamageBonusForNpc = CSng(SvrConfig.GetValue("CardDamageBonus1"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardDamageBonusLvl2")) Then GetCardDamageBonusForNpc = CSng(SvrConfig.GetValue("CardDamageBonus2"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardDamageBonusLvl3")) Then GetCardDamageBonusForNpc = CSng(SvrConfig.GetValue("CardDamageBonus3"))
End Function

Public Function GetCardDamageReductionForNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Single
    ' Returns the damage reduction multiplier based on the card level for THIS specific NPC
    ' Returns 1.0 if no reduction applies (full damage)
    ' Returns values < 1.0 for damage reduction (e.g., 0.8 = 80% damage = 20% reduction)
    Dim CardIndex As Integer
    Dim CardQuantity As Byte
    
    ' Get the card index for this NPC
    CardIndex = NpcList(NpcIndex).CollectibleCardIndex
    
    ' Validate card index
    If CardIndex < 1 Or CardIndex > MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
        GetCardDamageReductionForNpc = 1#
        Exit Function
    End If
    
    ' Get the quantity of this specific card the user has
    CardQuantity = GetUserCardQuantity(UserIndex, CardIndex)
    
    ' Apply damage reduction based on card level (quantity)
    ' Lower values = more reduction (we multiply the incoming damage by this)
    GetCardDamageReductionForNpc = CSng(SvrConfig.GetValue("CardDamageReductionBase"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardDamageReductionLvl1")) Then GetCardDamageReductionForNpc = CSng(SvrConfig.GetValue("CardDamageReduction1"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardDamageReductionLvl2")) Then GetCardDamageReductionForNpc = CSng(SvrConfig.GetValue("CardDamageReduction2"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardDamageReductionLvl3")) Then GetCardDamageReductionForNpc = CSng(SvrConfig.GetValue("CardDamageReduction3"))
End Function


Public Function GetCardDropBonusForNpc(ByVal UserIndex As Integer, ByRef Npc As t_Npc) As Single
    Dim CardIndex As Integer
    Dim CardQuantity As Byte
    
    CardIndex = Npc.CollectibleCardIndex
    
    If CardIndex < 1 Or CardIndex > MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
        GetCardDropBonusForNpc = 1#
        Exit Function
    End If
    
    CardQuantity = GetUserCardQuantity(UserIndex, CardIndex)
    
    GetCardDropBonusForNpc = CSng(SvrConfig.GetValue("CardDropBonusBase"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardDropBonusLvl1")) Then GetCardDropBonusForNpc = CSng(SvrConfig.GetValue("CardDropBonus1"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardDropBonusLvl2")) Then GetCardDropBonusForNpc = CSng(SvrConfig.GetValue("CardDropBonus2"))
End Function
