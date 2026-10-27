Attribute VB_Name = "CustomScenarios"
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

Public Enum e_DamageSourceType
    e_phisical
    e_magic
    e_pet
    e_dot
    e_trap
End Enum

Public Enum e_ScenarioInterfaces
    eInventoryInterface = 1
End Enum

Public Type e_Rank
    PlayerIndex As Integer
    Score As Integer
End Type

Private CustomScenarioList   As New Dictionary
Private ScenarioUpdateList() As IBaseScenario
Private AvailableUpdateSlots As t_IndexHeap
Private ActiveUpdateSlots    As t_IndexHeap
Const InitialUpdateSize = 20

Private Sub InitializeUpdateStacks()
    ReDim ScenarioUpdateList(InitialUpdateSize) As IBaseScenario
    ReDim AvailableUpdateSlots.IndexInfo(InitialUpdateSize)
    ReDim ActiveUpdateSlots.IndexInfo(InitialUpdateSize)
    Dim i As Integer
    For i = 1 To InitialUpdateSize
        AvailableUpdateSlots.IndexInfo(i) = InitialUpdateSize - (i - 1)
    Next i
    AvailableUpdateSlots.currentIndex = InitialUpdateSize
    ActiveUpdateSlots.currentIndex = 0
End Sub

Private Sub IncreaseArraySize(ByVal ExtraSlots As Integer)
    Dim NewSize As Integer
    NewSize = UBound(ScenarioUpdateList) + ExtraSlots
    ReDim Preserve ScenarioUpdateList(NewSize) As IBaseScenario
    ReDim Preserve AvailableUpdateSlots.IndexInfo(NewSize)
    ReDim Preserve ActiveUpdateSlots.IndexInfo(NewSize)
    Dim i As Integer
    For i = 1 To ExtraSlots
        AvailableUpdateSlots.IndexInfo(i) = NewSize - (i - 1)
    Next i
    AvailableUpdateSlots.currentIndex = ExtraSlots
End Sub

Public Function GetMap(ByVal mapIndex As Integer) As IBaseScenario
    On Error GoTo GetMap_Err:
    Set GetMap = Nothing
    If CustomScenarioList.Exists(mapIndex) Then
        Set GetMap = CustomScenarioList.Item(mapIndex)
    End If
    Exit Function
GetMap_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.GetMap " & mapIndex, Erl)
End Function

Public Function SetCustomScenario(ByVal mapIndex As Integer, ByRef Scenario As IBaseScenario) As Boolean
    ' Plan 05.002 wave 5. Dictionary.Add RAISES error 457 when the key already exists, and
    ' this Sub had no handler: the error escaped into the Class_Initialize / SetLobbyIndex of
    ' the scenario being registered and aborted the rest of its initialization. It now fails
    ' soft and LOUD, and returns whether it registered so a caller can abort. The scenarios
    ' call it with Call, so they keep compiling untouched.
    On Error GoTo SetCustomScenario_Err:
    If CustomScenarioList.Exists(mapIndex) Then
        Call LogError("CustomScenarios.SetCustomScenario: el mapa " & mapIndex & " ya tiene un escenario registrado. No se pisa: el escenario nuevo se queda sin ese mapa.")
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[Eventos] El mapa " & mapIndex & " ya tiene un escenario registrado: el evento nuevo no puede usarlo. Revisa el log.", e_TextChannel.TEXTCHANNEL_SERVER_STAFF, e_FontTypeNames.FONTTYPE_INFO))
        Exit Function
    End If
    Call CustomScenarioList.Add(mapIndex, Scenario)
    SetCustomScenario = True
    Exit Function
SetCustomScenario_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.SetCustomScenario " & mapIndex, Erl)
End Function

Public Sub RemoveCustomScenario(ByVal mapIndex As Integer)
    ' Plan 05.002 wave 5. Dictionary.Remove RAISES error 32811 when the key is missing. In
    ' every ClearMap this call is the FIRST line and RemoveUpdateScenario the LAST, so the
    ' error jumped to the ClearMap handler and the tick slot was never released. That always
    ' happens when the map was never registered (instance pool exhausted => MapNumber = -1).
    On Error GoTo RemoveCustomScenario_Err:
    If Not CustomScenarioList.Exists(mapIndex) Then Exit Sub
    Call CustomScenarioList.Remove(mapIndex)
    Exit Sub
RemoveCustomScenario_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.RemoveCustomScenario " & mapIndex, Erl)
End Sub

Public Function AddUpdateScenario(ByRef Scenario As IBaseScenario) As Integer
    On Error GoTo AddUpdateScenario_Err:
    Dim pos As Integer
    If AvailableUpdateSlots.currentIndex = 0 And ActiveUpdateSlots.currentIndex = 0 Then
        Call InitializeUpdateStacks
    End If
    If AvailableUpdateSlots.currentIndex = 0 Then
        Call IncreaseArraySize(InitialUpdateSize)
    End If
    pos = AvailableUpdateSlots.IndexInfo(AvailableUpdateSlots.currentIndex)
    AvailableUpdateSlots.currentIndex = AvailableUpdateSlots.currentIndex - 1
    ActiveUpdateSlots.IndexInfo(ActiveUpdateSlots.currentIndex) = pos
    ActiveUpdateSlots.currentIndex = ActiveUpdateSlots.currentIndex + 1
    Set ScenarioUpdateList(pos) = Scenario
    AddUpdateScenario = pos
    Exit Function
AddUpdateScenario_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.AddUpdateScenario", Erl)
End Function

Public Sub RemoveUpdateScenario(ByVal Index As Integer)
    ' Plan 05.002 wave 5. Hardened, not rewritten: the linear sweep (with the Exit For that
    ' main already has) keeps this idempotent. What was missing:
    '  - an own handler and a REAL range check: Debug.Assert is compiled out in release, so
    '    an out of range Index raised an unhandled error 9 and aborted the scenario close.
    '  - an overflow guard when giving the slot back (same one InstanceManager.ReleaseInstance
    '    had to learn).
    ' Index is now ByVal: every caller already passes it parenthesized (by value).
    On Error GoTo RemoveUpdateScenario_Err:
    If Index < LBound(ScenarioUpdateList) Or Index > UBound(ScenarioUpdateList) Then
        Call LogError("CustomScenarios.RemoveUpdateScenario: slot fuera de rango (" & Index & ")")
        Exit Sub
    End If
    Set ScenarioUpdateList(Index) = Nothing
    Dim i As Integer
    For i = 0 To ActiveUpdateSlots.currentIndex - 1
        If ActiveUpdateSlots.IndexInfo(i) = Index Then
            ActiveUpdateSlots.IndexInfo(i) = ActiveUpdateSlots.IndexInfo(ActiveUpdateSlots.currentIndex - 1)
            ActiveUpdateSlots.currentIndex = ActiveUpdateSlots.currentIndex - 1
            If AvailableUpdateSlots.currentIndex < UBound(AvailableUpdateSlots.IndexInfo) Then
                AvailableUpdateSlots.currentIndex = AvailableUpdateSlots.currentIndex + 1
                AvailableUpdateSlots.IndexInfo(AvailableUpdateSlots.currentIndex) = Index
            Else
                Call LogError("CustomScenarios.RemoveUpdateScenario: el heap de slots disponibles esta lleno, se descarta el slot " & Index)
            End If
            Exit For
        End If
    Next
    Exit Sub
RemoveUpdateScenario_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.RemoveUpdateScenario " & Index, Erl)
End Sub

Public Sub UpdateAll()
    ' Plan 05.002 wave 5. TWO problems lived in this loop:
    ' 1) A single On Error for the whole loop: the first scenario that raised aborted the tick
    '    of ALL the following ones, and since the For always starts at 0 it starved them
    '    forever (UpdateAll runs every 40 ms from modGameEvents.MaybeRunGameEvents).
    ' 2) ActiveUpdateSlots was iterated WHILE the scenarios mutate it: they call
    '    CloseScenario -> RemoveUpdateScenario from their own Update, and that compacts the
    '    heap by swapping with the last element. VB6 evaluates the For bound once, so the stale
    '    position was re-read and the moved scenario got two Update calls in the same tick.
    ' So the slot list is COPIED before walking it, and each Update runs in its own Sub, the
    ' only way in VB6 to get one handler per iteration without Resume Next.
    On Error GoTo UpdateAll_Err:
    Dim i         As Integer
    Dim SlotCount As Integer
    Dim Slots()   As Integer
    SlotCount = ActiveUpdateSlots.currentIndex
    If SlotCount <= 0 Then Exit Sub
    ReDim Slots(0 To SlotCount - 1) As Integer
    For i = 0 To SlotCount - 1
        Slots(i) = ActiveUpdateSlots.IndexInfo(i)
    Next i
    For i = 0 To SlotCount - 1
        Call UpdateSingleScenario(Slots(i))
    Next i
    Exit Sub
UpdateAll_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UpdateAll", Erl)
End Sub

Private Sub UpdateSingleScenario(ByVal Slot As Integer)
    ' The Is Nothing check is NOT redundant: it covers a scenario that was taken down during
    ' this very tick, because RemoveUpdateScenario leaves the slot Nothing before compacting.
    On Error GoTo UpdateSingleScenario_Err:
    If Slot < LBound(ScenarioUpdateList) Or Slot > UBound(ScenarioUpdateList) Then Exit Sub
    If ScenarioUpdateList(Slot) Is Nothing Then Exit Sub
    Call ScenarioUpdateList(Slot).Update
    Exit Sub
UpdateSingleScenario_Err:
    ' The error is NOT swallowed: it is logged with the slot to identify the scenario.
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UpdateAll slot:" & Slot, Erl)
End Sub

Public Sub Reset(ByVal MapNumber As Integer)
    On Error GoTo Reset_Err:
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(MapNumber)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.Reset
    Exit Sub
Reset_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.Reset " & MapNumber, Erl)
End Sub

Public Sub PlayerKillNpc(ByVal MapNumber As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
    On Error GoTo PlayerKillNpc_Err:
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(MapNumber)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.PlayerKillNpc(NpcIndex, UserIndex, SourceType, SourceIndex)
    Exit Sub
PlayerKillNpc_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.PlayerKillNpc map:" & MapNumber & " npc: " & NpcIndex & " user: " & UserIndex, Erl)
End Sub

Public Sub NPcKillPlayer(ByVal MapNumber As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
    On Error GoTo NPcKillPlayer_Err:
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(MapNumber)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.NPcKillPlayer(NpcIndex, UserIndex, SourceType, SourceIndex)
    Exit Sub
NPcKillPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.NPcKillPlayer", Erl)
End Sub

Public Sub PlayerKillPlayer(ByVal MapNumber As Integer, _
                            ByVal killerIndex As Integer, _
                            ByVal deadIndex As Integer, _
                            ByVal SourceType As e_DamageSourceType, _
                            ByVal SourceIndex As Integer)
    On Error GoTo PlayerKillPlayer_Err:
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(MapNumber)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.PlayerKillPlayer(killerIndex, deadIndex, SourceType, SourceIndex)
    Exit Sub
PlayerKillPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.PlayerKillPlayer", Erl)
End Sub

Public Sub UserDie(ByVal UserIndex As Integer)
    On Error GoTo UserDie_Err:
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(UserList(UserIndex).pos.Map)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.UserDie(UserIndex)
    Exit Sub
UserDie_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UserDie", Erl)
End Sub

Public Sub NpcDie(ByVal NpcIndex As Integer)
    On Error GoTo NpcDie_Err:
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(NpcList(NpcIndex).pos.Map)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.NpcDie(NpcIndex)
    Exit Sub
NpcDie_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.NpcDie", Erl)
End Sub

'inform that user want to pickup an item, return false if user can't pick up item at pos
Public Function UserCanPickUpItem(ByVal UserIndex As Integer) As Boolean
    On Error GoTo UserPickUpItem_Err:
    UserCanPickUpItem = True
    Dim Scenario           As IBaseScenario
    Dim InventoryInterface As IInventoryInterface
    Set Scenario = GetMap(UserList(UserIndex).pos.Map)
    If Scenario Is Nothing Then
        Exit Function
    End If
    If Not IsSet(Scenario.GetExtraInterfaces(), eInventoryInterface) Then
        Exit Function
    End If
    Set InventoryInterface = Scenario
    UserCanPickUpItem = InventoryInterface.UserCanPickUpItem(UserIndex)
    Exit Function
UserPickUpItem_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UserCanPickUpItem", Erl)
End Function

Public Sub UserDidPickupItem(ByVal UserIndex As Integer, ByVal ItemId As Integer)
    On Error GoTo UserDidPickupItem_Err:
    Dim Scenario           As IBaseScenario
    Dim InventoryInterface As IInventoryInterface
    Set Scenario = GetMap(UserList(UserIndex).pos.Map)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    If Not IsSet(Scenario.GetExtraInterfaces(), eInventoryInterface) Then
        Exit Sub
    End If
    Set InventoryInterface = Scenario
    Call InventoryInterface.UserDidPickupItem(UserIndex, ItemId)
    Exit Sub
UserDidPickupItem_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UserDidPickupItem", Erl)
End Sub

'inform that user want to drop an item
Public Sub UserDropItem(ByVal UserIndex As Integer, ByVal Slot As Integer, ByVal Map As Integer, ByVal TileX As Integer, ByVal TileY As Integer)
    On Error GoTo UserDropItem_Err:
    Dim Scenario           As IBaseScenario
    Dim InventoryInterface As IInventoryInterface
    Set Scenario = GetMap(Map)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    If Not IsSet(Scenario.GetExtraInterfaces(), eInventoryInterface) Then
        Exit Sub
    End If
    Set InventoryInterface = Scenario
    Call InventoryInterface.UserDropItem(UserIndex, Slot, Map, TileX, TileY)
    Exit Sub
UserDropItem_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UserDropItem", Erl)
End Sub

Public Function UserCanDropItem(ByVal UserIndex As Integer, ByVal Slot As Integer, ByVal Map As Integer, ByVal TileX As Integer, ByVal TileY As Integer) As Boolean
    On Error GoTo UserCanDropItem_Err:
    UserCanDropItem = True
    Dim Scenario           As IBaseScenario
    Dim InventoryInterface As IInventoryInterface
    Set Scenario = GetMap(Map)
    If Scenario Is Nothing Then
        Exit Function
    End If
    If Not IsSet(Scenario.GetExtraInterfaces(), eInventoryInterface) Then
        Exit Function
    End If
    Set InventoryInterface = Scenario
    UserCanDropItem = InventoryInterface.UserCanDropItemAt(UserIndex, Slot, Map, TileX, TileY)
    Exit Function
UserCanDropItem_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UserCanDropItem", Erl)
End Function

Public Function PrepareNewEvent(ByVal eventType As e_EventType, ByVal LobbyIndex As Integer) As Boolean
    ' Plan 05.002 wave 5. Returns True only if the lobby ended up with a LIVE scenario
    ' (built AND hooked to the tick). It used to be a Sub that verified nothing.
    '  - e_EventType.Generic (0) never had a Case. The lobby kept Scenario = Nothing and,
    '    since UpdateWaitingForPlayers is only called from the scenarios' Update, nobody
    '    advanced it: it never started, never expired and never cancelled, and it sat in
    '    AcceptingPlayers charging inscriptions while ActiveEventType() counted it as an
    '    active event, blocking every later public event.
    '  - The scenarios' Class_Initialize trap their own errors, so New ALWAYS returns an
    '    object, well built or not. One that failed before AddUpdateScenario never ticks:
    '    the same zombie lobby through another door.
    ' So the POSTCONDITION (it is in the tick) is verified instead of trusting the New.
    On Error GoTo PrepareNewEvent_Err:
    If LobbyIndex < LBound(LobbyList) Or LobbyIndex > UBound(LobbyList) Then
        Call LogError("CustomScenarios.PrepareNewEvent: lobby fuera de rango (" & LobbyIndex & ").")
        Exit Function
    End If
    Set LobbyList(LobbyIndex).Scenario = Nothing
    Select Case eventType
        Case e_EventType.NpcHunt
            Set LobbyList(LobbyIndex).Scenario = New ScenarioHunt
        Case e_EventType.DeathMatch
            Set LobbyList(LobbyIndex).Scenario = New ScenarioDeathMatch
        Case e_EventType.NavalBattle
            Set LobbyList(LobbyIndex).Scenario = New ScenarioNavalBoarding
        Case e_EventType.Generic
            ' No scenario ON PURPOSE: this is the MANUAL event of the GM (/CREAREVENTO LOBBY),
            ' started and ended by hand. ModLobby.UpdateLobbiesWithoutScenario advances it.
            ' Nothing to verify, so it reports success.
            PrepareNewEvent = True
            Exit Function
        Case Else
            ' CaptureTheFlag does not come through here yet (it keeps its own singleton until
            ' it is migrated to IBaseScenario), and any other value is a bug in the caller.
            Call LogError("CustomScenarios.PrepareNewEvent: tipo de evento no soportado (" & eventType & ") para el lobby " & LobbyIndex & ".")
            Exit Function
    End Select
    If LobbyList(LobbyIndex).Scenario Is Nothing Then Exit Function
    LobbyList(LobbyIndex).Scenario.SetLobbyIndex (LobbyIndex)
    ' Postcondition: Class_Initialize hooks the tick (AddUpdateScenario) only AFTER
    ' SetCustomScenario succeeded (plan 27.001). If either failed (duplicate key, missing
    ' .ini, an error swallowed by the constructor) the object exists but is NOT in the tick.
    If Not IsScenarioInUpdateList(LobbyList(LobbyIndex).Scenario) Then
        Call LogError("CustomScenarios.PrepareNewEvent: el escenario tipo " & eventType & " del lobby " & LobbyIndex & " no quedo registrado en el tick. Se descarta.")
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[Eventos] El escenario no pudo inicializarse (ver log). El evento no se abre.", e_TextChannel.TEXTCHANNEL_SERVER_STAFF, e_FontTypeNames.FONTTYPE_INFO))
        ' Plan 27.001: the constructor already took its map instance(s) from the pool before it
        ' failed. Reset gives them back (ClearMap + ReleaseInstance, all guarded by MapNumber > 0)
        ' and is safe on a half-built scenario: it swallows its own errors and RemoveUpdateScenario
        ' ignores slot 0. Without it every failed open leaked an instance map until restart.
        Call LobbyList(LobbyIndex).Scenario.Reset
        Set LobbyList(LobbyIndex).Scenario = Nothing
        Exit Function
    End If
    PrepareNewEvent = True
    Exit Function
PrepareNewEvent_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.PrepareNewEvent", Erl)
End Function

Public Function IsScenarioInUpdateList(ByRef Scenario As IBaseScenario) As Boolean
    ' Both sides are IBaseScenario, so Is compares the same interface pointer.
    On Error GoTo IsScenarioInUpdateList_Err:
    Dim i As Integer
    If Scenario Is Nothing Then Exit Function
    For i = 0 To ActiveUpdateSlots.currentIndex - 1
        If ScenarioUpdateList(ActiveUpdateSlots.IndexInfo(i)) Is Scenario Then
            IsScenarioInUpdateList = True
            Exit Function
        End If
    Next i
    Exit Function
IsScenarioInUpdateList_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.IsScenarioInUpdateList", Erl)
End Function

Public Sub ClearMap(ByVal MapNumber As Integer)
    On Error GoTo ClearMap_Err:
    Dim x As Long
    Dim y As Long
    For y = 1 To 99
        For x = 1 To 99
            If MapData(x, y, MapNumber).ObjInfo.ObjIndex > 0 Then
                If ItemNoEsDeMapa(MapData(x, y, MapNumber).ObjInfo.ObjIndex) Then
                    Call EraseObj(GetMaxInvOBJ(), MapNumber, x, y)
                End If
            End If
        Next x
    Next y
    Exit Sub
ClearMap_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.ClearMap", Erl)
    Resume Next
End Sub

Public Function IsEventActive() As Boolean
    ' Plan 05.002 wave 3: single source of truth, derived from what is alive.
    IsEventActive = (ModLobby.ActiveEventType() <> ModLobby.NO_ACTIVE_EVENT)
End Function

' Plan 27.001 (05.002 wave 7): "the map this player stands on belongs to a scenario".
' Deliberately NOT map numbers: EsMapaEvento() is a static list from MapasEspeciales.dat and events
' run on instances with a runtime-assigned number, so a gate based on it is blind for 3 of 4 modes.
' Not MapInfo(...).Seguro either: scenarios toggle it per phase (1 -> 0 -> 1) and the gate would
' blink. The per-map registry is how the engine already routes death, flags and NPCs.
Public Function IsUserInScenarioMap(ByVal UserIndex As Integer) As Boolean
    On Error GoTo IsUserInScenarioMap_Err
    IsUserInScenarioMap = False
    If UserIndex <= 0 Then Exit Function
    IsUserInScenarioMap = Not (GetMap(UserList(UserIndex).pos.Map) Is Nothing)
    Exit Function
IsUserInScenarioMap_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.IsUserInScenarioMap", Erl)
    IsUserInScenarioMap = False
End Function

' Duel = per-user flag (a duel does not live in a scenario). Event = the map the player stands on.
' The handler returns False on purpose: this is a gameplay restriction, not a security boundary. If
' something blows up it is preferable to let people work or trade than to lock the whole server's
' professions. The error is logged anyway.
Public Function IsUserInMatch(ByVal UserIndex As Integer) As Boolean
    On Error GoTo IsUserInMatch_Err
    IsUserInMatch = False
    If UserIndex <= 0 Then Exit Function
    If UserList(UserIndex).flags.EnReto Then
        IsUserInMatch = True
        Exit Function
    End If
    IsUserInMatch = IsUserInScenarioMap(UserIndex)
    Exit Function
IsUserInMatch_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.IsUserInMatch", Erl)
    IsUserInMatch = False
End Function

Public Sub UserDisconnected(ByVal UserIndex As Integer)
    Call RegisterDisconnectedUser(UserIndex)
End Sub

Public Sub UserConnected(ByVal UserIndex As Integer)
    Call RegisterReconnectedUser(UserIndex)
End Sub

Public Sub GetNextWaypointForNpc(ByVal NpcIndex As Integer, ByRef PosX As Integer, ByRef PosY As Integer)
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(NpcList(NpcIndex).pos.Map)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.GetNextWaypointForNpc(NpcIndex, PosX, PosY)
    Exit Sub
End Sub
