VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScenarioHunt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Const RequiredWeaponId = 1847

Const configPath = "../dat/Scenarios/NpcHunt.ini"
Const RequireItemMessage = 403
Const MillisecondInSecond = 1000
Implements IBaseScenario
Dim MatchNpc           As New Dictionary
Dim SpawnCreatureCount As Integer
Dim ElapsedTime        As Long
Dim LastFrameTime      As Long
Dim LastScoreUpdate    As Long
Dim ActiveNpc          As Integer
Dim SpawnedCreaturesId As New Dictionary
Dim MapNumber          As Integer
Dim UpdateIndex        As Integer
Dim MaxTime            As Long
Dim TargetScore        As Integer
Dim BroadCastInterval  As Long
Dim LastBroadCastTime  As Long
Dim frametime          As Long
Dim Board              As New ScoreBoard
Dim StartTimer         As t_Timer
Dim PrevMapState       As Integer
Dim MapDropItems       As Boolean
Dim PlayerRespawn      As New Dictionary
Dim PlayerRespawnTime  As Long
Dim LobbyIndex         As Integer

Friend Sub RegisterNpc(ByVal Tag, ByVal sectionSize, ByRef registry As Dictionary, ByRef File As clsIniManager)
    Dim i        As Integer
    Dim npcEntry As EventNpcDefinition
    For i = 1 To sectionSize
        Set npcEntry = New EventNpcDefinition
        npcEntry.NpcId = val(File.GetValue(Tag & i, "NPCID"))
        npcEntry.SpawnCount = val(File.GetValue(Tag & i, "count"))
        npcEntry.Score = val(File.GetValue(Tag & i, "score"))
        If Not registry.Exists(npcEntry.NpcId) Then
            Call registry.Add(npcEntry.NpcId, npcEntry)
        End If
    Next i
End Sub

Private Sub ConfigMapLobby()
    PrevMapState = MapInfo(MapNumber).Seguro
    MapDropItems = MapInfo(MapNumber).DropItems
    MapInfo(MapNumber).Seguro = 1
    MapInfo(MapNumber).DropItems = False
End Sub

Private Sub Class_Initialize()
    On Error GoTo Hunt_Initialize_Err:
    ' Plan 27.001: prime the frame clock; Update runs before StartGame and 0 would overflow.
    LastFrameTime = GetTickCountRaw()
    Dim validNpcCount As Integer
    Dim ScenarioIni   As clsIniManager
    Set ScenarioIni = New clsIniManager
    Debug.Assert FileExist(DatPath & configPath)
    Dim OriginalMap As Integer
    Call ScenarioIni.Initialize(DatPath & configPath)
    validNpcCount = val(ScenarioIni.GetValue("INIT", "TargetNpc"))
    OriginalMap = val(ScenarioIni.GetValue("INIT", "MapNumber"))
    MapNumber = InstanceManager.GetNextAvailableInstance
    Call InstanceManager.CloneMap(OriginalMap, MapNumber)
    MaxTime = val(ScenarioIni.GetValue("INIT", "MaxTime"))
    MaxTime = MaxTime * 1000
    PlayerRespawnTime = val(ScenarioIni.GetValue("INIT", "PlayerRespawnTime"))
    PlayerRespawnTime = PlayerRespawnTime * 1000
    TargetScore = val(ScenarioIni.GetValue("INIT", "TargetScore"))
    Call RegisterNpc("TARGET", validNpcCount, MatchNpc, ScenarioIni)
    ' Plan 27.001: register the map FIRST and hook the tick only on success, so a failed
    ' SetCustomScenario leaves UpdateIndex = 0 and PrepareNewEvent's postcondition sees it.
    If CustomScenarios.SetCustomScenario(MapNumber, Me) Then
        UpdateIndex = AddUpdateScenario(Me)
    End If
    LastBroadCastTime = 0
    Exit Sub
Hunt_Initialize_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.Initialize ", Erl)
End Sub

Private Sub CloseScenario()
    On Error GoTo CloseScenario_Err:
    Call ModLobby.ReturnAllPlayers(LobbyList(LobbyIndex))
    Set LobbyList(LobbyIndex).Scenario = Nothing
    Call ModLobby.UpdateLobbyState(LobbyList(LobbyIndex), e_LobbyState.UnInitilized)
    Call ClearMap
    If MapNumber > 0 Then
        Call ReleaseLobby(LobbyIndex)
        Call ReleaseInstance(MapNumber)
        MapNumber = 0
    End If
    Exit Sub
CloseScenario_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.CloseScenario ", Erl)
End Sub

Private Sub ClearMap()
    On Error GoTo ClearMap_Err:
    CustomScenarios.RemoveCustomScenario (MapNumber)
    CustomScenarios.RemoveUpdateScenario (UpdateIndex)
    UpdateIndex = 0
    MapInfo(MapNumber).Seguro = PrevMapState
    MapInfo(MapNumber).DropItems = MapDropItems
    Exit Sub
ClearMap_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.ClearMap", Erl)
End Sub

Public Sub IBaseScenario_SetLobbyIndex(ByVal value As Integer)
    LobbyIndex = value
    Call ConfigMapLobby
End Sub

Public Function IBaseScenario_GetScenarioName() As String
    IBaseScenario_GetScenarioName = "Caceria"
End Function

Public Sub IBaseScenario_BroadcastOpenScenario()
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MsgHuntScenearioIntro, "", e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
End Sub

Public Sub IBaseScenario_SendRules(ByVal UserIndex As Integer)
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageLocaleMsg(MsgHuntScenarioInstructions, "", e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
End Sub

Public Function IBaseScenario_GetExtraInterfaces() As Long
    IBaseScenario_GetExtraInterfaces = 0
End Function

Public Sub IBaseScenario_Reset()
    ' Plan 05.002 wave 5 (Reset contract). The ONLY caller is ModLobby.ForceReset
    ' (/lobby forcereset), which ALREADY returned and refunded the players and left
    ' RegisteredPlayers at 0 BEFORE calling here. So the lobby list cannot be iterated here.
    ' What belongs to the scenario and nobody did: kill the spawned creatures and give the
    ' instance map back to the pool. ReleaseLobby does NOT go here: giving the pool slot back
    ' is ForceReset's job.
    On Error GoTo Reset_Err
    Dim key As Variant
    For Each key In SpawnedCreaturesId.Keys
        Call QuitarNPC(key, eClearHunt)
    Next key
    Call SpawnedCreaturesId.RemoveAll
    Call PlayerRespawn.RemoveAll
    Call ClearMap
    If MapNumber > 0 Then
        Call ReleaseInstance(MapNumber)
        MapNumber = 0
    End If
    Exit Sub
Reset_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.Reset", Erl)
End Sub

Private Function IBaseScenario_RequiresSpawn() As Boolean
    IBaseScenario_RequiresSpawn = False
End Function

Private Function IBaseScenario_ValidateUser(ByVal UserIndex As Integer) As Integer
    On Error GoTo ValidateUser_Err:
    Dim i As Integer
    With UserList(UserIndex)
        For i = 1 To UBound(.invent.Object)
            If .invent.Object(i).ObjIndex = RequiredWeaponId Then
                IBaseScenario_ValidateUser = 0
                Exit Function
            End If
        Next i
        IBaseScenario_ValidateUser = RequireItemMessage
    End With
    Exit Function
ValidateUser_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.ValidateUser userIndex: " & UserIndex, Erl)
End Function

Private Sub IBaseScenario_WillSummonPlayer(ByVal UserIndex As Integer)
    Dim SpawnPos As t_WorldPos
    SpawnPos.Map = MapNumber
    SpawnPos.x = RandomNumber(12, 85)
    SpawnPos.y = RandomNumber(12, 85)
    LobbyList(LobbyIndex).SummonCoordinates = SpawnPos
    Call Board.AddPlayer(UserIndex)
End Sub

Private Sub IBaseScenario_UpdateLobbyState(ByVal oldState As e_LobbyState, ByRef NewState As e_LobbyState)
    If oldState = NewState Then
        Exit Sub
    End If
    If NewState = InProgress Then
        Call SetTimer(StartTimer, MillisecondInSecond)
    End If
    If NewState = Closed Then
        Call MatchComplete
    End If
End Sub

Private Sub StartGame()
    Dim key As Variant
    Dim i   As Integer
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        Dim SpawnPos As t_WorldPos
        SpawnPos.Map = MapNumber
        SpawnPos.x = RandomNumber(12, 85)
        SpawnPos.y = RandomNumber(12, 85)
        If LobbyList(LobbyIndex).Players(i).IsSummoned Then
            If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
                Call WarpToLegalPos(LobbyList(LobbyIndex).Players(i).User.ArrayIndex, MapNumber, SpawnPos.x, SpawnPos.y, True, True)
            End If
        End If
    Next i
    For Each key In MatchNpc.Keys
        For i = 1 To MatchNpc(key).SpawnCount
            SpawnNpc (MatchNpc(key).NpcId)
        Next i
    Next key
    LastFrameTime = GetTickCountRaw()
    BroadCastInterval = 60000
    Call SendData(SendTarget.toMap, MapNumber, PrepareMessageLocaleMsg(MSG_INICIE_CACERA, "", e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) 'Msg1730=¡Que inicie la cacería!
End Sub

Private Sub IBaseScenario_PlayerKillNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
    If LobbyList(LobbyIndex).State <> e_LobbyState.InProgress Then
        Exit Sub
    End If
    Dim npcType As Integer
    npcType = NpcList(NpcIndex).Numero
    If SourceType = e_phisical And SourceIndex = RequiredWeaponId Then
        If MatchNpc.Exists(npcType) Then
            If Board.UpdatePlayerScore(UserIndex, MatchNpc.Item(npcType).Score) >= TargetScore Then
                LobbyList(LobbyIndex).State = Completed
                Call SendData(SendTarget.toMap, MapNumber, PrepareMessageLocaleMsg(MSG_TENEMOS_GANADOR, vbNullString, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) 'Msg1731=¡Tenemos un ganador!
                Call MatchComplete
                Exit Sub
            End If
        End If
    Else
        Call WriteLocaleMsg(UserIndex, MSG_MUST_HIT_WITH_WEAPON_TO_SCORE_POINTS, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, ObjData(RequiredWeaponId).name)
    End If
    ' Plan 05.002 wave 5: Scripting.Dictionary.Remove on a missing key raises 32811 and this
    ' handler has no On Error of its own: the error went up to CustomScenarios.PlayerKillNpc,
    ' which swallows it, and the SpawnNpc below NEVER ran (the creature was not replaced).
    ' The hook fires for ANY npc killed by a user on the map, not only the hunt ones: pets that
    ' WarpMascotas re-summons on the instance map, summons, GM-created npcs. SpawnNpc goes
    ' INSIDE the If on purpose: replacing a creature for the death of someone else's pet would
    ' spawn the pet as prey.
    If SpawnedCreaturesId.Exists(NpcIndex) Then
        Call SpawnedCreaturesId.Remove(NpcIndex)
        Call SpawnNpc(npcType)
    End If
End Sub

Private Sub IBaseScenario_NPcKillPlayer(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
    'implements interface
End Sub

Private Sub IBaseScenario_PlayerKillPlayer(ByVal killerIndex As Integer, ByVal deadIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
    'implements interface
End Sub

Private Sub IBaseScenario_Update()
    Dim nowRaw    As Long
    Dim frametime As Long
    nowRaw = GetTickCountRaw()
    frametime = TicksElapsedLong(LastFrameTime, nowRaw)
    LastFrameTime = nowRaw
    If LobbyList(LobbyIndex).State = e_LobbyState.InProgress Then
        If StartTimer.Occurrences < 10 Then
            If UpdateTime(StartTimer, frametime) Then
                If StartTimer.Occurrences >= 10 Then
                    Call StartGame
                Else
                    Call SendData(SendTarget.toMap, MapNumber, PrepareMessageLocaleMsg(MSG_CACERA_INICIARA, 10 - StartTimer.Occurrences, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) ' Msg1819=La cacería iniciará en ¬1
                End If
            End If
            Exit Sub
        End If
        ElapsedTime = ElapsedTime + frametime
        If ElapsedTime - LastScoreUpdate > 30000 Then
            LastScoreUpdate = ElapsedTime
            Call Board.PublishScoreboard(SendTarget.toMap, MapNumber, LobbyList(LobbyIndex), False)
        End If
        If ElapsedTime > MaxTime Then
            Call SendData(SendTarget.toMap, MapNumber, PrepareMessageLocaleMsg(MSG_ACABA_TIEMPO, vbNullString, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) ' Msg1820=¡Se acabó el tiempo!
            Call MatchComplete
        End If
        Call RespawnPlayers
        Call BroadCastTimeUpdate
    ElseIf LobbyList(LobbyIndex).State = e_LobbyState.Completed Then
        ElapsedTime = ElapsedTime + frametime
        If ElapsedTime - LastScoreUpdate > 1000 Then
            LastScoreUpdate = ElapsedTime
            Call CloseCountDown
        End If
    ElseIf LobbyList(LobbyIndex).State = e_LobbyState.AcceptingPlayers Then
        Call UpdateWaitingForPlayers(frametime, LobbyList(LobbyIndex))
    ElseIf LobbyList(LobbyIndex).State = e_LobbyState.Closed Then
        Call CloseScenario
    End If
End Sub

Public Sub IBaseScenario_OnUserDisconnected(ByVal UserIndex As Integer)
    Call Board.DisconnectUser(UserIndex, UserList(UserIndex).Id)
End Sub

Public Sub IBaseScenario_OnUserReconnect(ByVal UserIndex As Integer)
    Call Board.RestoreDisconnectUser(UserIndex, UserList(UserIndex).Id)
End Sub

Private Sub MatchComplete()
    On Error GoTo MatchComplete_Err:
    Dim i As Integer
    If Not LobbyList(LobbyIndex).Canceled Then
        Call Board.PublishScoreboard(SendTarget.toMap, MapNumber, LobbyList(LobbyIndex), True)
        Call GiveRewards
    End If
    LobbyList(LobbyIndex).State = Completed
    Call CustomScenarios.ClearMap(MapNumber)
    ElapsedTime = MaxTime
    Dim key As Variant
    For Each key In SpawnedCreaturesId.Keys
        Call QuitarNPC(key, eClearHunt)
    Next key
    ' Plan 05.002 wave 5: empty the registry AFTER killing them. MatchComplete runs more than
    ' once (end by score or time, then CancelLobby -> UpdateLobbyState(Closed)) and NpcList
    ' indices are recycled (ReleaseNpc returns them to the IdNpcLibres heap). Without this the
    ' second pass does QuitarNPC on NPCs of OTHER maps that inherited those indices: ReleaseNpc
    ' only stops the double release of the SAME npc, via flags.NPCActive.
    Call SpawnedCreaturesId.RemoveAll
    Exit Sub
MatchComplete_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.MatchComplete ", Erl)
    Resume Next
End Sub

Private Sub SpawnNpc(ByVal npcType As Integer)
    Dim SpawnPos As t_WorldPos
    Dim NpcId    As Integer
    SpawnPos.Map = MapNumber
    SpawnPos.x = RandomNumber(12, 85)
    SpawnPos.y = RandomNumber(12, 85)
    NpcId = CrearNPC(npcType, MapNumber, SpawnPos)
    If NpcId > 0 Then
        NpcList(NpcId).flags.Respawn = 1
        NpcList(NpcId).Contadores.IntervaloRespawn = 0
        Call SpawnedCreaturesId.Add(NpcId, 0)
    Else
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageLocaleMsg(MSG_NPC_SPAWN_FAILED, npcType, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) ' Msg1786=No se pudo spawnear el npc ¬1, revisar.
        Call TraceError(Err.Number, "failed to spawn Npc " & npcType, "ScenarioHunt.SpawnNpc ", Erl)
    End If
End Sub

Private Sub CloseCountDown()
    Dim pendingTime As Integer
    pendingTime = 10 - (ElapsedTime - MaxTime) / 1000
    If pendingTime <= 0 Then
        Call CloseScenario
        Exit Sub
    End If
    Call SendData(SendTarget.toMap, MapNumber, PrepareMessageLocaleMsg(MSG_MAPA_CERRARA_SEGUNDOS, pendingTime, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) ' Msg1821=El mapa se cerrará en ¬1 segundos.
End Sub

Private Sub BroadCastTimeUpdate()
    If BroadCastInterval > (ElapsedTime - LastBroadCastTime) Then
        Exit Sub
    End If
    LastBroadCastTime = ElapsedTime
    Dim pendingTime As Long
    ' Plan 05.002 wave 5: this added `frametime` "to round up", but that was the MODULE level
    ' variable, which nobody ever writes (IBaseScenario_Update declares its own local Dim
    ' frametime that shadows it). It always added 0. The phantom term is removed so the code
    ' says what it does.
    pendingTime = (MaxTime - ElapsedTime) / 1000
    If pendingTime > MaxTime / 2000 Then
        Exit Sub
    ElseIf pendingTime >= 60 Then
        pendingTime = pendingTime / 60
        Call SendData(SendTarget.toMap, MapNumber, PrepareMessageLocaleMsg(MSG_QUEDAN_MINUTOS, pendingTime, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) ' Msg1822=Quedan ¬1 minutos.
        BroadCastInterval = 60000
    ElseIf pendingTime <= 10 Then
        Call SendData(SendTarget.toMap, MapNumber, PrepareMessageConsoleMsg(pendingTime, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
    End If
    If pendingTime <= 1 Then
        BroadCastInterval = 1000
    End If
End Sub

Public Sub IBaseScenario_GetNextWaypointForNpc(ByVal NpcIndex As Integer, ByRef PosX As Integer, ByRef PosY As Integer)
End Sub

Public Sub IBaseScenario_NpcDie(ByVal NpcIndex)
End Sub

Public Sub IBaseScenario_UserDie(ByVal UserIndex)
    Dim RespawnInfo As clsRespawnInfo
    Set RespawnInfo = New clsRespawnInfo
    RespawnInfo.RespawnTime = GetTickCountRaw()
    With UserList(UserIndex)
        Dim i As Integer
        For i = LBound(.invent.Object) To UBound(.invent.Object)
            If .invent.Object(i).Equipped Then
                Call RespawnInfo.AddEquipedElement(i)
            End If
        Next i
        If .flags.TomoPocion Then
            RespawnInfo.HasPowerUp = True
            RespawnInfo.PowerUpTime = .flags.DuracionEfecto
            RespawnInfo.Strength = .Stats.UserAtributos(1)
            RespawnInfo.Agility = .Stats.UserAtributos(2)
        End If
    End With
    ' Plan 05.002 wave 2 (class B): store the slot IDENTITY, not just its number. The UserList
    ' slot is recycled between players: without this, if the dead player disconnects before
    ' reviving, the next one to inherit the slot was teleported to the arena, revived and had
    ' Strength/Agility overwritten. The field already existed and nobody used it.
    RespawnInfo.UserSlotIteration = UserList(UserIndex).VersionId
    ' Plan 05.002 wave 5: UserMod.DoDamageOrHeal does not check flags.Muerto and ModifyHealth
    ' returns True every time MinHp <= 0, also on a corpse: a second hit (a poison tick, a
    ' delayed blast) re-enters CustomScenarios.UserDie and the Add with a repeated key raised
    ' 457. The re-death is ignored: the good entry is the FIRST one, which captured the
    ' equipment before it was unequipped.
    If Not PlayerRespawn.Exists(UserIndex) Then
        Call PlayerRespawn.Add(UserIndex, RespawnInfo)
    End If
End Sub

Private Sub Respawn(ByVal player As Integer)
    Dim SpawnPos As t_WorldPos
    SpawnPos.Map = MapNumber
    SpawnPos.x = RandomNumber(12, 85)
    SpawnPos.y = RandomNumber(12, 85)
    Call WarpToLegalPos(player, MapNumber, SpawnPos.x, SpawnPos.y, True, True)
    Call RevivirUsuario(player)
    UserList(player).Counters.timeFx = 3
    Call SendData(SendTarget.ToPCAliveArea, player, PrepareMessageParticleFX(UserList(player).Char.charindex, e_ParticleEffects.Corazones, 100, False, , UserList(player).pos.x, _
            UserList(player).pos.y))
    Call SendData(SendTarget.ToPCAliveArea, player, PrepareMessagePlayWave("104", UserList(player).pos.x, UserList(player).pos.y))
End Sub

Private Sub RespawnPlayers()
    Dim key         As Variant
    Dim Keys        As Variant
    Dim currentTime As Long
    Dim RespawnInfo As clsRespawnInfo
    Dim i           As Integer
    currentTime = GetTickCountRaw()
    Keys = PlayerRespawn.Keys
    For Each key In Keys
        Set RespawnInfo = PlayerRespawn.Item(key)
        ' Plan 05.002 wave 2 (class B): if the slot is no longer the same player, the entry is
        ' orphaned (died and disconnected). Drop it without touching whoever inherits the slot.
        If UserList(key).VersionId <> RespawnInfo.UserSlotIteration Then
            Call PlayerRespawn.Remove(key)
        ElseIf TicksElapsed(RespawnInfo.RespawnTime, currentTime) > PlayerRespawnTime Then
            Call Respawn(key)
            For i = 0 To RespawnInfo.EquipedElementCount - 1
                Call EquiparInvItem(key, RespawnInfo.GetActiveSlot(i))
            Next i
            If RespawnInfo.HasPowerUp Then
                With UserList(key)
                    .Stats.UserAtributos(1) = RespawnInfo.Strength
                    .Stats.UserAtributos(2) = RespawnInfo.Agility
                    .flags.DuracionEfecto = RespawnInfo.PowerUpTime
                    Call WriteFYA(key)
                End With
            End If
            Call PlayerRespawn.Remove(key)
        End If
    Next key
End Sub

Sub GiveRewards()
    On Error GoTo GiveRewards_Err:
    Dim Teams()         As String
    Dim Rewards(1 To 3) As Long
    Dim Reward          As Long
    Dim Ranking()       As e_Rank
    Dim RewardStr       As String
    Dim i               As Integer
    Dim j               As Integer
    If LobbyList(LobbyIndex).TeamSize = 0 Then
        LobbyList(LobbyIndex).TeamSize = 1
    End If
    If (LobbyList(LobbyIndex).RegisteredPlayers / LobbyList(LobbyIndex).TeamSize) = 2 Then
        Rewards(1) = LobbyList(LobbyIndex).AvailableInscriptionMoney * 0.55
        Rewards(2) = LobbyList(LobbyIndex).AvailableInscriptionMoney * 0.35
    Else
        Rewards(1) = LobbyList(LobbyIndex).AvailableInscriptionMoney * 0.4
        Rewards(2) = LobbyList(LobbyIndex).AvailableInscriptionMoney * 0.3
        Rewards(3) = LobbyList(LobbyIndex).AvailableInscriptionMoney * 0.2
    End If
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MsgMatchComplete, Me.IBaseScenario_GetScenarioName(), e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
    If LobbyList(LobbyIndex).TeamSize > 1 Then
        ReDim Teams(0 To LobbyList(LobbyIndex).RegisteredPlayers / LobbyList(LobbyIndex).TeamSize) As String
        Ranking = Board.GetTeamRanking(LobbyList(LobbyIndex))
        For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
            If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
                Teams(LobbyList(LobbyIndex).Players(i).team) = Teams(LobbyList(LobbyIndex).Players(i).team) & UserList(LobbyList(LobbyIndex).Players(i).User.ArrayIndex).name & " "
            End If
        Next i
        ' Plan 05.002 wave 5: GetTeamRanking returns rank(0 To TeamCount), so UBound is the team
        ' count, not 3. With 2 teams Ranking(3) is error 9, swallowed by GiveRewards_Err.
        ' Rewards() is declared 1 To 3, so the cap stays at 3.
        For i = 1 To IIf(UBound(Ranking) < 3, UBound(Ranking), 3)
            If LobbyList(LobbyIndex).AvailableInscriptionMoney > 0 Then
                Reward = Rewards(i) / LobbyList(LobbyIndex).TeamSize
                For j = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
                    If IsValidUserRef(LobbyList(LobbyIndex).Players(j).User) Then
                        If LobbyList(LobbyIndex).Players(j).team = Ranking(i).PlayerIndex Then
                            Call GiveGoldToPlayer(LobbyList(LobbyIndex), j, Reward)
                        End If
                    End If
                Next j
                RewardStr = ", ganaron " & PonerPuntos(Reward) & " Modenas de oro cada uno."
            End If
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_PUNTOSA, i & "¬" & Teams(Ranking(i).PlayerIndex) & "¬" & Ranking(i).Score & "¬" & RewardStr, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) ' Msg1823=¬1) ¬2: ¬3 puntos¬4
        Next i
    Else
        Ranking = Board.GetRanking(3)
        For i = 0 To UBound(Ranking) - 1
            ' Plan 05.002 wave 5: GetRanking leaves 0 in the places that were not filled and
            ' UserList starts at 1, so the broadcast read UserList(0) and raised error 9,
            ' aborting the rest of the awards. The ranking comes ordered: the first hole ends it.
            If Ranking(i).PlayerIndex <= 0 Then Exit For
            If LobbyList(LobbyIndex).AvailableInscriptionMoney > 0 Then
                For j = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
                    If IsValidUserRef(LobbyList(LobbyIndex).Players(j).User) Then
                        If LobbyList(LobbyIndex).Players(j).User.ArrayIndex = Ranking(i).PlayerIndex Then
                            Call GiveGoldToPlayer(LobbyList(LobbyIndex), j, Rewards(i + 1))
                            Exit For
                        End If
                    End If
                Next j
                RewardStr = ", gano " & PonerPuntos(Rewards(i + 1)) & " Modenas de oro."
            End If
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_PUNTOS_1824, (i + 1) & "¬" & UserList(Ranking(i).PlayerIndex).name & "¬" & Ranking(i).Score & "¬" & RewardStr, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) ' Msg1824=¬1) ¬2: ¬3 puntos¬4
        Next i
    End If
    Exit Sub
GiveRewards_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.GiveRewards ", Erl)
End Sub
