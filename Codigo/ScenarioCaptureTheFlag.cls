VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScenarioCaptureTheFlag"
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
' ============================================================================
' CAPTURE THE FLAG as an engine scenario (plan 27.001, port of 05.002 wave 4).
'
' Replaces clsCaptura.cls, which ran outside the lobby system with its own singleton
' (InstanciaCaptura), its own participant collection, its own charge/refund and its own state
' machine. Participants, charge, refund, state and disconnections are now handled by the lobby;
' what stays here are flags, rounds and bases.
'
' TIME. Update() is called by modGameEvents.MaybeRunGameEvents every GAME_EVENTS_INTERVAL_MS =
' 40 ms, i.e. 25 times per second, while clsCaptura ran once per second. Everything that is
' "per second" lives inside the If UpdateTime(SecondTimer, frametime) block (same pattern as
' ScenarioDeathMatch). Moving logic out of there speeds the mode up 25x.
'
' CONTRACT NOTES (the compiler does not see any of these; the 05.002 review found them the hard way):
'  - The scenario is the one that advances its own lobby (UpdateWaitingForPlayers / Closed branch
'    in Update). Nothing else ticks a lobby that has a scenario.
'  - The CTF builds its own two teams (SortearDosEquipos). The lobby sort only runs with
'    TeamType = eRandom and TeamSize > 1, which neither creation path sends.
'  - CustomScenarios.UserDie runs BEFORE UserMod.UserDie at every call site, so inside the hook the
'    player is not dead yet. Deaths are queued and resolved on the next Update.
'  - UpdateLobbyState(Closed) only notifies. The real close (return players, release maps and the
'    pool slot) is CloseScenario, run from Update.
'  - Cancel / forcereset reach the scenario with the lobby data already wiped, so flag cleanup
'    sweeps by flags.jugando_captura instead of walking the lobby player list.
' ============================================================================
Option Explicit
Implements IBaseScenario

Private Enum e_CtfRoundState
    EsperandoInicio = 0
    IniciaRonda = 1
    EnJuego = 2
    Terminado = 3
End Enum

Private Const WORLD_MAP_COUNT As Integer = 3
Private LobbyIndex              As Integer
Private UpdateSlot              As Integer
Private LastFrameTime           As Long
Private SecondTimer             As t_Timer
Private RoundState              As e_CtfRoundState
Private cantidad_rondas         As Byte
Private ronda_actual            As Byte
Private tiempo_ronda            As Integer
Private contadorBandera(1 To 2) As Integer
Private rondasGanadas(1 To 2)   As Integer
Private colorEquipo(1 To 2)     As String
Private mapaEquipo(1 To 2)      As Integer
Private xEquipo(1 To 2)         As Byte
Private yEquipo(1 To 2)         As Byte
Private objEquipo(1 To 2)       As Integer
' The three CTF maps are fixed world maps (two bases + waiting room). Registered one by one so a
' failed registration (map already owned by another scenario) never makes this scenario release
' a map it does not own.
Private WorldMaps(1 To WORLD_MAP_COUNT)       As Integer
Private MapRegistered(1 To WORLD_MAP_COUNT)   As Boolean
Private m_Closed                              As Boolean
' Deaths cannot be resolved inside the hook (see the contract notes): the dead player is queued
' and processed in the next Update, once the death is real. Same idea as ScenarioDeathMatch's
' clsRespawnInfo list.
Private Const MAX_MUERTES_PENDIENTES         As Integer = 64
Private MuertePendiente(1 To MAX_MUERTES_PENDIENTES) As t_UserReference
Private MuerteIntentos(1 To MAX_MUERTES_PENDIENTES)  As Integer
Private MuertesPendientes                    As Integer
' Watchdog for an empty event. CTF is the only mode without MaxTime: a round only ends if someone
' holds the flag, so it needs its own termination rule when nobody is left.
Private Const CAPTURA_SEGUNDOS_SIN_JUGADORES As Integer = 120
Private SegundosSinJugadores                 As Integer
' Watchdog for a start that failed: the lobby is InProgress but ArrancarEvento swallowed an error
' and RoundState never left EsperandoInicio.
Private Const CAPTURA_SEGUNDOS_ARRANQUE_MAX  As Integer = 60
Private SegundosEsperandoInicio              As Integer

Private Sub Class_Initialize()
    colorEquipo(1) = "azul"
    colorEquipo(2) = "rojo"
    mapaEquipo(1) = MAP_TEAM_1
    mapaEquipo(2) = MAP_TEAM_2
    xEquipo(1) = X_BANDERA_1
    xEquipo(2) = X_BANDERA_2
    yEquipo(1) = Y_BANDERA_1
    yEquipo(2) = Y_BANDERA_2
    objEquipo(1) = OBJ_CAPTURA_BANDERA_1
    objEquipo(2) = OBJ_CAPTURA_BANDERA_2
    WorldMaps(1) = MAP_TEAM_1
    WorldMaps(2) = MAP_TEAM_2
    WorldMaps(3) = MAP_SALA_ESPERA
    UpdateSlot = 0
    m_Closed = False
    RoundState = EsperandoInicio
End Sub

' ============================ IBaseScenario ============================

Private Sub IBaseScenario_SetLobbyIndex(ByVal value As Integer)
    On Error GoTo ErrHandler
    Dim i       As Integer
    Dim allOk   As Boolean
    LobbyIndex = value
    cantidad_rondas = LobbyList(LobbyIndex).RoundCount
    If cantidad_rondas < 1 Then cantidad_rondas = 1
    ' The three maps are the route by which deaths and flags reach the scenario
    ' (CustomScenarios.GetMap(pos.Map)).
    allOk = True
    For i = 1 To WORLD_MAP_COUNT
        MapRegistered(i) = CustomScenarios.SetCustomScenario(WorldMaps(i), Me)
        If Not MapRegistered(i) Then allOk = False
    Next i
    If Not allOk Then
        ' SetCustomScenario already logged which map is taken. Give back the ones that were free
        ' and stay out of the tick: PrepareNewEvent sees the missing postcondition and aborts.
        Call LogError("ScenarioCaptureTheFlag.SetLobbyIndex: a CTF map is already registered by another scenario, the event is not started.")
        Call DesregistrarMapas
        Exit Sub
    End If
    UpdateSlot = CustomScenarios.AddUpdateScenario(Me)
    LastFrameTime = GetTickCountRaw()
    Call SetTimer(SecondTimer, 1000)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.SetLobbyIndex", Erl)
End Sub

Private Function IBaseScenario_GetScenarioName() As String
    IBaseScenario_GetScenarioName = "Captura de la Bandera"
End Function

Private Sub IBaseScenario_BroadcastOpenScenario()
    On Error GoTo ErrHandler
    'Msg1688: same text clsCaptura.inicializar announced.
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_INICIADO_CAPTURA_BANDERA_PARTICIPANTES_PRECIO_INSCRIPCION, LobbyList(LobbyIndex).MaxPlayers & "¬" & PonerPuntos(LobbyList(LobbyIndex).InscriptionPrice) & "¬" & LobbyList(LobbyIndex).MinLevel & "¬" & LobbyList(LobbyIndex).MaxLevel, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.BroadcastOpenScenario", Erl)
End Sub

Private Sub IBaseScenario_SendRules(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Call WriteConsoleMsg(UserIndex, "Captura de la Bandera: llevate la bandera enemiga a tu base y aguantala " & CAPTURA_TIEMPO_BANDERA & " segundos para ganar la ronda.", e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.SendRules", Erl)
End Sub

Private Function IBaseScenario_GetExtraInterfaces() As Long
    IBaseScenario_GetExtraInterfaces = 0
End Function

Private Function IBaseScenario_RequiresSpawn() As Boolean
    ' Fixed maps: the GM does not need to set a spawn point.
    IBaseScenario_RequiresSpawn = False
End Function

Private Function IBaseScenario_ValidateUser(ByVal UserIndex As Integer) As Integer
    ' Engine contract: return the MESSAGE ID of the rejection (>0) or 0 if the player can join.
    ' The message is sent by the lobby (ModLobby.CanPlayerJoin), not here. The lobby already
    ' validates gold, jail trigger, level and capacity; only CTF-specific rules go here, with the
    ' same messages clsCaptura used.
    On Error GoTo ErrHandler
    IBaseScenario_ValidateUser = 0
    With UserList(UserIndex)
        If MapInfo(.pos.Map).Seguro = 0 Then
            'Msg928
            IBaseScenario_ValidateUser = MSG_SOLAMENTE_PODRAS_INGRESAR_EVENTO_ZONA_SEGURA
        ElseIf .Counters.Pena > 0 Then
            'Msg930
            IBaseScenario_ValidateUser = MSG_NO_PUEDES_ANOTARTE_EVENTO_MIENTRAS_ENCARCELADO
        ElseIf .flags.Muerto = 1 Then
            'Msg931
            IBaseScenario_ValidateUser = MSG_NO_PUEDES_ANOTARTE_EVENTO_ESTANDO_MUERTO
        ElseIf .flags.Comerciando Then
            'Msg932
            IBaseScenario_ValidateUser = MSG_NO_PUEDES_INGRESAR_EVENTO_MIENTRAS_COMERCIANDO
        ElseIf .flags.Montado = 1 Then
            'Msg933
            IBaseScenario_ValidateUser = MSG_NO_PUEDES_INGRESAR_EVENTO_ESTANDO_MONTADO
        ElseIf .flags.Navegando = 1 Then
            'Msg934
            IBaseScenario_ValidateUser = MSG_NO_PUEDES_INGRESAR_EVENTO_MIENTRAS_NAVEGANDO
        End If
    End With
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.ValidateUser", Erl)
    IBaseScenario_ValidateUser = 0
End Function

Private Sub IBaseScenario_WillSummonPlayer(ByVal UserIndex As Integer)
    ' Summoned players go to the waiting room, as before.
    On Error GoTo ErrHandler
    Dim SpawnPos As t_WorldPos
    SpawnPos.Map = MAP_SALA_ESPERA
    SpawnPos.x = RandomNumber(MIN_SALA_ESPERA_X, MAX_SALA_ESPERA_X)
    SpawnPos.y = RandomNumber(MIN_SALA_ESPERA_Y, MAX_SALA_ESPERA_Y)
    LobbyList(LobbyIndex).SummonCoordinates = SpawnPos
    UserList(UserIndex).flags.jugando_captura = 1
    UserList(UserIndex).flags.tiene_bandera = 0
    ' jugando_captura_timer / jugando_captura_muertes are NOT touched here; the per-round reset
    ' lives in IniciarRondaNueva. Note this also runs on reconnect, but TCP.ResetUserFlags already
    ' zeroes the timer at login, so a relog never kept the death penalty; OnUserReconnect is the
    ' one that re-arms the timer so a player coming back mid-round is sent to his base.
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.WillSummonPlayer", Erl)
End Sub

Private Sub IBaseScenario_UpdateLobbyState(ByVal oldState As e_LobbyState, ByRef NewState As e_LobbyState)
    On Error GoTo ErrHandler
    If oldState = NewState Then Exit Sub
    If NewState = e_LobbyState.InProgress Then
        Call ArrancarEvento
    ElseIf NewState = e_LobbyState.Closed Then
        Call LimpiarEscenario
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.UpdateLobbyState", Erl)
End Sub

Private Sub IBaseScenario_Update()
    On Error GoTo ErrHandler
    Dim nowRaw    As Long
    Dim frametime As Long
    nowRaw = GetTickCountRaw()
    frametime = CLng(TicksElapsed(LastFrameTime, nowRaw))
    LastFrameTime = nowRaw
    ' Death queue at engine pace (40 ms) so the warp to the waiting room feels immediate.
    If MuertesPendientes > 0 Then Call ProcesarMuertesPendientes
    ' The scenario advances its own lobby. Without this branch the CTF stays in AcceptingPlayers
    ' forever (it neither starts when full nor when the time is up) and never closes when it ends.
    If LobbyList(LobbyIndex).State = e_LobbyState.AcceptingPlayers Then
        Call UpdateWaitingForPlayers(frametime, LobbyList(LobbyIndex))
        Exit Sub
    ElseIf LobbyList(LobbyIndex).State = e_LobbyState.Closed Then
        Call CloseScenario
        Exit Sub
    End If
    ' EVERYTHING below runs once per second. Update() is entered 25 times per second.
    If Not UpdateTime(SecondTimer, frametime) Then Exit Sub
    ' Termination watchdog. If the last player drops, TickEnJuego iterates invalid references and
    ' nothing ever happens: the lobby stays InProgress forever, the slot never goes back to the
    ' pool and ActiveEventType() keeps reporting a public event, so no other one can be opened
    ' until a restart. This guard only covers the empty-lobby case; a start that failed with
    ' players still connected (ArrancarEvento swallowed an error, RoundState stays EsperandoInicio
    ' with the lobby InProgress) is covered by the separate check right below.
    If RoundState = EsperandoInicio And LobbyList(LobbyIndex).State = e_LobbyState.InProgress Then
        SegundosEsperandoInicio = SegundosEsperandoInicio + 1
        If SegundosEsperandoInicio >= CAPTURA_SEGUNDOS_ARRANQUE_MAX Then
            Call LogError("[Evento] CTF cancelled: the round did not start " & CAPTURA_SEGUNDOS_ARRANQUE_MAX & "s after the lobby went InProgress.")
            ' Same path as the watchdog below: refund + Closed, CloseScenario releases on next tick.
            Call ModLobby.CancelLobby(LobbyList(LobbyIndex))
            Exit Sub
        End If
    End If
    If RoundState <> Terminado And LobbyList(LobbyIndex).State = e_LobbyState.InProgress Then
        If JugadoresConectados() = 0 Then
            SegundosSinJugadores = SegundosSinJugadores + 1
            If SegundosSinJugadores >= CAPTURA_SEGUNDOS_SIN_JUGADORES Then
                Call LogError("[Evento] CTF closed by watchdog: " & CAPTURA_SEGUNDOS_SIN_JUGADORES & "s without connected players.")
                ' CancelLobby refunds the inscription (offline players too, via the DB) and leaves
                ' the lobby Closed. The real close (maps and slot) is CloseScenario on the next
                ' tick, through the Closed branch above.
                Call ModLobby.CancelLobby(LobbyList(LobbyIndex))
                Exit Sub
            End If
        Else
            SegundosSinJugadores = 0
        End If
    End If
    If RoundState = IniciaRonda Then
        Call TickCuentaRegresivaRonda
    ElseIf RoundState = EnJuego Then
        Call TickEnJuego
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.Update", Erl)
End Sub

Private Sub IBaseScenario_UserDie(ByVal UserIndex)
    ' The player is NOT dead yet here (see the contract notes). What does not depend on Muerto is
    ' done now; the warp and the resurrection are queued. The hook also fires for ANY player dying
    ' on a registered CTF map, including the waiting room, so only actual participants count.
    On Error GoTo ErrHandler
    If RoundState <> EnJuego Then Exit Sub
    With UserList(UserIndex)
        If .flags.jugando_captura <> 1 Then Exit Sub
        If .flags.CurrentTeam < 1 Or .flags.CurrentTeam > 2 Then Exit Sub
        'Msg1715
        Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_EQUIPO_MUERTO_BATALLA, .name & "¬" & colorEquipo(.flags.CurrentTeam), e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
        Call TirarBanderaDe(UserIndex)
        .flags.jugando_captura_muertes = .flags.jugando_captura_muertes + 1
        .flags.jugando_captura_timer = .flags.jugando_captura_muertes * CAPTURA_TIEMPO_POR_MUERTE + CAPTURA_TIEMPO_MUERTE
        'Msg1846
        Call WriteLocaleMsg(UserIndex, MSG_VOLVERAS_BASE_SEGUNDOS, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, CStr(.flags.jugando_captura_timer))
    End With
    Call EncolarMuerte(UserIndex)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.UserDie", Erl)
End Sub

Private Function JugadoresConectados() As Integer
    ' Counts registered players with a live session. RegisteredPlayers is NOT enough: a
    ' disconnection does not remove the player from the lobby (RegisterDisconnectedUserOnLobby only
    ' marks Connected = False and returns him to his position), so it keeps counting people who are
    ' gone.
    On Error GoTo ErrHandler
    Dim i As Integer
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        If LobbyList(LobbyIndex).Players(i).Connected Then
            If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
                JugadoresConectados = JugadoresConectados + 1
            End If
        End If
    Next i
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.JugadoresConectados", Erl)
End Function

Private Sub EncolarMuerte(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Dim i As Integer
    For i = 1 To MuertesPendientes
        If MuertePendiente(i).ArrayIndex = UserIndex Then Exit Sub
    Next i
    If MuertesPendientes >= MAX_MUERTES_PENDIENTES Then
        ' Should not happen (CTF capacity is 48 and the queue drains every 40 ms), but if it does
        ' resolve it right away rather than lose the player.
        Call DevolverALaSalaMuerto(UserIndex)
        Exit Sub
    End If
    MuertesPendientes = MuertesPendientes + 1
    Call SetUserRef(MuertePendiente(MuertesPendientes), UserIndex)
    MuerteIntentos(MuertesPendientes) = 0
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.EncolarMuerte", Erl)
End Sub

Private Sub ProcesarMuertesPendientes()
    On Error GoTo ErrHandler
    Dim i  As Integer
    Dim ui As Integer
    ' Backwards: the queue is compacted when an element is taken out.
    For i = MuertesPendientes To 1 Step -1
        MuerteIntentos(i) = MuerteIntentos(i) + 1
        If Not IsValidUserRef(MuertePendiente(i)) Then
            Call SacarDeLaCola(i)
        Else
            ui = MuertePendiente(i).ArrayIndex
            If UserList(ui).flags.Muerto = 1 Then
                Call DevolverALaSalaMuerto(ui)
                Call SacarDeLaCola(i)
            ElseIf MuerteIntentos(i) >= 25 Then
                ' One second waiting for a Muerto = 1 that never came: something else revived
                ' him (a GM, a priest). Let him go untouched.
                Call SacarDeLaCola(i)
            End If
        End If
    Next i
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.ProcesarMuertesPendientes", Erl)
End Sub

Private Sub SacarDeLaCola(ByVal Index As Integer)
    If Index < MuertesPendientes Then
        MuertePendiente(Index) = MuertePendiente(MuertesPendientes)
        MuerteIntentos(Index) = MuerteIntentos(MuertesPendientes)
    End If
    Call ClearUserRef(MuertePendiente(MuertesPendientes))
    MuerteIntentos(MuertesPendientes) = 0
    MuertesPendientes = MuertesPendientes - 1
End Sub

Private Sub DevolverALaSalaMuerto(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Call EnviarUsuarioSala(UserIndex)
    Call ResucitarOCurar(UserIndex)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.DevolverALaSalaMuerto", Erl)
End Sub

Private Sub IBaseScenario_OnUserDisconnected(ByVal UserIndex As Integer)
    ' The lobby already marks the slot as disconnected and returns the player; here the flag is
    ' only dropped so it does not get lost in limbo.
    On Error GoTo ErrHandler
    Call TirarBanderaDe(UserIndex)
    UserList(UserIndex).flags.jugando_captura = 0
    UserList(UserIndex).flags.tiene_bandera = 0
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.OnUserDisconnected", Erl)
End Sub

Private Sub IBaseScenario_OnUserReconnect(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    If RoundState = EnJuego Or RoundState = IniciaRonda Then
        UserList(UserIndex).flags.jugando_captura = 1
        Call EnviarUsuarioSala(UserIndex)
        ' Mid-round the waiting room is a dead end: TickEnJuego only sends a player to his base
        ' when jugando_captura_timer counts down to 0, and ResetUserFlags zeroed it at login.
        ' Arm it like a death respawn (CAPTURA_TIEMPO_MUERTE) so he is returned to his base.
        If RoundState = EnJuego Then
            UserList(UserIndex).flags.jugando_captura_timer = CAPTURA_TIEMPO_MUERTE
        End If
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.OnUserReconnect", Erl)
End Sub

Private Sub IBaseScenario_Reset()
    ' Reset contract: release the maps, leave the tick and leave the player flags clean. The
    ' players themselves are returned by the lobby (RemovePlayerFromEvent), and ForceReset empties
    ' RegisteredPlayers BEFORE calling here, so LimpiarEscenario cannot rely on the lobby list.
    ' Idempotent: DesregistrarMapas only releases what was registered and UpdateSlot = 0 means
    ' "already out of the tick".
    On Error GoTo ErrHandler
    Call LimpiarEscenario
    Call DesregistrarMapas
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.Reset", Erl)
End Sub

Private Sub IBaseScenario_PlayerKillNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
End Sub

Private Sub IBaseScenario_NPcKillPlayer(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
End Sub

Private Sub IBaseScenario_PlayerKillPlayer(ByVal killerIndex As Integer, ByVal deadIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
End Sub

Private Sub IBaseScenario_NpcDie(ByVal NpcIndex)
End Sub

Private Sub IBaseScenario_GetNextWaypointForNpc(ByVal NpcIndex As Integer, ByRef PosX As Integer, ByRef PosY As Integer)
End Sub

' ======================= CTF own logic =======================

Private Sub ArrancarEvento()
    ' The lobby already charged the inscription and summoned everyone to the waiting room.
    ' The teams are built by the CTF, NOT by SortTeams (see the contract notes).
    On Error GoTo ErrHandler
    Dim i    As Integer
    Dim ui   As Integer
    Dim team As Integer
    'Msg1714
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_EVENTO_CAPTURA_BANDERA_INICIADO, vbNullString, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
    Call SortearDosEquipos
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
            ui = LobbyList(LobbyIndex).Players(i).User.ArrayIndex
            team = LobbyList(LobbyIndex).Players(i).team
            If team < 1 Then team = 1
            If team > 2 Then team = 2
            With UserList(ui)
                .flags.CurrentTeam = team
                .flags.tiene_bandera = 0
                Call SendData(SendTarget.ToAll, ui, PrepareMessageUpdateFlag(0, .Char.charindex))
                'Msg1845
                Call WriteLocaleMsg(ui, MSG_TOCADO_PARTICIPAR_EQUIPO_DEBERAS_PROTEGER_COMPANEROS_EQUIPO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, colorEquipo(team))
            End With
        End If
    Next i
    ronda_actual = 0
    rondasGanadas(1) = 0
    rondasGanadas(2) = 0
    Call IniciarRondaNueva
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.ArrancarEvento", Erl)
End Sub

Private Sub IniciarRondaNueva()
    On Error GoTo ErrHandler
    Dim i  As Integer
    Dim ui As Integer
    RoundState = IniciaRonda
    ronda_actual = ronda_actual + 1
    tiempo_ronda = CAPTURA_TIEMPO_INICIO_RONDA
    contadorBandera(1) = 0
    contadorBandera(2) = 0
    'Msg1720
    Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_RONDA_DARA_COMIENZO_SEGUNDOS, ronda_actual & "¬" & tiempo_ronda, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
            ui = LobbyList(LobbyIndex).Players(i).User.ArrayIndex
            Call SendData(SendTarget.ToAll, ui, PrepareMessageUpdateFlag(0, UserList(ui).Char.charindex))
            UserList(ui).flags.jugando_captura_muertes = 0
            UserList(ui).flags.jugando_captura_timer = 0
            Call TirarBanderaDe(ui)
            Call EnviarUsuarioSala(ui)
        End If
    Next i
    ' Players no longer hold flags: remove any flag left on the floor before restoring the two.
    Call EliminarBanderasDelMapa
    Call RestaurarBandera(1)
    Call RestaurarBandera(2)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.IniciarRondaNueva", Erl)
End Sub

Private Sub TickCuentaRegresivaRonda()
    On Error GoTo ErrHandler
    Dim i As Integer
    tiempo_ronda = tiempo_ronda - 1
    If tiempo_ronda Mod 30 = 0 And tiempo_ronda > 5 Then
        'Msg1710
        Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_SEGUNDOS_COMIENZA_RONDA, tiempo_ronda & "¬" & ronda_actual, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
    ElseIf tiempo_ronda <= 5 And tiempo_ronda > 0 Then
        'Msg1711
        Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_RONDA_COMIENZA, CStr(tiempo_ronda), e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
    ElseIf tiempo_ronda <= 0 Then
        'Msg1712
        Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_YAAA, vbNullString, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
        For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
            If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
                Call EnviarUsuarioBase(LobbyList(LobbyIndex).Players(i).User.ArrayIndex)
            End If
        Next i
        RoundState = EnJuego
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.TickCuentaRegresivaRonda", Erl)
End Sub

Private Sub TickEnJuego()
    On Error GoTo ErrHandler
    Dim i                     As Integer
    Dim ui                    As Integer
    Dim team                  As Byte
    Dim estaEnBandera(1 To 2) As Boolean
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
            ui = LobbyList(LobbyIndex).Players(i).User.ArrayIndex
            With UserList(ui)
                team = .flags.CurrentTeam
                If team >= 1 And team <= 2 Then
                    If .pos.Map = mapaEquipo(team) Then
                        If Abs(.pos.x - xEquipo(team)) <= 8 And Abs(.pos.y - yEquipo(team)) <= 5 Then
                            If TieneObjetos(objEquipo(IIf(team = 1, 2, 1)), 1, ui) Then
                                contadorBandera(team) = contadorBandera(team) + 1
                                estaEnBandera(team) = True
                                If contadorBandera(team) >= CAPTURA_TIEMPO_BANDERA Then
                                    Call FinalizarRonda(team)
                                    Exit Sub
                                Else
                                    'Msg1713
                                    Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_ATENCION_EQUIPO_TIENE_BANDERA_ENEMIGA_BASE, .name & "¬" & colorEquipo(team) & "¬" & (CAPTURA_TIEMPO_BANDERA - contadorBandera(team)), e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
                                End If
                            End If
                        End If
                    ElseIf .flags.jugando_captura_timer > 0 Then
                        .flags.jugando_captura_timer = .flags.jugando_captura_timer - 1
                        If .flags.jugando_captura_timer = 0 Then
                            Call EnviarUsuarioBase(ui)
                        End If
                    End If
                End If
            End With
        End If
    Next i
    If Not estaEnBandera(1) Then contadorBandera(1) = 0
    If Not estaEnBandera(2) Then contadorBandera(2) = 0
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.TickEnJuego", Erl)
End Sub

Private Sub FinalizarRonda(ByVal team As Byte)
    On Error GoTo ErrHandler
    Dim rondasParaGanar As Integer
    'Msg1725
    Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_EQUIPO_GANADO_RONDA, colorEquipo(team) & "¬" & ronda_actual, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
    Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))
    contadorBandera(1) = 0
    contadorBandera(2) = 0
    rondasGanadas(team) = rondasGanadas(team) + 1
    rondasParaGanar = (cantidad_rondas \ 2) + 1
    If rondasGanadas(team) >= rondasParaGanar Then
        Call TerminarEvento(CInt(team))
    Else
        Call IniciarRondaNueva
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.FinalizarRonda", Erl)
End Sub

Private Sub TerminarEvento(ByVal team As Integer)
    ' Closing pattern of the plan: 1) mark finished, 2) pay over the lobby list without mutating
    ' it, 3) announce, 4) let the lobby release.
    On Error GoTo ErrHandler
    Dim i          As Integer
    Dim ui         As Integer
    Dim recompensa As Long
    If RoundState = Terminado Then Exit Sub
    RoundState = Terminado
    If team >= 1 And team <= 2 Then
        recompensa = LobbyList(LobbyIndex).InscriptionPrice * 2
        For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
            If LobbyList(LobbyIndex).Players(i).team = team Then
                ' GiveGoldToPlayer also pays a disconnected winner (credited in the database).
                If GiveGoldToPlayer(LobbyList(LobbyIndex), i, recompensa) Then
                    If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
                        ui = LobbyList(LobbyIndex).Players(i).User.ArrayIndex
                        'Msg1847
                        Call WriteLocaleMsg(ui, MSG_HAS_RECUPERADO_INSCRIPCION_GANADO_MONEDAS_ORO_ADICIONALES, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, PonerPuntos(LobbyList(LobbyIndex).InscriptionPrice))
                    End If
                End If
            End If
        Next i
        'Msg1687
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_EVENTO_CAPTURA_BANDERA_TERMINADO_GANADOR_EQUIPO, colorEquipo(team), e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
    Else
        ' Draw (not reachable today: FinalizarRonda always passes a team). Nobody wins, so the
        ' pool goes back to everybody instead of being lost.
        For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
            Call RefundPlayer(LobbyList(LobbyIndex), i)
        Next i
        'Msg1686 (draw)
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_EVENTO_CAPTURA_BANDERA_TERMINADO_EMPATE, vbNullString, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
    End If
    ' Moving to Closed does not close anything by itself: UpdateLobbyState only notifies and
    ' changes the state. The real close (return players, release maps, give the pool slot back) is
    ' CloseScenario on the next Update, like the other scenarios. Without it the players stayed
    ' trapped and one lobby leaked per finished CTF.
    Call ModLobby.UpdateLobbyState(LobbyList(LobbyIndex), e_LobbyState.Closed)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.TerminarEvento", Erl)
End Sub

Private Sub LimpiarEscenario()
    ' Cannot walk the lobby list nor look at CurrentTeam to know who holds a flag: both close
    ' paths arrive here with that data already erased (CancelLobby runs ReturnAllPlayers, which
    ' sets CurrentTeam = 0, before notifying the scenario, and ForceReset empties
    ' RegisteredPlayers before calling Reset). The player would keep the enemy banner and win the
    ' next CTF round just by standing on his base.
    ' So: sweep by flags.jugando_captura (the criterion SendTarget.ToJugadoresCaptura uses, which
    ' nobody else touches) and remove BOTH flags without asking which team they belonged to.
    On Error GoTo ErrHandler
    Dim ui As Integer
    RoundState = Terminado
    MuertesPendientes = 0
    For ui = 1 To LastUser
        If UserList(ui).flags.jugando_captura = 1 Then
            Call SacarBanderasDe(ui)
            UserList(ui).flags.jugando_captura = 0
            UserList(ui).flags.tiene_bandera = 0
            UserList(ui).flags.CurrentTeam = 0
            UserList(ui).flags.jugando_captura_timer = 0
            UserList(ui).flags.jugando_captura_muertes = 0
        End If
    Next ui
    Call EliminarBanderasDelMapa
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.LimpiarEscenario", Erl)
End Sub

Private Sub EliminarBanderasDelMapa()
    ' Removes every flag object (both teams) lying on the CTF maps this scenario owns: the
    ' restored ones at the bases and the ones dropped by dead/disconnected carriers. Without it a
    ' stray flag survives the round (duplicate flag) or the event (stray flag in the world).
    On Error GoTo ErrHandler
    Dim i As Integer
    Dim x As Integer
    Dim y As Integer
    For i = 1 To WORLD_MAP_COUNT
        If MapRegistered(i) Then
            For y = YMinMapSize To YMaxMapSize
                For x = XMinMapSize To XMaxMapSize
                    If MapData(x, y, WorldMaps(i)).ObjInfo.ObjIndex = OBJ_CAPTURA_BANDERA_1 Or MapData(x, y, WorldMaps(i)).ObjInfo.ObjIndex = OBJ_CAPTURA_BANDERA_2 Then
                        Call EraseObj(MapData(x, y, WorldMaps(i)).ObjInfo.amount, WorldMaps(i), x, y)
                    End If
                Next x
            Next y
        End If
    Next i
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.EliminarBanderasDelMapa", Erl)
End Sub

Private Sub SacarBanderasDe(ByVal UserIndex As Integer)
    ' Cleanup version: removes both flags and does NOT drop them on the floor (the event is over).
    ' TirarBanderaDe is the in-game one and does leave them where the player died.
    On Error GoTo ErrHandler
    Dim actualizar As Boolean
    If TieneObjetos(OBJ_CAPTURA_BANDERA_1, 1, UserIndex) Then
        Call QuitarObjetos(OBJ_CAPTURA_BANDERA_1, 1, UserIndex)
        actualizar = True
    End If
    If TieneObjetos(OBJ_CAPTURA_BANDERA_2, 1, UserIndex) Then
        Call QuitarObjetos(OBJ_CAPTURA_BANDERA_2, 1, UserIndex)
        actualizar = True
    End If
    If actualizar Then
        UserList(UserIndex).flags.tiene_bandera = 0
        Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(0, UserList(UserIndex).Char.charindex))
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.SacarBanderasDe", Erl)
End Sub

Private Sub DesregistrarMapas()
    On Error GoTo ErrHandler
    Dim i As Integer
    ' Only the maps this scenario registered itself are released.
    For i = 1 To WORLD_MAP_COUNT
        If MapRegistered(i) Then
            Call CustomScenarios.RemoveCustomScenario(WorldMaps(i))
            MapRegistered(i) = False
        End If
    Next i
    If UpdateSlot > 0 Then
        Call CustomScenarios.RemoveUpdateScenario(UpdateSlot)
        UpdateSlot = 0
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.DesregistrarMapas", Erl)
End Sub

Private Sub CloseScenario()
    ' Same contract as ScenarioDeathMatch.CloseScenario: return the players, drop the reference,
    ' take the lobby back to UnInitilized, unregister and release the slot. The order matters:
    ' first clean flags (while the player is still on the CTF map), then return him.
    On Error GoTo ErrHandler
    If m_Closed Then Exit Sub
    m_Closed = True
    Call LimpiarEscenario
    Call ModLobby.ReturnAllPlayers(LobbyList(LobbyIndex))
    Set LobbyList(LobbyIndex).Scenario = Nothing
    Call ModLobby.UpdateLobbyState(LobbyList(LobbyIndex), e_LobbyState.UnInitilized)
    Call DesregistrarMapas
    Call ReleaseLobby(LobbyIndex)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.CloseScenario", Erl)
End Sub

Private Sub SortearDosEquipos()
    ' Two teams balanced by level, serpentine 1-2-2-1-1-2... as clsCaptura.iniciarCaptura did.
    ' With an odd headcount the legacy CANCELLED the whole event (and lacked the Exit Sub); here
    ' the lowest-level player is removed and refunded, which is what SortTeams already does for
    ' the other modes and does not punish the rest.
    On Error GoTo ErrHandler
    Dim i         As Integer
    Dim ui        As Integer
    Dim best      As Integer
    Dim bestLvl   As Integer
    Dim team      As Integer
    Dim direction As Integer
    With LobbyList(LobbyIndex)
        If .RegisteredPlayers Mod 2 = 1 Then
            best = -1
            bestLvl = 32767
            For i = 0 To .RegisteredPlayers - 1
                If IsValidUserRef(.Players(i).User) Then
                    If UserList(.Players(i).User.ArrayIndex).Stats.ELV <= bestLvl Then
                        bestLvl = UserList(.Players(i).User.ArrayIndex).Stats.ELV
                        best = i
                    End If
                Else
                    ' A disconnected player is the first to go.
                    best = i
                    Exit For
                End If
            Next i
            If best >= 0 Then
                ' KickPlayer returns and refunds but does not tell the scenario: the player kept
                ' jugando_captura = 1 and went on receiving the event messages (and could not
                ' enter a challenge) until the close.
                If IsValidUserRef(.Players(best).User) Then
                    ui = .Players(best).User.ArrayIndex
                    Call SacarBanderasDe(ui)
                    UserList(ui).flags.jugando_captura = 0
                    UserList(ui).flags.CurrentTeam = 0
                    UserList(ui).flags.jugando_captura_timer = 0
                    UserList(ui).flags.jugando_captura_muertes = 0
                    Call WriteLocaleMsg(ui, MsgNotEnoughPlayerForTeam, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
                End If
                Call KickPlayer(LobbyList(LobbyIndex), best)
            End If
        End If
        For i = 0 To .RegisteredPlayers - 1
            .Players(i).team = 0
        Next i
        team = 1
        direction = 1
        Do
            best = -1
            bestLvl = -1
            For i = 0 To .RegisteredPlayers - 1
                If .Players(i).team = 0 And IsValidUserRef(.Players(i).User) Then
                    If UserList(.Players(i).User.ArrayIndex).Stats.ELV > bestLvl Then
                        bestLvl = UserList(.Players(i).User.ArrayIndex).Stats.ELV
                        best = i
                    End If
                End If
            Next i
            If best < 0 Then Exit Do
            .Players(best).team = team
            UserList(.Players(best).User.ArrayIndex).flags.CurrentTeam = team
            team = team + direction
            If team > 2 Then
                direction = -1
                team = 2
            ElseIf team < 1 Then
                direction = 1
                team = 1
            End If
        Loop
        ' Players who disconnected between the inscription and the start have no level to read;
        ' they are spread alternating so the headcount is not unbalanced.
        For i = 0 To .RegisteredPlayers - 1
            If .Players(i).team = 0 Then
                .Players(i).team = team
                team = team + direction
                If team > 2 Then
                    direction = -1
                    team = 2
                ElseIf team < 1 Then
                    direction = 1
                    team = 1
                End If
            End If
        Next i
        .TeamSortDone = True
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.SortearDosEquipos", Erl)
End Sub

Private Sub EnviarUsuarioBase(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Dim pos_map As Integer
    Dim pos_X   As Byte
    Dim pos_Y   As Byte
    If UserList(UserIndex).flags.CurrentTeam = 1 Then
        pos_map = MAP_TEAM_1
        pos_X = X_TEAM_1
        pos_Y = Y_TEAM_1
    ElseIf UserList(UserIndex).flags.CurrentTeam = 2 Then
        pos_map = MAP_TEAM_2
        pos_X = X_TEAM_2
        pos_Y = Y_TEAM_2
    Else
        Exit Sub
    End If
    Call FindLegalPos(UserIndex, pos_map, pos_X, pos_Y)
    Call WarpUserChar(UserIndex, pos_map, pos_X, pos_Y, True)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.EnviarUsuarioBase", Erl)
End Sub

Private Sub EnviarUsuarioSala(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Dim pos_X As Byte
    Dim pos_Y As Byte
    pos_X = RandomNumber(MIN_SALA_ESPERA_X, MAX_SALA_ESPERA_X)
    pos_Y = RandomNumber(MIN_SALA_ESPERA_Y, MAX_SALA_ESPERA_Y)
    Call FindLegalPos(UserIndex, MAP_SALA_ESPERA, pos_X, pos_Y)
    Call WarpUserChar(UserIndex, MAP_SALA_ESPERA, pos_X, pos_Y, True)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.EnviarUsuarioSala", Erl)
End Sub

Private Sub RestaurarBandera(ByVal team As Byte)
    On Error GoTo ErrHandler
    Dim MiObj As t_Obj
    Dim MiPos As t_WorldPos
    MiObj.amount = 1
    MiObj.ObjIndex = objEquipo(team)
    MiPos.Map = mapaEquipo(team)
    MiPos.x = xEquipo(team)
    MiPos.y = yEquipo(team)
    If MapData(MiPos.x, MiPos.y, MiPos.Map).ObjInfo.amount <= 0 Then
        Call TirarItemAlPiso(MiPos, MiObj, False)
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.RestaurarBandera", Erl)
End Sub

Private Sub TirarBanderaDe(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Dim team       As Byte
    Dim objBandera As Integer
    Dim MiObj      As t_Obj
    With UserList(UserIndex)
        team = .flags.CurrentTeam
        If team = 1 Then
            objBandera = OBJ_CAPTURA_BANDERA_2
        ElseIf team = 2 Then
            objBandera = OBJ_CAPTURA_BANDERA_1
        Else
            Exit Sub
        End If
        If TieneObjetos(objBandera, 1, UserIndex) Then
            .flags.tiene_bandera = 0
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(0, .Char.charindex))
            Call QuitarObjetos(objBandera, 1, UserIndex)
            If RoundState = EnJuego Then
                MiObj.amount = 1
                MiObj.ObjIndex = objBandera
                Call TirarItemAlPiso(.pos, MiObj, False)
                'Msg1719
                Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_ATENCION_BANDERA_EQUIPO_ENCUENTRA_1719, colorEquipo(team) & "¬" & .pos.Map & "," & .pos.x & "," & .pos.y, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.TirarBanderaDe", Erl)
End Sub

' =============== Public API for the flag call sites ===============

' Called by ModLobby.ReturnPlayer BEFORE it clears CurrentTeam (GM kick / returnplayer / returnall):
' TirarBanderaDe needs the team to know which flag the player may carry.
Public Sub OnPlayerRemoved(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Call TirarBanderaDe(UserIndex)
    UserList(UserIndex).flags.jugando_captura = 0
    UserList(UserIndex).flags.tiene_bandera = 0
    UserList(UserIndex).flags.jugando_captura_timer = 0
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.OnPlayerRemoved", Erl)
End Sub
' Called from InvUsuario through ModCaptura.GetCapturaEscenario (CustomScenarios.GetMap + TypeOf).

Public Function TomaBandera(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
    On Error GoTo ErrHandler
    Dim banderaTeam As Byte
    TomaBandera = True
    If ObjIndex = OBJ_CAPTURA_BANDERA_1 Then
        banderaTeam = 1
    ElseIf ObjIndex = OBJ_CAPTURA_BANDERA_2 Then
        banderaTeam = 2
    Else
        Exit Function
    End If
    With UserList(UserIndex)
        If banderaTeam = .flags.CurrentTeam Then
            If .pos.x = xEquipo(banderaTeam) And .pos.y = yEquipo(banderaTeam) Then
                TomaBandera = False
                Exit Function
            End If
        End If
        If .flags.CurrentTeam <> banderaTeam Then
            .flags.tiene_bandera = banderaTeam
        Else
            .flags.tiene_bandera = 0
        End If
        Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(.flags.tiene_bandera, .Char.charindex))
    End With
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.TomaBandera", Erl)
    TomaBandera = True
End Function

Public Sub QuitarBandera(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
    On Error GoTo ErrHandler
    Dim banderaTeam As Byte
    If ObjIndex = OBJ_CAPTURA_BANDERA_1 Then
        banderaTeam = 1
    ElseIf ObjIndex = OBJ_CAPTURA_BANDERA_2 Then
        banderaTeam = 2
    Else
        Exit Sub
    End If
    With UserList(UserIndex)
        If .flags.CurrentTeam < 1 Or .flags.CurrentTeam > 2 Then Exit Sub
        If banderaTeam <> .flags.CurrentTeam Then
            'Msg1716
            Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_ATENCION_EQUIPO_TOMADO_BANDERA_EQUIPO, .name & "¬" & colorEquipo(.flags.CurrentTeam) & "¬" & colorEquipo(banderaTeam), e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
            .flags.tiene_bandera = banderaTeam
        Else
            .flags.tiene_bandera = 0
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(0, .Char.charindex))
            Call QuitarObjetos(ObjIndex, 1, UserIndex)
            Call RestaurarBandera(banderaTeam)
            'Msg1717
            Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_ATENCION_BANDERA_EQUIPO_SIDO_RECUPERADA, colorEquipo(banderaTeam), e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.QuitarBandera", Erl)
End Sub

Public Sub TiraBandera(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
    On Error GoTo ErrHandler
    Dim team As Byte
    If ObjIndex = OBJ_CAPTURA_BANDERA_1 Then
        team = 1
    ElseIf ObjIndex = OBJ_CAPTURA_BANDERA_2 Then
        team = 2
    Else
        Exit Sub
    End If
    ' The capture counter belongs to the team HOLDING the flag, i.e. the other one.
    contadorBandera(IIf(team = 1, 2, 1)) = 0
    With UserList(UserIndex)
        .flags.tiene_bandera = 0
        Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(0, .Char.charindex))
        'Msg1718
        Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_ATENCION_BANDERA_EQUIPO_ENCUENTRA, colorEquipo(team) & "¬" & .pos.Map & "," & .pos.x & "," & .pos.y, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.TiraBandera", Erl)
End Sub

' NOTE: there is no GM-cancel routine here. /cancelarevento goes through HandleCancelarEvento ->
' CancelLobby, which refunds, returns the players and moves the lobby to Closed; the scenario is
' told through UpdateLobbyState and cleans itself up.
