Attribute VB_Name = "ModCaptura"
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
' Plan 27.001 (05.002 wave 4): CAPTURA_TIEMPO_ESPERA (90 s) was removed. Its only consumer was
' clsCaptura, which is gone. The registration window of the CTF is set by the lobby:
' ModLobby.WaitingForPlayersTime = 300000 ms (5 minutes).
#If DEBUGGING Then
    Public Const CAPTURA_TIEMPO_INICIO_RONDA = 10 '60 'Tiempo hasta que se inicia la ronda
#Else
    Public Const CAPTURA_TIEMPO_INICIO_RONDA = 60 '60 'Tiempo hasta que se inicia la ronda
#End If
Public Const CAPTURA_TIEMPO_BANDERA = 7 'Tiempo que tiene que estar el user con la bandera en la base
Public Const CAPTURA_TIEMPO_MUERTE = 5 'Tiempo que tarda para poder revivir cuando muere
Public Const CAPTURA_TIEMPO_POR_MUERTE = 2 'Multiplicador de tiempo x veces que murio
Public Const MAP_SALA_ESPERA       As Integer = 278

Public Const MAP_TEAM_1            As Integer = 275
Public Const X_TEAM_1              As Integer = 43
Public Const Y_TEAM_1              As Integer = 51
Public Const X_BANDERA_1           As Integer = 37
Public Const Y_BANDERA_1           As Integer = 51
Public Const MAP_TEAM_2            As Integer = 277
Public Const X_TEAM_2              As Integer = 65
Public Const Y_TEAM_2              As Integer = 55
Public Const X_BANDERA_2           As Integer = 71
Public Const Y_BANDERA_2           As Integer = 55
Public Const MIN_SALA_ESPERA_X     As Byte = 45
Public Const MIN_SALA_ESPERA_Y     As Byte = 68
Public Const MAX_SALA_ESPERA_X     As Byte = 56
Public Const MAX_SALA_ESPERA_Y     As Byte = 73
Public Const OBJ_CAPTURA_BANDERA_1 As Integer = 3674 'Estandarte Azul
Public Const OBJ_CAPTURA_BANDERA_2 As Integer = 3675 'Estandarte Rojo
' B1: round time limit bounds (seconds), see CtfClampRoundSeconds
Public Const CAPTURA_TIEMPO_MAX_RONDA_DEFAULT As Long = 600
Public Const CAPTURA_TIEMPO_MAX_RONDA_MIN     As Long = 60
Public Const CAPTURA_TIEMPO_MAX_RONDA_MAX     As Long = 3600
' Plan 27.001 (05.002 wave 4): the InstanciaCaptura singleton is gone. The CTF is an engine
' scenario (ScenarioCaptureTheFlag) that lives inside its lobby and is resolved through the map
' where the player stands, like every other scenario.
Public Function GetCapturaEscenario(ByVal UserIndex As Integer) As ScenarioCaptureTheFlag
    On Error GoTo ErrHandler
    Dim sc As IBaseScenario
    Set GetCapturaEscenario = Nothing
    If UserIndex <= 0 Then Exit Function
    Set sc = CustomScenarios.GetMap(UserList(UserIndex).pos.Map)
    If sc Is Nothing Then Exit Function
    If TypeOf sc Is ScenarioCaptureTheFlag Then
        Set GetCapturaEscenario = sc
    End If
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModCaptura.GetCapturaEscenario", Erl)
    Set GetCapturaEscenario = Nothing
End Function

' ---------------------------------------------------------------------------
' B1 (plan 2026-11-24-revision-eventos-automaticos): round time limit of the CTF.
' Pure helpers (no globals) so UnitTesting can cover them.
' ---------------------------------------------------------------------------
' Value read from Scenarios/CaptureTheFlag.ini [INIT] TiempoMaxRondaSeg. 0 (missing key, empty or
' not a number) means "use the default"; anything else is clamped to 60..3600.
Public Function CtfClampRoundSeconds(ByVal rawValue As Long) As Long
    If rawValue <= 0 Then
        CtfClampRoundSeconds = CAPTURA_TIEMPO_MAX_RONDA_DEFAULT
    ElseIf rawValue < CAPTURA_TIEMPO_MAX_RONDA_MIN Then
        CtfClampRoundSeconds = CAPTURA_TIEMPO_MAX_RONDA_MIN
    ElseIf rawValue > CAPTURA_TIEMPO_MAX_RONDA_MAX Then
        CtfClampRoundSeconds = CAPTURA_TIEMPO_MAX_RONDA_MAX
    Else
        CtfClampRoundSeconds = rawValue
    End If
End Function

' Who wins the event when the LAST round expires without a winner: the team with more rounds won.
' Returns 1 or 2, or 0 for a tie (the caller refunds everybody through the draw path).
Public Function CtfTimeoutWinner(ByVal roundsWonTeam1 As Integer, ByVal roundsWonTeam2 As Integer) As Integer
    If roundsWonTeam1 > roundsWonTeam2 Then
        CtfTimeoutWinner = 1
    ElseIf roundsWonTeam2 > roundsWonTeam1 Then
        CtfTimeoutWinner = 2
    Else
        CtfTimeoutWinner = 0
    End If
End Function
