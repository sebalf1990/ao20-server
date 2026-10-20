Attribute VB_Name = "ModPhoenix"
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
Option Explicit

Public PhoenixLastSpawnTimestamp As Long
Public PhoenixMapPool() As Integer
Private m_LastPhoenixSpawnAttempt As Long
Public IsPhoenixAlive As Boolean
Public Const PHOENIX_NPC_INDEX = 1373
Private Const PHOENIX_BROADCAST_MSG_ID = 2159
Private PhoenixSpawnPosition As t_WorldPos

Public Sub MaybeSpawnFenix()
    On Error GoTo MaybeSpawnFenix_Err
    If UBound(PhoenixMapPool) < 1 Then Exit Sub
    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()
    If TicksElapsed(m_LastPhoenixSpawnAttempt, nowRaw) < IntervalPhoenixSpawn Then Exit Sub
    m_LastPhoenixSpawnAttempt = nowRaw
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    If Not IsPhoenixAlive Then
        ' Fenix (plan 02.001 Ola 6): spawn en un mapa random que NO sea dungeon (como el oficial)
        Dim cand As Integer, intentos As Integer, m As Integer
        Dim nIndex As Integer
        cand = 0
        For intentos = 1 To 50
            m = RandomNumber(1, NumMaps)
            If MapInfo(m).zone <> "DUNGEON" Then
                cand = m
                Exit For
            End If
        Next intentos
        If cand = 0 Then Exit Sub
        PhoenixSpawnPosition.Map = cand
        ' Solo marcar vivo si el spawn ubico al Fenix (ClosestLegalPos devuelve 0 si no hay tile valido)
        nIndex = SpawnNpc(PHOENIX_NPC_INDEX, PhoenixSpawnPosition, 1, False, False, 0, SND_WARP)
        If nIndex <> 0 Then
            Call modSendData.SendData(ToAll, 0, PrepareMessageLocaleMsg(PHOENIX_BROADCAST_MSG_ID, vbNullString, e_TextChannel.TEXTCHANNEL_FACTION, e_FontTypeNames.FONTTYPE_CITIZEN))
            Call modSendData.SendData(ToAll, 0, PrepareMessagePlayWave(e_SoundEffects.PhoenixRebirth, NO_3D_SOUND, NO_3D_SOUND, False, 0))
            IsPhoenixAlive = True
        End If
    End If
    Call PerformTimeLimitCheck(PerformanceTimer, "ModPhoenix.MaybeSpawnFenix", 100)
    Exit Sub
MaybeSpawnFenix_Err:
    Call TraceError(Err.Number, Err.Description, "ModPhoenix.MaybeSpawnFenix", Erl)
End Sub


Public Sub LoadPhoenixModule()
    On Error GoTo LoadPhoenixModule_Err
    m_LastPhoenixSpawnAttempt = GetTickCountRaw()
    If Not FileExist(DatPath & "PhoenixMapPool.dat", vbArchive) Then
        Debug.Assert False
        Call LogError("Missing file PhoenixMapPool.Dat")
        Exit Sub
    End If
    Dim IniFile     As clsIniManager
    Set IniFile = New clsIniManager
    Call IniFile.Initialize(DatPath & "PhoenixMapPool.dat")
    Dim MaxPhoenixMaps
    MaxPhoenixMaps = val(IniFile.GetValue("INIT", "MaxPhoenixMaps"))
    If MaxPhoenixMaps < 1 Then
        Call LogError("PhoenixMapPool.dat: MaxPhoenixMaps invalido: " & MaxPhoenixMaps)
        Exit Sub
    End If
    ReDim Preserve PhoenixMapPool(1 To MaxPhoenixMaps)
    Dim i As Integer
    For i = 1 To MaxPhoenixMaps
        PhoenixMapPool(i) = CLng(val(IniFile.GetValue("Maps", "Map" & i)))
    Next i
    PhoenixSpawnPosition.x = 50
    PhoenixSpawnPosition.y = 50
    Exit Sub
LoadPhoenixModule_Err:
    Call TraceError(Err.Number, Err.Description, "ModPhoenix.LoadPhoenixModule", Erl)
End Sub
