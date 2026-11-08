Attribute VB_Name = "modWorkFatigue"
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
' ==========================================================================
' Work fatigue (plan 2026-11-08-reemplazo-hambre-sed-energia, E2).
'
' Replaces the energy bar as the brake on work macros. A character that works
' (fishing, logging, mining, smelting, blacksmith/tailor/alchemy/carpentry
' crafting) for a random time between MinMinutos and MaxMinutos of ACTIVE work
' gets a warning and the work stops. Tool trades must UNEQUIP the tool; the
' rest countdown (SegundosDescanso) starts when the tool is unequipped. Smelting
' has no tool: the countdown starts at the warning. Equipping the tool again (or
' trying to work) before the rest is over restarts the countdown.
'
' Toggle: work_fatigue (feature_toggle.ini). Params: intervalos.ini [WORKFATIGUE]
' MinMinutos / MaxMinutos / SegundosDescanso.
'
' State lives in server memory keyed by character name (survives a relog, is
' lost when the server restarts; no database column). The state machine
' (WF_* functions) is pure and takes the clock as a parameter so it can be
' unit tested (Tests/Unit_WorkFatigue.bas). The WorkFatigue_* functions are the
' UserIndex-level wrappers used by the game code.
' ==========================================================================

' Cached feature toggles (refreshed by SetFeatureToggle / LoadFeatureToggles).
Public WorkFatigueEnabled          As Boolean
Public HungerThirstStaminaDisabled As Boolean

Public Const WORK_FATIGUE_TOGGLE As String = "work_fatigue"
Public Const HUNGER_THIRST_STAMINA_TOGGLE As String = "disable_hunger_thirst_stamina"

Public Const WF_DEFAULT_MIN_MINUTES As Long = 12
Public Const WF_DEFAULT_MAX_MINUTES As Long = 18
Public Const WF_DEFAULT_REST_SECONDS As Long = 5
Public Const WF_CONFIG_MAX As Long = 600
' Saved states older than this are dropped on login (a long break counts as rest).
Public Const WF_STATE_MAX_AGE_MINUTES As Long = 60
' Longest pause between two executed work actions that still counts as continuous work.
Public Const WF_MAX_GAP_MS As Long = 15000
' Minimum time between two repeated "still tired" messages (macro spam guard).
Public Const WF_MSG_THROTTLE_MS As Long = 1500

Public Enum e_WorkFatiguePhase
    WF_PHASE_WORKING = 0
    WF_PHASE_WAIT_UNEQUIP = 1
    WF_PHASE_RESTING = 2
End Enum

Public Enum e_WorkFatigueEvent
    WF_EVT_NONE = 0
    WF_EVT_FATIGUED = 1
    WF_EVT_BLOCKED_TIRED = 2
    WF_EVT_BLOCKED_STILL = 3
    WF_EVT_REST_DONE = 4
End Enum

Public Enum e_WorkFatigueKind
    WF_KIND_TOOL_AUTO = 1
    WF_KIND_SMELT = 2
    WF_KIND_CRAFT = 3
End Enum

Public Type t_WorkFatigue
    Phase As Long
    ThresholdMs As Long
    ActiveMs As Long
    LastWorkTick As Long
    RestStart As Long
    NeedsUnequip As Boolean
    LastMsgTick As Long
End Type

Private m_MinMinutes   As Long
Private m_MaxMinutes   As Long
Private m_RestSeconds  As Long
Private m_States()     As t_WorkFatigue
Private m_Capacity     As Long
Private m_Saved        As Dictionary

' --------------------------------------------------------------------------
' Config and toggle cache
' --------------------------------------------------------------------------
Public Sub WorkFatigue_SetConfig(ByVal minMinutes As Long, ByVal maxMinutes As Long, ByVal restSeconds As Long)
    ' Clamp to 1..600 so a typo in the ini can neither disable the fatigue nor overflow a Long.
    If minMinutes <= 0 Then minMinutes = WF_DEFAULT_MIN_MINUTES
    If minMinutes > WF_CONFIG_MAX Then minMinutes = WF_CONFIG_MAX
    If maxMinutes < minMinutes Then maxMinutes = minMinutes
    If maxMinutes > WF_CONFIG_MAX Then maxMinutes = WF_CONFIG_MAX
    If restSeconds <= 0 Then restSeconds = WF_DEFAULT_REST_SECONDS
    If restSeconds > WF_CONFIG_MAX Then restSeconds = WF_CONFIG_MAX
    m_MinMinutes = minMinutes
    m_MaxMinutes = maxMinutes
    m_RestSeconds = restSeconds
End Sub

Public Function WorkFatigue_MinMinutes() As Long
    WorkFatigue_MinMinutes = IIf(m_MinMinutes > 0, m_MinMinutes, WF_DEFAULT_MIN_MINUTES)
End Function

Public Function WorkFatigue_MaxMinutes() As Long
    WorkFatigue_MaxMinutes = IIf(m_MaxMinutes > 0, m_MaxMinutes, WF_DEFAULT_MAX_MINUTES)
End Function

Public Function WorkFatigue_RestSeconds() As Long
    WorkFatigue_RestSeconds = IIf(m_RestSeconds > 0, m_RestSeconds, WF_DEFAULT_REST_SECONDS)
End Function

Public Sub WorkFatigue_ResetToggleCache()
    WorkFatigueEnabled = False
    HungerThirstStaminaDisabled = False
End Sub

Public Sub WorkFatigue_OnToggleChanged(ByVal toggleName As String, ByVal state As Boolean)
    Select Case toggleName
        Case WORK_FATIGUE_TOGGLE
            WorkFatigueEnabled = state
        Case HUNGER_THIRST_STAMINA_TOGGLE
            HungerThirstStaminaDisabled = state
    End Select
End Sub

' --------------------------------------------------------------------------
' Pure state machine (clock injected)
' --------------------------------------------------------------------------
Public Sub WF_InitState(ByRef s As t_WorkFatigue)
    s.Phase = WF_PHASE_WORKING
    s.ThresholdMs = 0
    s.ActiveMs = 0
    s.LastWorkTick = 0
    s.RestStart = 0
    s.NeedsUnequip = False
    s.LastMsgTick = 0
End Sub

Public Function WF_DrawThresholdMs(ByVal minMinutes As Long, ByVal maxMinutes As Long) As Long
    If minMinutes <= 0 Then minMinutes = WF_DEFAULT_MIN_MINUTES
    If maxMinutes < minMinutes Then maxMinutes = minMinutes
    WF_DrawThresholdMs = RandomNumber(minMinutes * 60000, maxMinutes * 60000)
End Function

' Completes the rest when its time is over. Returns WF_EVT_REST_DONE once.
Public Function WF_Poll(ByRef s As t_WorkFatigue, ByVal nowTick As Long, ByVal restMs As Long) As e_WorkFatigueEvent
    WF_Poll = WF_EVT_NONE
    If s.Phase <> WF_PHASE_RESTING Then Exit Function
    If TicksElapsed(s.RestStart, nowTick) >= restMs Then
        s.Phase = WF_PHASE_WORKING
        s.ThresholdMs = 0
        s.ActiveMs = 0
        s.LastWorkTick = 0
        s.NeedsUnequip = False
        WF_Poll = WF_EVT_REST_DONE
    End If
End Function

' Gate before a work action. NONE / REST_DONE = the work may go on.
' BLOCKED_TIRED = fatigued, tool not dropped yet. BLOCKED_STILL = the rest is not over.
Public Function WF_Gate(ByRef s As t_WorkFatigue, ByVal nowTick As Long, ByVal restMs As Long, ByVal toolEquipped As Boolean) As e_WorkFatigueEvent
    WF_Gate = WF_Poll(s, nowTick, restMs)
    Select Case s.Phase
        Case WF_PHASE_WORKING
            Exit Function
        Case WF_PHASE_WAIT_UNEQUIP
            If toolEquipped Then
                WF_Gate = WF_EVT_BLOCKED_TIRED
            Else
                ' No tool in hand (a blacksmith can craft without one): nothing to drop, the rest starts now.
                s.Phase = WF_PHASE_RESTING
                s.RestStart = nowTick
                WF_Gate = WF_EVT_BLOCKED_STILL
            End If
        Case WF_PHASE_RESTING
            If s.NeedsUnequip Then
                ' A tool in hand means it was equipped again before the rest was over.
                If toolEquipped Then s.Phase = WF_PHASE_WAIT_UNEQUIP
            Else
                s.RestStart = nowTick
            End If
            WF_Gate = WF_EVT_BLOCKED_STILL
    End Select
End Function

' Counts one EXECUTED work action. Only the time since the previous executed action
' counts, and only when the pause is short: a second call in the same millisecond adds
' nothing (no double counting) and an idle gap or a relog does not give free minutes.
Public Function WF_RegisterWork(ByRef s As t_WorkFatigue, ByVal nowTick As Long, ByVal needsUnequip As Boolean, ByVal minMinutes As Long, ByVal maxMinutes As Long) As e_WorkFatigueEvent
    WF_RegisterWork = WF_EVT_NONE
    If s.Phase <> WF_PHASE_WORKING Then Exit Function
    If s.ThresholdMs <= 0 Then s.ThresholdMs = WF_DrawThresholdMs(minMinutes, maxMinutes)
    If s.LastWorkTick <> 0 Then
        Dim delta As Double
        delta = TicksElapsed(s.LastWorkTick, nowTick)
        If delta <= WF_MAX_GAP_MS Then s.ActiveMs = s.ActiveMs + CLng(delta)
    End If
    s.LastWorkTick = NonZeroTickFrom(nowTick)
    If s.ActiveMs >= s.ThresholdMs Then
        s.Phase = IIf(needsUnequip, WF_PHASE_WAIT_UNEQUIP, WF_PHASE_RESTING)
        s.NeedsUnequip = needsUnequip
        s.RestStart = nowTick
        s.ActiveMs = 0
        s.LastWorkTick = 0
        WF_RegisterWork = WF_EVT_FATIGUED
    End If
End Function

' The player dropped the work tool: the rest countdown starts now.
Public Function WF_OnUnequip(ByRef s As t_WorkFatigue, ByVal nowTick As Long) As Boolean
    If s.Phase = WF_PHASE_WAIT_UNEQUIP And s.NeedsUnequip Then
        s.Phase = WF_PHASE_RESTING
        s.RestStart = nowTick
        WF_OnUnequip = True
    End If
End Function

' The player equipped a work tool. Before the rest is over the countdown is cancelled
' and restarts on the next unequip.
Public Function WF_OnEquip(ByRef s As t_WorkFatigue, ByVal nowTick As Long, ByVal restMs As Long) As e_WorkFatigueEvent
    WF_OnEquip = WF_Poll(s, nowTick, restMs)
    If s.Phase = WF_PHASE_RESTING And s.NeedsUnequip Then
        s.Phase = WF_PHASE_WAIT_UNEQUIP
        WF_OnEquip = WF_EVT_BLOCKED_STILL
    End If
End Function

' Compact text form used to keep the state of a character that is offline.
' The last field is the wall-clock time of the save (days as a Double, "." decimal).
Public Function WF_Serialize(ByRef s As t_WorkFatigue, ByVal savedAt As Date) As String
    WF_Serialize = s.Phase & "|" & s.ThresholdMs & "|" & s.ActiveMs & "|" & s.LastWorkTick & "|" & s.RestStart & "|" & IIf(s.NeedsUnequip, "1", "0") & "|" & Trim$(Str$(CDbl(savedAt)))
End Function

' False (and a default state) when the text is corrupt or older than WF_STATE_MAX_AGE_MINUTES.
Public Function WF_Deserialize(ByVal Text As String, ByRef s As t_WorkFatigue, ByVal nowDate As Date) As Boolean
    On Error GoTo WF_Deserialize_Err
    Dim parts() As String
    Dim ageMinutes As Double
    Call WF_InitState(s)
    parts = Split(Text, "|")
    If UBound(parts) <> 6 Then Exit Function
    s.Phase = val(parts(0))
    s.ThresholdMs = val(parts(1))
    s.ActiveMs = val(parts(2))
    s.LastWorkTick = CLng(parts(3))
    s.RestStart = CLng(parts(4))
    s.NeedsUnequip = (parts(5) = "1")
    ageMinutes = (CDbl(nowDate) - val(parts(6))) * 1440#
    If s.Phase < WF_PHASE_WORKING Or s.Phase > WF_PHASE_RESTING Or ageMinutes > WF_STATE_MAX_AGE_MINUTES Then
        Call WF_InitState(s)
        Exit Function
    End If
    WF_Deserialize = True
    Exit Function
WF_Deserialize_Err:
    Call WF_InitState(s)
    WF_Deserialize = False
End Function

Public Function WF_IsDefault(ByRef s As t_WorkFatigue) As Boolean
    WF_IsDefault = (s.Phase = WF_PHASE_WORKING And s.ActiveMs = 0 And s.ThresholdMs = 0)
End Function

' --------------------------------------------------------------------------
' Offline store (by character name) - survives a relog, not a server restart
' --------------------------------------------------------------------------
Public Sub WF_StoreByName(ByVal charName As String, ByRef s As t_WorkFatigue, ByVal savedAt As Date)
    If m_Saved Is Nothing Then Set m_Saved = New Dictionary
    Dim key As String
    key = UCase$(charName)
    If m_Saved.Exists(key) Then m_Saved.Remove key
    If Not WF_IsDefault(s) Then Call m_Saved.Add(key, WF_Serialize(s, savedAt))
End Sub

Public Function WF_LoadByName(ByVal charName As String, ByRef s As t_WorkFatigue, ByVal nowDate As Date) As Boolean
    Call WF_InitState(s)
    If m_Saved Is Nothing Then Exit Function
    Dim key As String
    key = UCase$(charName)
    If Not m_Saved.Exists(key) Then Exit Function
    WF_LoadByName = WF_Deserialize(CStr(m_Saved.Item(key)), s, nowDate)
    m_Saved.Remove key
End Function

' --------------------------------------------------------------------------
' UserIndex level wrappers (game code)
' Every public wrapper has its own error handler: an error here must never abort
' the RunAutomatedActions loop or a work action (the gate fails open).
' --------------------------------------------------------------------------
Private Sub EnsureSlot(ByVal UserIndex As Integer)
    If UserIndex <= m_Capacity Then Exit Sub
    Dim oldCapacity As Long
    oldCapacity = m_Capacity
    m_Capacity = UserIndex + 64
    If oldCapacity = 0 Then
        ReDim m_States(1 To m_Capacity) As t_WorkFatigue
    Else
        ReDim Preserve m_States(1 To m_Capacity) As t_WorkFatigue
    End If
End Sub

Private Function GetRestMs() As Long
    GetRestMs = WorkFatigue_RestSeconds() * 1000
End Function

Private Sub StopWork(ByVal UserIndex As Integer, ByVal kind As e_WorkFatigueKind)
    If kind = WF_KIND_CRAFT Or UserList(UserIndex).flags.UsandoMacro Then
        Call WriteMacroTrabajoToggle(UserIndex, False)
    End If
    Call ResetUserAutomatedActions(UserIndex)
End Sub

Private Function ShouldNotify(ByVal UserIndex As Integer, ByVal nowTick As Long) As Boolean
    If m_States(UserIndex).LastMsgTick <> 0 Then
        If TicksElapsed(m_States(UserIndex).LastMsgTick, nowTick) < WF_MSG_THROTTLE_MS Then Exit Function
    End If
    m_States(UserIndex).LastMsgTick = NonZeroTickFrom(nowTick)
    ShouldNotify = True
End Function

Private Sub SendRestedMessage(ByVal UserIndex As Integer)
    Call WriteLocaleMsg(UserIndex, MSG_WORK_FATIGUE_RESTED, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
End Sub

' Msg9020 takes the rest seconds as parameter 1.
Private Sub SendTiredMessage(ByVal UserIndex As Integer)
    Call WriteLocaleMsg(UserIndex, MSG_WORK_FATIGUE_TIRED, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO, WorkFatigue_RestSeconds())
End Sub

' Loads the state of the character that is logging in (also clears what a previous
' occupant of the UserIndex left behind). States saved more than WF_STATE_MAX_AGE_MINUTES
' ago are discarded: a long break counts as rest.
Public Sub WorkFatigue_OnLogin(ByVal UserIndex As Integer)
    On Error GoTo WorkFatigue_OnLogin_Err
    Call EnsureSlot(UserIndex)
    Call WF_LoadByName(UserList(UserIndex).name, m_States(UserIndex), Now)
    Exit Sub
WorkFatigue_OnLogin_Err:
    Call TraceError(Err.Number, Err.Description, "modWorkFatigue.WorkFatigue_OnLogin", Erl)
End Sub

Public Sub WorkFatigue_OnLogout(ByVal UserIndex As Integer)
    On Error GoTo WorkFatigue_OnLogout_Err
    If UserIndex > m_Capacity Then Exit Sub
    If LenB(UserList(UserIndex).name) = 0 Then Exit Sub
    Call WF_StoreByName(UserList(UserIndex).name, m_States(UserIndex), Now)
    Call WF_InitState(m_States(UserIndex))
    Exit Sub
WorkFatigue_OnLogout_Err:
    Call TraceError(Err.Number, Err.Description, "modWorkFatigue.WorkFatigue_OnLogout", Erl)
End Sub

' Gate used before a work action. False = do not work (the work was stopped and the
' player was told why).
Public Function WorkFatigue_CanWork(ByVal UserIndex As Integer, ByVal kind As e_WorkFatigueKind) As Boolean
    On Error GoTo WorkFatigue_CanWork_Err
    WorkFatigue_CanWork = True
    If Not WorkFatigueEnabled Then Exit Function
    Call EnsureSlot(UserIndex)
    Dim nowTick As Long
    Dim evt     As e_WorkFatigueEvent
    nowTick = GetTickCountRaw()
    evt = WF_Gate(m_States(UserIndex), nowTick, GetRestMs(), UserList(UserIndex).invent.EquippedWorkingToolObjIndex > 0)
    Select Case evt
        Case WF_EVT_REST_DONE
            Call SendRestedMessage(UserIndex)
        Case WF_EVT_BLOCKED_TIRED
            WorkFatigue_CanWork = False
            Call StopWork(UserIndex, kind)
            If ShouldNotify(UserIndex, nowTick) Then Call SendTiredMessage(UserIndex)
        Case WF_EVT_BLOCKED_STILL
            WorkFatigue_CanWork = False
            Call StopWork(UserIndex, kind)
            If ShouldNotify(UserIndex, nowTick) Then
                Call WriteLocaleMsg(UserIndex, MSG_WORK_FATIGUE_STILL_TIRED, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO, WorkFatigue_RestSeconds())
            End If
    End Select
    Exit Function
WorkFatigue_CanWork_Err:
    WorkFatigue_CanWork = True
    Call TraceError(Err.Number, Err.Description, "modWorkFatigue.WorkFatigue_CanWork", Erl)
End Function

' Call once per work action that was really executed (not for rejected attempts).
Public Sub WorkFatigue_RegisterWork(ByVal UserIndex As Integer, ByVal kind As e_WorkFatigueKind)
    On Error GoTo WorkFatigue_RegisterWork_Err
    If Not WorkFatigueEnabled Then Exit Sub
    Call EnsureSlot(UserIndex)
    Dim evt As e_WorkFatigueEvent
    evt = WF_RegisterWork(m_States(UserIndex), GetTickCountRaw(), kind <> WF_KIND_SMELT, WorkFatigue_MinMinutes(), WorkFatigue_MaxMinutes())
    If evt <> WF_EVT_FATIGUED Then Exit Sub
    Call StopWork(UserIndex, kind)
    m_States(UserIndex).LastMsgTick = NonZeroTickFrom(GetTickCountRaw())
    If kind = WF_KIND_SMELT Then
        Call WriteLocaleMsg(UserIndex, MSG_WORK_FATIGUE_SMELT, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO, WorkFatigue_RestSeconds())
    Else
        Call SendTiredMessage(UserIndex)
    End If
    Exit Sub
WorkFatigue_RegisterWork_Err:
    Call TraceError(Err.Number, Err.Description, "modWorkFatigue.WorkFatigue_RegisterWork", Erl)
End Sub

Public Sub WorkFatigue_OnToolUnequipped(ByVal UserIndex As Integer)
    On Error GoTo WorkFatigue_OnToolUnequipped_Err
    If Not WorkFatigueEnabled Then Exit Sub
    Call EnsureSlot(UserIndex)
    Call WF_OnUnequip(m_States(UserIndex), GetTickCountRaw())
    Exit Sub
WorkFatigue_OnToolUnequipped_Err:
    Call TraceError(Err.Number, Err.Description, "modWorkFatigue.WorkFatigue_OnToolUnequipped", Erl)
End Sub

Public Sub WorkFatigue_OnToolEquipped(ByVal UserIndex As Integer)
    On Error GoTo WorkFatigue_OnToolEquipped_Err
    If Not WorkFatigueEnabled Then Exit Sub
    Call EnsureSlot(UserIndex)
    Dim nowTick As Long
    Dim evt     As e_WorkFatigueEvent
    nowTick = GetTickCountRaw()
    evt = WF_OnEquip(m_States(UserIndex), nowTick, GetRestMs())
    Select Case evt
        Case WF_EVT_REST_DONE
            Call SendRestedMessage(UserIndex)
        Case WF_EVT_BLOCKED_STILL
            If ShouldNotify(UserIndex, nowTick) Then
                Call WriteLocaleMsg(UserIndex, MSG_WORK_FATIGUE_STILL_TIRED, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO, WorkFatigue_RestSeconds())
            End If
    End Select
    Exit Sub
WorkFatigue_OnToolEquipped_Err:
    Call TraceError(Err.Number, Err.Description, "modWorkFatigue.WorkFatigue_OnToolEquipped", Erl)
End Sub

' Periodic call (RunAutomatedActions): tells the player when the rest is over.
Public Sub WorkFatigue_Update(ByVal UserIndex As Integer)
    On Error GoTo WorkFatigue_Update_Err
    If Not WorkFatigueEnabled Then Exit Sub
    If UserIndex > m_Capacity Then Exit Sub
    If m_States(UserIndex).Phase <> WF_PHASE_RESTING Then Exit Sub
    If WF_Poll(m_States(UserIndex), GetTickCountRaw(), GetRestMs()) = WF_EVT_REST_DONE Then
        Call SendRestedMessage(UserIndex)
    End If
    Exit Sub
WorkFatigue_Update_Err:
    Call TraceError(Err.Number, Err.Description, "modWorkFatigue.WorkFatigue_Update", Erl)
End Sub
