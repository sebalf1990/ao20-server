Attribute VB_Name = "Unit_WorkFatigue"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Work fatigue and spell gate test suite (plan 2026-11-08-reemplazo-hambre-sed-energia)
' Pure state machine of modWorkFatigue with an injected clock, the spell cooldown
' rules of modSpellCooldowns and the MinLevel / 5000 guards of modHechizos.
' ==========================================================================
Private Const REST_MS As Long = 5000
Private Const TEST_NOW As Date = #6/15/2026 12:00:00 PM#

Public Function test_suite_work_fatigue() As Boolean
    Call UnitTesting.RunTest("test_wf_threshold_in_range", test_wf_threshold_in_range())
    Call UnitTesting.RunTest("test_wf_threshold_varies", test_wf_threshold_varies())
    Call UnitTesting.RunTest("test_wf_fatigues_after_threshold", test_wf_fatigues_after_threshold())
    Call UnitTesting.RunTest("test_wf_idle_gap_counts_nothing", test_wf_idle_gap_counts_nothing())
    Call UnitTesting.RunTest("test_wf_same_tick_no_double_count", test_wf_same_tick_no_double_count())
    Call UnitTesting.RunTest("test_wf_clock_crossing_2_31", test_wf_clock_crossing_2_31())
    Call UnitTesting.RunTest("test_wf_countdown_starts_on_unequip", test_wf_countdown_starts_on_unequip())
    Call UnitTesting.RunTest("test_wf_rest_done_after_unequip_time", test_wf_rest_done_after_unequip_time())
    Call UnitTesting.RunTest("test_wf_early_reequip_restarts", test_wf_early_reequip_restarts())
    Call UnitTesting.RunTest("test_wf_early_attempt_with_tool_restarts", test_wf_early_attempt_with_tool_restarts())
    Call UnitTesting.RunTest("test_wf_smelting_counts_from_warning", test_wf_smelting_counts_from_warning())
    Call UnitTesting.RunTest("test_wf_smelting_early_attempt_restarts", test_wf_smelting_early_attempt_restarts())
    Call UnitTesting.RunTest("test_wf_new_threshold_each_cycle", test_wf_new_threshold_each_cycle())
    Call UnitTesting.RunTest("test_wf_relog_persists_fatigue", test_wf_relog_persists_fatigue())
    Call UnitTesting.RunTest("test_wf_relog_keeps_partial_count", test_wf_relog_keeps_partial_count())
    Call UnitTesting.RunTest("test_wf_default_state_not_stored", test_wf_default_state_not_stored())
    Call UnitTesting.RunTest("test_wf_bad_saved_text_is_default", test_wf_bad_saved_text_is_default())
    Call UnitTesting.RunTest("test_wf_no_tool_in_wait_starts_rest", test_wf_no_tool_in_wait_starts_rest())
    Call UnitTesting.RunTest("test_wf_old_saved_state_discarded", test_wf_old_saved_state_discarded())
    Call UnitTesting.RunTest("test_wf_config_clamped", test_wf_config_clamped())
    Call UnitTesting.RunTest("test_spell_cooldown_old_stamp_discarded", test_spell_cooldown_old_stamp_discarded())
    Call UnitTesting.RunTest("test_spell_cooldown_active_window", test_spell_cooldown_active_window())
    Call UnitTesting.RunTest("test_spell_cooldown_fixed_ignores_elven_wood", test_spell_cooldown_fixed_ignores_elven_wood())
    Call UnitTesting.RunTest("test_spell_cooldown_store_take", test_spell_cooldown_store_take())
    Call UnitTesting.RunTest("test_spell_min_level_gate", test_spell_min_level_gate())
    Call UnitTesting.RunTest("test_spell_unavailable_5000_guard", test_spell_unavailable_5000_guard())
    test_suite_work_fatigue = True
End Function

' Registers work every stepMs from startTick until the state changes or maxSteps is reached.
' Returns the last event; endTick gets the tick of the last action.
Private Function RunUntilFatigue(ByRef s As t_WorkFatigue, ByVal startTick As Long, ByVal stepMs As Long, ByVal maxSteps As Long, ByRef endTick As Long) As e_WorkFatigueEvent
    Dim i As Long
    endTick = startTick
    For i = 1 To maxSteps
        endTick = endTick + stepMs
        RunUntilFatigue = WF_RegisterWork(s, endTick, True, 1, 1)
        If RunUntilFatigue = WF_EVT_FATIGUED Then Exit Function
    Next i
End Function

Private Function test_wf_threshold_in_range() As Boolean
    On Error GoTo Err_Handler
    test_wf_threshold_in_range = True
    Dim i As Long, v As Long
    For i = 1 To 300
        v = WF_DrawThresholdMs(12, 18)
        If v < 12 * 60000 Or v > 18 * 60000 Then test_wf_threshold_in_range = False: Exit Function
    Next i
    ' min = max is a fixed value
    If WF_DrawThresholdMs(2, 2) <> 120000 Then test_wf_threshold_in_range = False: Exit Function
    Exit Function
Err_Handler:
    test_wf_threshold_in_range = False
End Function

Private Function test_wf_threshold_varies() As Boolean
    On Error GoTo Err_Handler
    Dim first As Long, i As Long
    first = WF_DrawThresholdMs(12, 18)
    For i = 1 To 20
        If WF_DrawThresholdMs(12, 18) <> first Then test_wf_threshold_varies = True: Exit Function
    Next i
    Exit Function
Err_Handler:
    test_wf_threshold_varies = False
End Function

' 1 minute threshold, one action every 4 s: not fatigued at 56 s, fatigued at 60 s.
Private Function test_wf_fatigues_after_threshold() As Boolean
    On Error GoTo Err_Handler
    Dim s As t_WorkFatigue, endTick As Long, evt As e_WorkFatigueEvent
    Call WF_InitState(s)
    s.ThresholdMs = 60000
    ' first action only arms the clock (nothing counted yet)
    If WF_RegisterWork(s, 1000, True, 1, 1) <> WF_EVT_NONE Then Exit Function
    If s.ActiveMs <> 0 Then Exit Function
    evt = RunUntilFatigue(s, 1000, 4000, 14, endTick)
    ' 14 steps of 4 s = 56 s counted: still working
    If evt <> WF_EVT_NONE Then Exit Function
    If s.ActiveMs <> 56000 Then Exit Function
    evt = WF_RegisterWork(s, endTick + 4000, True, 1, 1)
    If evt <> WF_EVT_FATIGUED Then Exit Function
    If s.Phase <> WF_PHASE_WAIT_UNEQUIP Then Exit Function
    test_wf_fatigues_after_threshold = True
    Exit Function
Err_Handler:
    test_wf_fatigues_after_threshold = False
End Function

Private Function test_wf_idle_gap_counts_nothing() As Boolean
    On Error GoTo Err_Handler
    Dim s As t_WorkFatigue
    Call WF_InitState(s)
    s.ThresholdMs = 60000
    Call WF_RegisterWork(s, 1000, True, 1, 1)
    Call WF_RegisterWork(s, 5000, True, 1, 1)
    If s.ActiveMs <> 4000 Then Exit Function
    ' 10 minutes later: the pause is not work
    Call WF_RegisterWork(s, 5000 + 600000, True, 1, 1)
    If s.ActiveMs <> 4000 Then Exit Function
    test_wf_idle_gap_counts_nothing = True
    Exit Function
Err_Handler:
    test_wf_idle_gap_counts_nothing = False
End Function

Private Function test_wf_same_tick_no_double_count() As Boolean
    On Error GoTo Err_Handler
    Dim s As t_WorkFatigue
    Call WF_InitState(s)
    s.ThresholdMs = 60000
    Call WF_RegisterWork(s, 1000, True, 1, 1)
    Call WF_RegisterWork(s, 5000, True, 1, 1)
    Call WF_RegisterWork(s, 5000, True, 1, 1)
    Call WF_RegisterWork(s, 5000, True, 1, 1)
    If s.ActiveMs <> 4000 Then Exit Function
    test_wf_same_tick_no_double_count = True
    Exit Function
Err_Handler:
    test_wf_same_tick_no_double_count = False
End Function

' Ticks go from &H7FFFFFFF to &H80000000 (negative Long): counting must not break.
Private Function test_wf_clock_crossing_2_31() As Boolean
    On Error GoTo Err_Handler
    Dim s As t_WorkFatigue
    Call WF_InitState(s)
    s.ThresholdMs = 60000
    Call WF_RegisterWork(s, &H7FFFFFFE, True, 1, 1)
    Call WF_RegisterWork(s, &H80000000 + 3998, True, 1, 1)
    If s.ActiveMs <> 4000 Then Exit Function
    test_wf_clock_crossing_2_31 = True
    Exit Function
Err_Handler:
    test_wf_clock_crossing_2_31 = False
End Function

Private Sub MakeFatigued(ByRef s As t_WorkFatigue, ByVal needsUnequip As Boolean, ByVal atTick As Long)
    Call WF_InitState(s)
    s.ThresholdMs = 8000
    Call WF_RegisterWork(s, atTick - 4000, needsUnequip, 1, 1)
    Call WF_RegisterWork(s, atTick, needsUnequip, 1, 1)
    Call WF_RegisterWork(s, atTick + 4000, needsUnequip, 1, 1)
End Sub

' Tool trade: the 5 seconds do NOT run until the tool is dropped.
Private Function test_wf_countdown_starts_on_unequip() As Boolean
    On Error GoTo Err_Handler
    Dim s As t_WorkFatigue
    Call MakeFatigued(s, True, 10000)
    If s.Phase <> WF_PHASE_WAIT_UNEQUIP Then Exit Function
    ' a minute later, tool still in hand: still blocked, no rest ran
    If WF_Gate(s, 14000 + 60000, REST_MS, True) <> WF_EVT_BLOCKED_TIRED Then Exit Function
    If s.Phase <> WF_PHASE_WAIT_UNEQUIP Then Exit Function
    If Not WF_OnUnequip(s, 80000) Then Exit Function
    If s.Phase <> WF_PHASE_RESTING Then Exit Function
    If WF_Gate(s, 80000 + 4999, REST_MS, False) <> WF_EVT_BLOCKED_STILL Then Exit Function
    test_wf_countdown_starts_on_unequip = True
    Exit Function
Err_Handler:
    test_wf_countdown_starts_on_unequip = False
End Function

Private Function test_wf_rest_done_after_unequip_time() As Boolean
    On Error GoTo Err_Handler
    Dim s As t_WorkFatigue
    Call MakeFatigued(s, True, 10000)
    Call WF_OnUnequip(s, 20000)
    If WF_Gate(s, 20000 + REST_MS, REST_MS, False) <> WF_EVT_REST_DONE Then Exit Function
    If s.Phase <> WF_PHASE_WORKING Then Exit Function
    ' works again and the next threshold is drawn on the first action
    If s.ThresholdMs <> 0 Then Exit Function
    Call WF_RegisterWork(s, 30000, True, 1, 1)
    If s.ThresholdMs <> 60000 Then Exit Function
    test_wf_rest_done_after_unequip_time = True
    Exit Function
Err_Handler:
    test_wf_rest_done_after_unequip_time = False
End Function

' Equip again before the 5 s: still tired, and the 5 s run again from the NEXT unequip.
Private Function test_wf_early_reequip_restarts() As Boolean
    On Error GoTo Err_Handler
    Dim s As t_WorkFatigue
    Call MakeFatigued(s, True, 10000)
    Call WF_OnUnequip(s, 20000)
    If WF_OnEquip(s, 23000, REST_MS) <> WF_EVT_BLOCKED_STILL Then Exit Function
    If s.Phase <> WF_PHASE_WAIT_UNEQUIP Then Exit Function
    ' 20 s later, tool still equipped: the old countdown does not count
    If WF_Gate(s, 43000, REST_MS, True) <> WF_EVT_BLOCKED_TIRED Then Exit Function
    Call WF_OnUnequip(s, 50000)
    If WF_Gate(s, 54000, REST_MS, False) <> WF_EVT_BLOCKED_STILL Then Exit Function
    If WF_Gate(s, 55000, REST_MS, False) <> WF_EVT_REST_DONE Then Exit Function
    test_wf_early_reequip_restarts = True
    Exit Function
Err_Handler:
    test_wf_early_reequip_restarts = False
End Function

Private Function test_wf_early_attempt_with_tool_restarts() As Boolean
    On Error GoTo Err_Handler
    Dim s As t_WorkFatigue
    Call MakeFatigued(s, True, 10000)
    Call WF_OnUnequip(s, 20000)
    ' tried to work with the tool in hand 2 s later
    If WF_Gate(s, 22000, REST_MS, True) <> WF_EVT_BLOCKED_STILL Then Exit Function
    If s.Phase <> WF_PHASE_WAIT_UNEQUIP Then Exit Function
    ' even much later the rest has not happened
    If WF_Gate(s, 90000, REST_MS, True) <> WF_EVT_BLOCKED_TIRED Then Exit Function
    test_wf_early_attempt_with_tool_restarts = True
    Exit Function
Err_Handler:
    test_wf_early_attempt_with_tool_restarts = False
End Function

' Smelting has no tool: the 5 s run from the warning.
Private Function test_wf_smelting_counts_from_warning() As Boolean
    On Error GoTo Err_Handler
    Dim s As t_WorkFatigue
    Call MakeFatigued(s, False, 10000)
    If s.Phase <> WF_PHASE_RESTING Then Exit Function
    If WF_Gate(s, 14000 + 4999, REST_MS, False) <> WF_EVT_BLOCKED_STILL Then Exit Function
    ' the early attempt restarted the 5 s from 18999
    If WF_Gate(s, 18999 + 4999, REST_MS, False) <> WF_EVT_BLOCKED_STILL Then Exit Function
    test_wf_smelting_counts_from_warning = True
    Exit Function
Err_Handler:
    test_wf_smelting_counts_from_warning = False
End Function

Private Function test_wf_smelting_early_attempt_restarts() As Boolean
    On Error GoTo Err_Handler
    Dim s As t_WorkFatigue
    Call MakeFatigued(s, False, 10000)
    ' no attempt in between: rest over 5 s after the warning
    If WF_Gate(s, 14000 + REST_MS, REST_MS, False) <> WF_EVT_REST_DONE Then Exit Function
    test_wf_smelting_early_attempt_restarts = True
    Exit Function
Err_Handler:
    test_wf_smelting_early_attempt_restarts = False
End Function

' 6 full cycles: every one draws a threshold inside [1,2] minutes and they are not all equal.
Private Function test_wf_new_threshold_each_cycle() As Boolean
    On Error GoTo Err_Handler
    Dim s As t_WorkFatigue, cycle As Long, tick As Long, i As Long
    Dim thresholds(1 To 6) As Long
    Dim allEqual As Boolean
    allEqual = True
    Call WF_InitState(s)
    tick = 1000
    For cycle = 1 To 6
        Do
            tick = tick + 4000
            If WF_RegisterWork(s, tick, True, 1, 2) = WF_EVT_FATIGUED Then Exit Do
            i = i + 1
            If i > 1000 Then Exit Function
        Loop
        thresholds(cycle) = s.ThresholdMs
        If thresholds(cycle) < 60000 Or thresholds(cycle) > 120000 Then Exit Function
        If thresholds(cycle) <> thresholds(1) Then allEqual = False
        Call WF_OnUnequip(s, tick)
        tick = tick + REST_MS
        If WF_Poll(s, tick, REST_MS) <> WF_EVT_REST_DONE Then Exit Function
    Next cycle
    If allEqual Then Exit Function
    test_wf_new_threshold_each_cycle = True
    Exit Function
Err_Handler:
    test_wf_new_threshold_each_cycle = False
End Function

' Relog in the middle of the fatigue (warning given, tool equipped): comes back tired.
Private Function test_wf_relog_persists_fatigue() As Boolean
    On Error GoTo Err_Handler
    Dim s As t_WorkFatigue, back As t_WorkFatigue
    Call MakeFatigued(s, True, 10000)
    Call WF_StoreByName("Pepe_Test_Fatiga", s, TEST_NOW)
    If Not WF_LoadByName("pepe_test_fatiga", back, TEST_NOW) Then Exit Function
    If back.Phase <> WF_PHASE_WAIT_UNEQUIP Then Exit Function
    If Not back.NeedsUnequip Then Exit Function
    If WF_Gate(back, 99999, REST_MS, True) <> WF_EVT_BLOCKED_TIRED Then Exit Function
    test_wf_relog_persists_fatigue = True
    Exit Function
Err_Handler:
    test_wf_relog_persists_fatigue = False
End Function

' Relog with half the work done: the count is neither lost nor given away.
Private Function test_wf_relog_keeps_partial_count() As Boolean
    On Error GoTo Err_Handler
    Dim s As t_WorkFatigue, back As t_WorkFatigue
    Call WF_InitState(s)
    s.ThresholdMs = 60000
    Call WF_RegisterWork(s, 1000, True, 1, 1)
    Call WF_RegisterWork(s, 5000, True, 1, 1)
    Call WF_RegisterWork(s, 9000, True, 1, 1)
    Call WF_StoreByName("Pepe_Test_Parcial", s, TEST_NOW)
    If Not WF_LoadByName("Pepe_Test_Parcial", back, TEST_NOW) Then Exit Function
    If back.ActiveMs <> 8000 Or back.ThresholdMs <> 60000 Then Exit Function
    ' first action after the relog (long pause): counts nothing
    Call WF_RegisterWork(back, 9000 + 300000, True, 1, 1)
    If back.ActiveMs <> 8000 Then Exit Function
    test_wf_relog_keeps_partial_count = True
    Exit Function
Err_Handler:
    test_wf_relog_keeps_partial_count = False
End Function

Private Function test_wf_default_state_not_stored() As Boolean
    On Error GoTo Err_Handler
    Dim s As t_WorkFatigue, back As t_WorkFatigue
    Call WF_InitState(s)
    Call WF_StoreByName("Pepe_Test_Default", s, TEST_NOW)
    If WF_LoadByName("Pepe_Test_Default", back, TEST_NOW) Then Exit Function
    If Not WF_IsDefault(back) Then Exit Function
    test_wf_default_state_not_stored = True
    Exit Function
Err_Handler:
    test_wf_default_state_not_stored = False
End Function

Private Function test_wf_bad_saved_text_is_default() As Boolean
    On Error GoTo Err_Handler
    Dim s As t_WorkFatigue
    If WF_Deserialize("basura", s, TEST_NOW) Then Exit Function
    If Not WF_IsDefault(s) Then Exit Function
    If WF_Deserialize("9|1|1|1|1|1|45000.5", s, TEST_NOW) Then Exit Function
    If Not WF_IsDefault(s) Then Exit Function
    test_wf_bad_saved_text_is_default = True
    Exit Function
Err_Handler:
    test_wf_bad_saved_text_is_default = False
End Function

' Blacksmith without a tool in hand: nothing to drop, the rest starts at the next attempt.
Private Function test_wf_no_tool_in_wait_starts_rest() As Boolean
    On Error GoTo Err_Handler
    Dim s As t_WorkFatigue
    Call MakeFatigued(s, True, 10000)
    If s.Phase <> WF_PHASE_WAIT_UNEQUIP Then Exit Function
    If WF_Gate(s, 20000, REST_MS, False) <> WF_EVT_BLOCKED_STILL Then Exit Function
    If s.Phase <> WF_PHASE_RESTING Then Exit Function
    If WF_Gate(s, 24999, REST_MS, False) <> WF_EVT_BLOCKED_STILL Then Exit Function
    If WF_Gate(s, 25000, REST_MS, False) <> WF_EVT_REST_DONE Then Exit Function
    test_wf_no_tool_in_wait_starts_rest = True
    Exit Function
Err_Handler:
    test_wf_no_tool_in_wait_starts_rest = False
End Function

' A state saved more than 60 minutes ago is dropped (a long break counts as rest).
Private Function test_wf_old_saved_state_discarded() As Boolean
    On Error GoTo Err_Handler
    Dim s As t_WorkFatigue, back As t_WorkFatigue
    Call MakeFatigued(s, True, 10000)
    Call WF_StoreByName("Pepe_Test_Viejo", s, TEST_NOW)
    If Not WF_LoadByName("Pepe_Test_Viejo", back, DateAdd("n", 59, TEST_NOW)) Then Exit Function
    Call WF_StoreByName("Pepe_Test_Viejo", s, TEST_NOW)
    If WF_LoadByName("Pepe_Test_Viejo", back, DateAdd("n", 61, TEST_NOW)) Then Exit Function
    If Not WF_IsDefault(back) Then Exit Function
    test_wf_old_saved_state_discarded = True
    Exit Function
Err_Handler:
    test_wf_old_saved_state_discarded = False
End Function

Private Function test_wf_config_clamped() As Boolean
    On Error GoTo Err_Handler
    Dim oldMin As Long, oldMax As Long, oldRest As Long
    oldMin = WorkFatigue_MinMinutes()
    oldMax = WorkFatigue_MaxMinutes()
    oldRest = WorkFatigue_RestSeconds()
    Call WorkFatigue_SetConfig(0, 99999, 99999)
    If WorkFatigue_MinMinutes() <> WF_DEFAULT_MIN_MINUTES Then GoTo Restore
    If WorkFatigue_MaxMinutes() <> 600 Then GoTo Restore
    If WorkFatigue_RestSeconds() <> 600 Then GoTo Restore
    Call WorkFatigue_SetConfig(900, 5, 5)
    If WorkFatigue_MinMinutes() <> 600 Or WorkFatigue_MaxMinutes() <> 600 Then GoTo Restore
    Call WorkFatigue_SetConfig(3, 7, 5)
    If WorkFatigue_MinMinutes() <> 3 Or WorkFatigue_MaxMinutes() <> 7 Then GoTo Restore
    test_wf_config_clamped = True
Restore:
    Call WorkFatigue_SetConfig(oldMin, oldMax, oldRest)
    Exit Function
Err_Handler:
    test_wf_config_clamped = False
End Function

Private Function test_spell_cooldown_old_stamp_discarded() As Boolean
    On Error GoTo Err_Handler
    Call SpellCooldown_Store("Pepe_Test_Cd2", 14, 777, TEST_NOW)
    If SpellCooldown_Take("Pepe_Test_Cd2", 14, DateAdd("n", 59, TEST_NOW)) <> 777 Then Exit Function
    Call SpellCooldown_Store("Pepe_Test_Cd2", 14, 777, TEST_NOW)
    If SpellCooldown_Take("Pepe_Test_Cd2", 14, DateAdd("n", 61, TEST_NOW)) <> 0 Then Exit Function
    test_spell_cooldown_old_stamp_discarded = True
    Exit Function
Err_Handler:
    test_spell_cooldown_old_stamp_discarded = False
End Function

Private Function test_spell_cooldown_active_window() As Boolean
    On Error GoTo Err_Handler
    If Not SpellCooldown_IsActive(1000, 1000 + 19999, 20) Then Exit Function
    If SpellCooldown_IsActive(1000, 1000 + 20000, 20) Then Exit Function
    If SpellCooldown_IsActive(0, 5000, 20) Then Exit Function
    If SpellCooldown_IsActive(1000, 2000, 0) Then Exit Function
    ' stamp taken just before the 2^31 crossing
    If Not SpellCooldown_IsActive(&H7FFFFFF0, &H80000000 + 100, 20) Then Exit Function
    test_spell_cooldown_active_window = True
    Exit Function
Err_Handler:
    test_spell_cooldown_active_window = False
End Function

Private Function test_spell_cooldown_fixed_ignores_elven_wood() As Boolean
    On Error GoTo Err_Handler
    If SpellCooldown_EffectiveSeconds(20, True, 2) <> 20 Then Exit Function
    If SpellCooldown_EffectiveSeconds(20, False, 0) <> 20 Then Exit Function
    If SpellCooldown_EffectiveSeconds(20, False, 1) <> 10 Then Exit Function
    If SpellCooldown_EffectiveSeconds(20, False, 2) <> 5 Then Exit Function
    test_spell_cooldown_fixed_ignores_elven_wood = True
    Exit Function
Err_Handler:
    test_spell_cooldown_fixed_ignores_elven_wood = False
End Function

Private Function test_spell_cooldown_store_take() As Boolean
    On Error GoTo Err_Handler
    Call SpellCooldown_Store("Pepe_Test_Cd", 14, 123456, TEST_NOW)
    If SpellCooldown_Take("pepe_test_cd", 15, TEST_NOW) <> 0 Then Exit Function
    If SpellCooldown_Take("pepe_test_cd", 14, TEST_NOW) <> 123456 Then Exit Function
    ' taken once: gone
    If SpellCooldown_Take("pepe_test_cd", 14, TEST_NOW) <> 0 Then Exit Function
    test_spell_cooldown_store_take = True
    Exit Function
Err_Handler:
    test_spell_cooldown_store_take = False
End Function

Private Function test_spell_min_level_gate() As Boolean
    On Error GoTo Err_Handler
    ' 0 = no requirement
    If Not SpellMeetsMinLevel(1, 0) Then Exit Function
    If SpellMeetsMinLevel(16, 17) Then Exit Function
    If Not SpellMeetsMinLevel(17, 17) Then Exit Function
    If Not SpellMeetsMinLevel(30, 24) Then Exit Function
    test_spell_min_level_gate = True
    Exit Function
Err_Handler:
    test_spell_min_level_gate = False
End Function

Private Function test_spell_unavailable_5000_guard() As Boolean
    On Error GoTo Err_Handler
    If Not IsSpellUnavailable(5000) Then Exit Function
    If Not IsSpellUnavailable(6000) Then Exit Function
    If IsSpellUnavailable(4999) Then Exit Function
    If IsSpellUnavailable(0) Then Exit Function
    test_spell_unavailable_5000_guard = True
    Exit Function
Err_Handler:
    test_spell_unavailable_5000_guard = False
End Function

#End If
