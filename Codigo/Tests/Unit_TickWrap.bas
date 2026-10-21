Attribute VB_Name = "Unit_TickWrap"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Tick Wrap Test Suite (plan 21.003, fase 2)
' Describes the CORRECT behavior of the pure tick rules in modTickRules with
' injected ticks: normal, negative (Windows uptime 24.8-49.7 days), across
' &H7FFFFFFF -> &H80000000 and across the real wrap &HFFFFFFFF -> 0.
' Negative values are built with hex literals (&H80000000 = -2147483648).
' An overflow (error 6) inside a rule is caught by each test's handler and
' counts as a failed test, never as a crash of the runner.
' Tests marked [control] pass with the signed arithmetic too: they guard the
' fix against over-correcting. The others are red until phase 4.
' ==========================================================================
Public Function test_suite_tick_wrap() As Boolean
    Call UnitTesting.RunTest("test_cooldown_normal_in_cooldown", test_cooldown_normal_in_cooldown())
    Call UnitTesting.RunTest("test_cooldown_normal_expired", test_cooldown_normal_expired())
    Call UnitTesting.RunTest("test_cooldown_zero_duration_never_active", test_cooldown_zero_duration_never_active())
    Call UnitTesting.RunTest("test_cooldown_last_zero_means_never", test_cooldown_last_zero_means_never())
    Call UnitTesting.RunTest("test_cooldown_both_negative_in_cooldown", test_cooldown_both_negative_in_cooldown())
    Call UnitTesting.RunTest("test_cooldown_both_negative_expired", test_cooldown_both_negative_expired())
    Call UnitTesting.RunTest("test_cooldown_exact_border_normal", test_cooldown_exact_border_normal())
    Call UnitTesting.RunTest("test_cooldown_exact_border_negative_ticks", test_cooldown_exact_border_negative_ticks())
    Call UnitTesting.RunTest("test_cooldown_cross_2_31", test_cooldown_cross_2_31())
    Call UnitTesting.RunTest("test_cooldown_true_wrap", test_cooldown_true_wrap())
    Call UnitTesting.RunTest("test_stamp_zero_is_not_set", test_stamp_zero_is_not_set())
    Call UnitTesting.RunTest("test_stamp_positive_is_set", test_stamp_positive_is_set())
    Call UnitTesting.RunTest("test_stamp_negative_is_set", test_stamp_negative_is_set())
    Call UnitTesting.RunTest("test_stamp_minus_one_is_set", test_stamp_minus_one_is_set())
    Call UnitTesting.RunTest("test_poison_expiry_normal", test_poison_expiry_normal())
    Call UnitTesting.RunTest("test_poison_expiry_cross_2_31_not_expired", test_poison_expiry_cross_2_31_not_expired())
    Call UnitTesting.RunTest("test_poison_expiry_cross_2_31_expired", test_poison_expiry_cross_2_31_expired())
    Call UnitTesting.RunTest("test_poison_expiry_exact_border_cross_2_31", test_poison_expiry_exact_border_cross_2_31())
    Call UnitTesting.RunTest("test_poison_expiry_both_negative", test_poison_expiry_both_negative())
    Call UnitTesting.RunTest("test_poison_expiry_true_wrap", test_poison_expiry_true_wrap())
    Call UnitTesting.RunTest("test_poison_expiry_inverted_args_guard", test_poison_expiry_inverted_args_guard())
    Call UnitTesting.RunTest("test_ticks_elapsed_inverted_args_not_silent", test_ticks_elapsed_inverted_args_not_silent())
    Call UnitTesting.RunTest("test_radar_due_normal", test_radar_due_normal())
    Call UnitTesting.RunTest("test_radar_due_last_zero_now_negative", test_radar_due_last_zero_now_negative())
    Call UnitTesting.RunTest("test_radar_due_cross_2_31", test_radar_due_cross_2_31())
    Call UnitTesting.RunTest("test_radar_due_both_negative", test_radar_due_both_negative())
    Call UnitTesting.RunTest("test_radar_due_true_wrap", test_radar_due_true_wrap())
    Call UnitTesting.RunTest("test_session_elapsed_normal", test_session_elapsed_normal())
    Call UnitTesting.RunTest("test_session_elapsed_cross_2_31", test_session_elapsed_cross_2_31())
    Call UnitTesting.RunTest("test_session_elapsed_both_negative", test_session_elapsed_both_negative())
    Call UnitTesting.RunTest("test_session_elapsed_true_wrap", test_session_elapsed_true_wrap())
    Call UnitTesting.RunTest("test_session_elapsed_huge_clamped", test_session_elapsed_huge_clamped())
    Call UnitTesting.RunTest("test_fight_duration_normal", test_fight_duration_normal())
    Call UnitTesting.RunTest("test_fight_duration_cross_2_31", test_fight_duration_cross_2_31())
    Call UnitTesting.RunTest("test_fight_duration_true_wrap", test_fight_duration_true_wrap())
    Call UnitTesting.RunTest("test_fight_duration_huge_clamped", test_fight_duration_huge_clamped())
    Call UnitTesting.RunTest("test_tick_older_normal", test_tick_older_normal())
    Call UnitTesting.RunTest("test_tick_older_cross_2_31", test_tick_older_cross_2_31())
    Call UnitTesting.RunTest("test_tick_older_true_wrap", test_tick_older_true_wrap())
    Call UnitTesting.RunTest("test_csv_positive_ticks", test_csv_positive_ticks())
    Call UnitTesting.RunTest("test_csv_tick_past_2_31", test_csv_tick_past_2_31())
    Call UnitTesting.RunTest("test_csv_tick_minus_one", test_csv_tick_minus_one())
    Call UnitTesting.RunTest("test_nonzero_tick_from_zero_is_one", test_nonzero_tick_from_zero_is_one())
    Call UnitTesting.RunTest("test_nonzero_tick_from_keeps_positive", test_nonzero_tick_from_keeps_positive())
    Call UnitTesting.RunTest("test_nonzero_tick_from_keeps_negative", test_nonzero_tick_from_keeps_negative())
    Call UnitTesting.RunTest("test_nonzero_tick_from_keeps_2_31", test_nonzero_tick_from_keeps_2_31())
    Call UnitTesting.RunTest("test_sentinel_zero_tick_is_in_cooldown", test_sentinel_zero_tick_is_in_cooldown())
    test_suite_tick_wrap = True
End Function

' [control] Normal ticks: used 2 s ago with a 5 s cooldown is active.
Private Function test_cooldown_normal_in_cooldown() As Boolean
    On Error GoTo Err_Handler
    test_cooldown_normal_in_cooldown = True
    If Not (CooldownActive(10000, 8000, 5000)) Then test_cooldown_normal_in_cooldown = False: Exit Function
    If Not (CooldownActive(12999, 8000, 5000)) Then test_cooldown_normal_in_cooldown = False: Exit Function
    Exit Function
Err_Handler:
    test_cooldown_normal_in_cooldown = False
End Function

' [control] Normal ticks: used 12 s ago with a 5 s cooldown is not active.
Private Function test_cooldown_normal_expired() As Boolean
    On Error GoTo Err_Handler
    test_cooldown_normal_expired = True
    If Not (Not CooldownActive(20000, 8000, 5000)) Then test_cooldown_normal_expired = False: Exit Function
    Exit Function
Err_Handler:
    test_cooldown_normal_expired = False
End Function

' [control] A zero cooldown is never active, with normal or negative ticks.
Private Function test_cooldown_zero_duration_never_active() As Boolean
    On Error GoTo Err_Handler
    test_cooldown_zero_duration_never_active = True
    If Not (Not CooldownActive(10000, 8000, 0)) Then test_cooldown_zero_duration_never_active = False: Exit Function
    If Not (Not CooldownActive(&H80002000, &H80001000, 0)) Then test_cooldown_zero_duration_never_active = False: Exit Function
    Exit Function
Err_Handler:
    test_cooldown_zero_duration_never_active = False
End Function

' [control] last = 0 means "never used": no cooldown, whatever now is.
Private Function test_cooldown_last_zero_means_never() As Boolean
    On Error GoTo Err_Handler
    test_cooldown_last_zero_means_never = True
    If Not (Not CooldownActive(10000, 0, 5000)) Then test_cooldown_last_zero_means_never = False: Exit Function
    If Not (Not CooldownActive(&H80000100, 0, 5000)) Then test_cooldown_last_zero_means_never = False: Exit Function
    If Not (Not CooldownActive(&H100, 0, 5000)) Then test_cooldown_last_zero_means_never = False: Exit Function
    Exit Function
Err_Handler:
    test_cooldown_last_zero_means_never = False
End Function

' Both ticks negative (Windows uptime 24.8-49.7 days): a negative stamp is a set stamp.
Private Function test_cooldown_both_negative_in_cooldown() As Boolean
    On Error GoTo Err_Handler
    test_cooldown_both_negative_in_cooldown = True
    If Not (CooldownActive(&H80002000, &H80001000, 5000)) Then test_cooldown_both_negative_in_cooldown = False: Exit Function
    Exit Function
Err_Handler:
    test_cooldown_both_negative_in_cooldown = False
End Function

' [control] Both ticks negative and 30 s apart: expired.
Private Function test_cooldown_both_negative_expired() As Boolean
    On Error GoTo Err_Handler
    test_cooldown_both_negative_expired = True
    If Not (Not CooldownActive(&H80008530, &H80001000, 5000)) Then test_cooldown_both_negative_expired = False: Exit Function
    Exit Function
Err_Handler:
    test_cooldown_both_negative_expired = False
End Function

' [control] Exact border with normal ticks: elapsed = cooldown is expired (strict <).
Private Function test_cooldown_exact_border_normal() As Boolean
    On Error GoTo Err_Handler
    test_cooldown_exact_border_normal = True
    If Not (Not CooldownActive(15000, 10000, 5000)) Then test_cooldown_exact_border_normal = False: Exit Function
    If Not (CooldownActive(14999, 10000, 5000)) Then test_cooldown_exact_border_normal = False: Exit Function
    Exit Function
Err_Handler:
    test_cooldown_exact_border_normal = False
End Function

' Exact border with negative ticks: 4999 ms active, 5000 ms expired.
Private Function test_cooldown_exact_border_negative_ticks() As Boolean
    On Error GoTo Err_Handler
    test_cooldown_exact_border_negative_ticks = True
    If Not (Not CooldownActive(&H80002388, &H80001000, 5000)) Then test_cooldown_exact_border_negative_ticks = False: Exit Function
    If Not (CooldownActive(&H80002387, &H80001000, 5000)) Then test_cooldown_exact_border_negative_ticks = False: Exit Function
    Exit Function
Err_Handler:
    test_cooldown_exact_border_negative_ticks = False
End Function

' last just before &H7FFFFFFF, now just after: 512 ms apart is active, 6000 ms is expired.
Private Function test_cooldown_cross_2_31() As Boolean
    On Error GoTo Err_Handler
    test_cooldown_cross_2_31 = True
    If Not (CooldownActive(&H80000100, &H7FFFFF00, 5000)) Then test_cooldown_cross_2_31 = False: Exit Function
    If Not (Not CooldownActive(&H80001670, &H7FFFFF00, 5000)) Then test_cooldown_cross_2_31 = False: Exit Function
    Exit Function
Err_Handler:
    test_cooldown_cross_2_31 = False
End Function

' Real wrap &HFFFFFFFF -> 0: 512 ms apart is active, 8448 ms is expired.
Private Function test_cooldown_true_wrap() As Boolean
    On Error GoTo Err_Handler
    test_cooldown_true_wrap = True
    If Not (CooldownActive(&H100, &HFFFFFF00, 5000)) Then test_cooldown_true_wrap = False: Exit Function
    If Not (Not CooldownActive(&H2000, &HFFFFFF00, 5000)) Then test_cooldown_true_wrap = False: Exit Function
    Exit Function
Err_Handler:
    test_cooldown_true_wrap = False
End Function

' [control] Stamp 0 means "never set".
Private Function test_stamp_zero_is_not_set() As Boolean
    On Error GoTo Err_Handler
    test_stamp_zero_is_not_set = True
    If Not (Not TickStampIsSet(0)) Then test_stamp_zero_is_not_set = False: Exit Function
    Exit Function
Err_Handler:
    test_stamp_zero_is_not_set = False
End Function

' [control] Positive stamps are set.
Private Function test_stamp_positive_is_set() As Boolean
    On Error GoTo Err_Handler
    test_stamp_positive_is_set = True
    If Not (TickStampIsSet(1)) Then test_stamp_positive_is_set = False: Exit Function
    If Not (TickStampIsSet(&H7FFFFFFF)) Then test_stamp_positive_is_set = False: Exit Function
    Exit Function
Err_Handler:
    test_stamp_positive_is_set = False
End Function

' A negative stamp (uptime past 2^31 ms) must count as set.
Private Function test_stamp_negative_is_set() As Boolean
    On Error GoTo Err_Handler
    test_stamp_negative_is_set = True
    If Not (TickStampIsSet(&H80000000)) Then test_stamp_negative_is_set = False: Exit Function
    If Not (TickStampIsSet(&H80000001)) Then test_stamp_negative_is_set = False: Exit Function
    If Not (TickStampIsSet(&HC0000000)) Then test_stamp_negative_is_set = False: Exit Function
    Exit Function
Err_Handler:
    test_stamp_negative_is_set = False
End Function

' Stamp -1 (&HFFFFFFFF, last ms before the wrap) must count as set.
Private Function test_stamp_minus_one_is_set() As Boolean
    On Error GoTo Err_Handler
    test_stamp_minus_one_is_set = True
    If Not (TickStampIsSet(&HFFFFFFFF)) Then test_stamp_minus_one_is_set = False: Exit Function
    If Not (TickStampIsSet(-1)) Then test_stamp_minus_one_is_set = False: Exit Function
    Exit Function
Err_Handler:
    test_stamp_minus_one_is_set = False
End Function

' [control] Normal ticks: strict > against the max duration.
Private Function test_poison_expiry_normal() As Boolean
    On Error GoTo Err_Handler
    test_poison_expiry_normal = True
    If Not (Not PoisonApplicationExpired(50000, 1000, 60000)) Then test_poison_expiry_normal = False: Exit Function
    If Not (Not PoisonApplicationExpired(61000, 1000, 60000)) Then test_poison_expiry_normal = False: Exit Function
    If Not (PoisonApplicationExpired(61001, 1000, 60000)) Then test_poison_expiry_normal = False: Exit Function
    If Not (PoisonApplicationExpired(62000, 1000, 60000)) Then test_poison_expiry_normal = False: Exit Function
    Exit Function
Err_Handler:
    test_poison_expiry_normal = False
End Function

' Applied before &H7FFFFFFF, used 512 ms later after it: not expired.
Private Function test_poison_expiry_cross_2_31_not_expired() As Boolean
    On Error GoTo Err_Handler
    test_poison_expiry_cross_2_31_not_expired = True
    If Not (Not PoisonApplicationExpired(&H80000100, &H7FFFFF00, 60000)) Then test_poison_expiry_cross_2_31_not_expired = False: Exit Function
    Exit Function
Err_Handler:
    test_poison_expiry_cross_2_31_not_expired = False
End Function

' Applied before &H7FFFFFFF, used 100512 ms later after it: expired.
Private Function test_poison_expiry_cross_2_31_expired() As Boolean
    On Error GoTo Err_Handler
    test_poison_expiry_cross_2_31_expired = True
    If Not (PoisonApplicationExpired(&H800187A0, &H7FFFFF00, 60000)) Then test_poison_expiry_cross_2_31_expired = False: Exit Function
    Exit Function
Err_Handler:
    test_poison_expiry_cross_2_31_expired = False
End Function

' Exact border across 2^31: 60000 ms not expired, 60001 ms expired.
Private Function test_poison_expiry_exact_border_cross_2_31() As Boolean
    On Error GoTo Err_Handler
    test_poison_expiry_exact_border_cross_2_31 = True
    If Not (Not PoisonApplicationExpired(&H8000E960, &H7FFFFF00, 60000)) Then test_poison_expiry_exact_border_cross_2_31 = False: Exit Function
    If Not (PoisonApplicationExpired(&H8000E961, &H7FFFFF00, 60000)) Then test_poison_expiry_exact_border_cross_2_31 = False: Exit Function
    Exit Function
Err_Handler:
    test_poison_expiry_exact_border_cross_2_31 = False
End Function

' [control] Both ticks negative: 30 s not expired, 70 s expired (max 60 s).
Private Function test_poison_expiry_both_negative() As Boolean
    On Error GoTo Err_Handler
    test_poison_expiry_both_negative = True
    If Not (Not PoisonApplicationExpired(&H80008530, &H80001000, 60000)) Then test_poison_expiry_both_negative = False: Exit Function
    If Not (PoisonApplicationExpired(&H80012170, &H80001000, 60000)) Then test_poison_expiry_both_negative = False: Exit Function
    Exit Function
Err_Handler:
    test_poison_expiry_both_negative = False
End Function

' [control] Real wrap: 512 ms not expired, 131328 ms expired.
Private Function test_poison_expiry_true_wrap() As Boolean
    On Error GoTo Err_Handler
    test_poison_expiry_true_wrap = True
    If Not (Not PoisonApplicationExpired(&H100, &HFFFFFF00, 60000)) Then test_poison_expiry_true_wrap = False: Exit Function
    If Not (PoisonApplicationExpired(&H20000, &HFFFFFF00, 60000)) Then test_poison_expiry_true_wrap = False: Exit Function
    Exit Function
Err_Handler:
    test_poison_expiry_true_wrap = False
End Function

' [control] Argument order (now, applied, max): applied 1 s ago with max 60 s is NOT expired.
Private Function test_poison_expiry_inverted_args_guard() As Boolean
    On Error GoTo Err_Handler
    test_poison_expiry_inverted_args_guard = True
    If Not (Not PoisonApplicationExpired(51000, 50000, 60000)) Then test_poison_expiry_inverted_args_guard = False: Exit Function
    If Not (Not PoisonApplicationExpired(&H80005000, &H80004C18, 60000)) Then test_poison_expiry_inverted_args_guard = False: Exit Function
    Exit Function
Err_Handler:
    test_poison_expiry_inverted_args_guard = False
End Function

' [control] TicksElapsed(start, now): swapped arguments give ~4.29e9, not a small number.
Private Function test_ticks_elapsed_inverted_args_not_silent() As Boolean
    On Error GoTo Err_Handler
    test_ticks_elapsed_inverted_args_not_silent = True
    If Not (TicksElapsed(50000, 51000) = 1000) Then test_ticks_elapsed_inverted_args_not_silent = False: Exit Function
    If Not (TicksElapsed(51000, 50000) > 4000000000#) Then test_ticks_elapsed_inverted_args_not_silent = False: Exit Function
    Exit Function
Err_Handler:
    test_ticks_elapsed_inverted_args_not_silent = False
End Function

' [control] Normal ticks: due at >= 300 ms; last = 0 with a big now is due.
Private Function test_radar_due_normal() As Boolean
    On Error GoTo Err_Handler
    test_radar_due_normal = True
    If Not (RadarDue(1300, 1000)) Then test_radar_due_normal = False: Exit Function
    If Not (RadarDue(1301, 1000)) Then test_radar_due_normal = False: Exit Function
    If Not (Not RadarDue(1299, 1000)) Then test_radar_due_normal = False: Exit Function
    If Not (RadarDue(1000000, 0)) Then test_radar_due_normal = False: Exit Function
    Exit Function
Err_Handler:
    test_radar_due_normal = False
End Function

' Freshly logged user (last = 0) with a negative now must be due.
Private Function test_radar_due_last_zero_now_negative() As Boolean
    On Error GoTo Err_Handler
    test_radar_due_last_zero_now_negative = True
    If Not (RadarDue(&H80000100, 0)) Then test_radar_due_last_zero_now_negative = False: Exit Function
    If Not (RadarDue(&HFFFFFF00, 0)) Then test_radar_due_last_zero_now_negative = False: Exit Function
    Exit Function
Err_Handler:
    test_radar_due_last_zero_now_negative = False
End Function

' Across 2^31: 512 ms apart is due, 32 ms apart is not.
Private Function test_radar_due_cross_2_31() As Boolean
    On Error GoTo Err_Handler
    test_radar_due_cross_2_31 = True
    If Not (RadarDue(&H80000100, &H7FFFFF00)) Then test_radar_due_cross_2_31 = False: Exit Function
    If Not (Not RadarDue(&H80000010, &H7FFFFFF0)) Then test_radar_due_cross_2_31 = False: Exit Function
    Exit Function
Err_Handler:
    test_radar_due_cross_2_31 = False
End Function

' [control] Both ticks negative: exactly 300 ms due, 299 ms not.
Private Function test_radar_due_both_negative() As Boolean
    On Error GoTo Err_Handler
    test_radar_due_both_negative = True
    If Not (RadarDue(&H8000112C, &H80001000)) Then test_radar_due_both_negative = False: Exit Function
    If Not (Not RadarDue(&H8000112B, &H80001000)) Then test_radar_due_both_negative = False: Exit Function
    Exit Function
Err_Handler:
    test_radar_due_both_negative = False
End Function

' [control] Real wrap: 512 ms due, 32 ms not.
Private Function test_radar_due_true_wrap() As Boolean
    On Error GoTo Err_Handler
    test_radar_due_true_wrap = True
    If Not (RadarDue(&H100, &HFFFFFF00)) Then test_radar_due_true_wrap = False: Exit Function
    If Not (Not RadarDue(&H10, &HFFFFFFF0)) Then test_radar_due_true_wrap = False: Exit Function
    Exit Function
Err_Handler:
    test_radar_due_true_wrap = False
End Function

' [control] Normal ticks.
Private Function test_session_elapsed_normal() As Boolean
    On Error GoTo Err_Handler
    test_session_elapsed_normal = True
    If Not (SessionElapsedMs(5000, 1000) = 4000) Then test_session_elapsed_normal = False: Exit Function
    Exit Function
Err_Handler:
    test_session_elapsed_normal = False
End Function

' Session crossing &H7FFFFFFF -> &H80000000 lasts 512 ms, not an overflow.
Private Function test_session_elapsed_cross_2_31() As Boolean
    On Error GoTo Err_Handler
    test_session_elapsed_cross_2_31 = True
    If Not (SessionElapsedMs(&H80000100, &H7FFFFF00) = 512) Then test_session_elapsed_cross_2_31 = False: Exit Function
    Exit Function
Err_Handler:
    test_session_elapsed_cross_2_31 = False
End Function

' [control] Both ticks negative.
Private Function test_session_elapsed_both_negative() As Boolean
    On Error GoTo Err_Handler
    test_session_elapsed_both_negative = True
    If Not (SessionElapsedMs(&H80016F90, &H80001000) = 90000) Then test_session_elapsed_both_negative = False: Exit Function
    Exit Function
Err_Handler:
    test_session_elapsed_both_negative = False
End Function

' [control] Real wrap.
Private Function test_session_elapsed_true_wrap() As Boolean
    On Error GoTo Err_Handler
    test_session_elapsed_true_wrap = True
    If Not (SessionElapsedMs(&H100, &HFFFFFF00) = 512) Then test_session_elapsed_true_wrap = False: Exit Function
    Exit Function
Err_Handler:
    test_session_elapsed_true_wrap = False
End Function

' Sessions longer than 2^31-1 ms are clamped to 2147483647 (never negative).
Private Function test_session_elapsed_huge_clamped() As Boolean
    On Error GoTo Err_Handler
    test_session_elapsed_huge_clamped = True
    If Not (SessionElapsedMs(&H7FFFFFFF, 0) = 2147483647) Then test_session_elapsed_huge_clamped = False: Exit Function
    If Not (SessionElapsedMs(&HF0000000, 1000) = 2147483647) Then test_session_elapsed_huge_clamped = False: Exit Function
    If Not (SessionElapsedMs(&H7FFFFFFF, &H80000000) = 2147483647) Then test_session_elapsed_huge_clamped = False: Exit Function
    Exit Function
Err_Handler:
    test_session_elapsed_huge_clamped = False
End Function

' [control] Normal ticks.
Private Function test_fight_duration_normal() As Boolean
    On Error GoTo Err_Handler
    test_fight_duration_normal = True
    If Not (FightDurationMs(40000, 10000) = 30000) Then test_fight_duration_normal = False: Exit Function
    Exit Function
Err_Handler:
    test_fight_duration_normal = False
End Function

' A fight crossing 2^31 lasts 32 ms, not an overflow.
Private Function test_fight_duration_cross_2_31() As Boolean
    On Error GoTo Err_Handler
    test_fight_duration_cross_2_31 = True
    If Not (FightDurationMs(&H80000010, &H7FFFFFF0) = 32) Then test_fight_duration_cross_2_31 = False: Exit Function
    Exit Function
Err_Handler:
    test_fight_duration_cross_2_31 = False
End Function

' [control] Real wrap.
Private Function test_fight_duration_true_wrap() As Boolean
    On Error GoTo Err_Handler
    test_fight_duration_true_wrap = True
    If Not (FightDurationMs(&H10, &HFFFFFFF0) = 32) Then test_fight_duration_true_wrap = False: Exit Function
    Exit Function
Err_Handler:
    test_fight_duration_true_wrap = False
End Function

' Durations longer than 2^31-1 ms are clamped to 2147483647.
Private Function test_fight_duration_huge_clamped() As Boolean
    On Error GoTo Err_Handler
    test_fight_duration_huge_clamped = True
    If Not (FightDurationMs(&HF0000000, 1000) = 2147483647) Then test_fight_duration_huge_clamped = False: Exit Function
    Exit Function
Err_Handler:
    test_fight_duration_huge_clamped = False
End Function

' [control] Normal ticks: a is older than b only when it was stamped first.
Private Function test_tick_older_normal() As Boolean
    On Error GoTo Err_Handler
    test_tick_older_normal = True
    If Not (TickIsOlder(10000, 1000, 5000)) Then test_tick_older_normal = False: Exit Function
    If Not (Not TickIsOlder(10000, 5000, 1000)) Then test_tick_older_normal = False: Exit Function
    If Not (Not TickIsOlder(10000, 5000, 5000)) Then test_tick_older_normal = False: Exit Function
    Exit Function
Err_Handler:
    test_tick_older_normal = False
End Function

' Across 2^31 the older stamp is the one before &H7FFFFFFF, in both directions.
Private Function test_tick_older_cross_2_31() As Boolean
    On Error GoTo Err_Handler
    test_tick_older_cross_2_31 = True
    If Not (TickIsOlder(&H80001000, &H7FFFFF00, &H80000100)) Then test_tick_older_cross_2_31 = False: Exit Function
    If Not (Not TickIsOlder(&H80001000, &H80000100, &H7FFFFF00)) Then test_tick_older_cross_2_31 = False: Exit Function
    Exit Function
Err_Handler:
    test_tick_older_cross_2_31 = False
End Function

' [control] Across the real wrap the older stamp is the one before &HFFFFFFFF, in both directions.
Private Function test_tick_older_true_wrap() As Boolean
    On Error GoTo Err_Handler
    test_tick_older_true_wrap = True
    If Not (TickIsOlder(&H1000, &HFFFFFF00, &H100)) Then test_tick_older_true_wrap = False: Exit Function
    If Not (Not TickIsOlder(&H1000, &H100, &HFFFFFF00)) Then test_tick_older_true_wrap = False: Exit Function
    Exit Function
Err_Handler:
    test_tick_older_true_wrap = False
End Function

' [control] Positive ticks print as is.
Private Function test_csv_positive_ticks() As Boolean
    On Error GoTo Err_Handler
    test_csv_positive_ticks = True
    If Not (TickToCsv(12345) = "12345") Then test_csv_positive_ticks = False: Exit Function
    If Not (TickToCsv(0) = "0") Then test_csv_positive_ticks = False: Exit Function
    If Not (TickToCsv(&H7FFFFFFF) = "2147483647") Then test_csv_positive_ticks = False: Exit Function
    Exit Function
Err_Handler:
    test_csv_positive_ticks = False
End Function

' ts_ms is unsigned: &H80000000 prints 2147483648.
Private Function test_csv_tick_past_2_31() As Boolean
    On Error GoTo Err_Handler
    test_csv_tick_past_2_31 = True
    If Not (TickToCsv(&H80000000) = "2147483648") Then test_csv_tick_past_2_31 = False: Exit Function
    If Not (TickToCsv(&H80000001) = "2147483649") Then test_csv_tick_past_2_31 = False: Exit Function
    Exit Function
Err_Handler:
    test_csv_tick_past_2_31 = False
End Function

' ts_ms is unsigned: -1 prints 4294967295.
Private Function test_csv_tick_minus_one() As Boolean
    On Error GoTo Err_Handler
    test_csv_tick_minus_one = True
    If Not (TickToCsv(-1) = "4294967295") Then test_csv_tick_minus_one = False: Exit Function
    If Not (TickToCsv(&HFFFFFF00) = "4294967040") Then test_csv_tick_minus_one = False: Exit Function
    Exit Function
Err_Handler:
    test_csv_tick_minus_one = False
End Function

' NonZeroTickFrom: 0 is reserved as "never", so a stored 0 becomes 1.
Private Function test_nonzero_tick_from_zero_is_one() As Boolean
    On Error GoTo Err_Handler
    test_nonzero_tick_from_zero_is_one = True
    If Not (NonZeroTickFrom(0) = 1) Then test_nonzero_tick_from_zero_is_one = False: Exit Function
    Exit Function
Err_Handler:
    test_nonzero_tick_from_zero_is_one = False
End Function

' NonZeroTickFrom leaves positive ticks alone.
Private Function test_nonzero_tick_from_keeps_positive() As Boolean
    On Error GoTo Err_Handler
    test_nonzero_tick_from_keeps_positive = True
    If Not (NonZeroTickFrom(1) = 1) Then test_nonzero_tick_from_keeps_positive = False: Exit Function
    If Not (NonZeroTickFrom(12345) = 12345) Then test_nonzero_tick_from_keeps_positive = False: Exit Function
    If Not (NonZeroTickFrom(&H7FFFFFFF) = &H7FFFFFFF) Then test_nonzero_tick_from_keeps_positive = False: Exit Function
    Exit Function
Err_Handler:
    test_nonzero_tick_from_keeps_positive = False
End Function

' NonZeroTickFrom leaves negative ticks alone (uptime past 2^31 ms).
Private Function test_nonzero_tick_from_keeps_negative() As Boolean
    On Error GoTo Err_Handler
    test_nonzero_tick_from_keeps_negative = True
    If Not (NonZeroTickFrom(-5) = -5) Then test_nonzero_tick_from_keeps_negative = False: Exit Function
    If Not (NonZeroTickFrom(-1) = -1) Then test_nonzero_tick_from_keeps_negative = False: Exit Function
    Exit Function
Err_Handler:
    test_nonzero_tick_from_keeps_negative = False
End Function

' NonZeroTickFrom leaves &H80000000 alone (the first negative tick).
Private Function test_nonzero_tick_from_keeps_2_31() As Boolean
    On Error GoTo Err_Handler
    test_nonzero_tick_from_keeps_2_31 = True
    If Not (NonZeroTickFrom(&H80000000) = &H80000000) Then test_nonzero_tick_from_keeps_2_31 = False: Exit Function
    Exit Function
Err_Handler:
    test_nonzero_tick_from_keeps_2_31 = False
End Function

' Sentinel integration: a stamp that the clock read as exactly 0 is stored as 1 and still
' counts as "used 99 ms ago" (cooldown active), instead of being treated as "never used".
Private Function test_sentinel_zero_tick_is_in_cooldown() As Boolean
    On Error GoTo Err_Handler
    test_sentinel_zero_tick_is_in_cooldown = True
    If Not (CooldownActive(100, NonZeroTickFrom(0), 5000)) Then test_sentinel_zero_tick_is_in_cooldown = False: Exit Function
    Exit Function
Err_Handler:
    test_sentinel_zero_tick_is_in_cooldown = False
End Function

#End If
