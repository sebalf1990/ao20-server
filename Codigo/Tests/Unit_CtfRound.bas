Attribute VB_Name = "Unit_CtfRound"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' CTF round time limit (plan 2026-11-24-revision-eventos-automaticos, B1).
' Pure helpers of ModCaptura: CtfClampRoundSeconds and CtfTimeoutWinner.
' ==========================================================================

Public Function test_suite_ctf_round() As Boolean
    Call UnitTesting.RunTest("test_ctf_clamp_default_when_missing", test_ctf_clamp_default_when_missing())
    Call UnitTesting.RunTest("test_ctf_clamp_range", test_ctf_clamp_range())
    Call UnitTesting.RunTest("test_ctf_clamp_inside_range_unchanged", test_ctf_clamp_inside_range_unchanged())
    Call UnitTesting.RunTest("test_ctf_timeout_winner_by_rounds", test_ctf_timeout_winner_by_rounds())
    Call UnitTesting.RunTest("test_ctf_timeout_tie_is_draw", test_ctf_timeout_tie_is_draw())
    test_suite_ctf_round = True
End Function

Private Function test_ctf_clamp_default_when_missing() As Boolean
    test_ctf_clamp_default_when_missing = (CtfClampRoundSeconds(0) = 600 And CtfClampRoundSeconds(-5) = 600)
End Function

Private Function test_ctf_clamp_range() As Boolean
    test_ctf_clamp_range = (CtfClampRoundSeconds(1) = 60 And CtfClampRoundSeconds(59) = 60 And CtfClampRoundSeconds(3601) = 3600 And CtfClampRoundSeconds(99999) = 3600)
End Function

Private Function test_ctf_clamp_inside_range_unchanged() As Boolean
    test_ctf_clamp_inside_range_unchanged = (CtfClampRoundSeconds(60) = 60 And CtfClampRoundSeconds(600) = 600 And CtfClampRoundSeconds(3600) = 3600)
End Function

Private Function test_ctf_timeout_winner_by_rounds() As Boolean
    test_ctf_timeout_winner_by_rounds = (CtfTimeoutWinner(1, 0) = 1 And CtfTimeoutWinner(0, 1) = 2 And CtfTimeoutWinner(2, 1) = 1)
End Function

Private Function test_ctf_timeout_tie_is_draw() As Boolean
    test_ctf_timeout_tie_is_draw = (CtfTimeoutWinner(0, 0) = 0 And CtfTimeoutWinner(1, 1) = 0)
End Function

#End If
