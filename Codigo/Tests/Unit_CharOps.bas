Attribute VB_Name = "Unit_CharOps"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Char-ops bridge (plan 2026-11-26-borrar-personaje): the PURE parts of ModAccountBridge, i.e. op line parsing
' and ack line formatting. The DB-backed parts (mark/set deleted) are exercised with the probe
' (tools/testclient scenarios), not here: these suites do not write to the database.
' ==========================================================================

Public Function test_suite_charops() As Boolean
    Call UnitTesting.RunTest("test_charops_parse_ok", test_charops_parse_ok())
    Call UnitTesting.RunTest("test_charops_parse_normalizes", test_charops_parse_normalizes())
    Call UnitTesting.RunTest("test_charops_parse_keeps_semicolons_in_email", test_charops_parse_keeps_semicolons_in_email())
    Call UnitTesting.RunTest("test_charops_parse_rejects_malformed", test_charops_parse_rejects_malformed())
    Call UnitTesting.RunTest("test_charops_parse_keeps_opid_when_malformed", test_charops_parse_keeps_opid_when_malformed())
    Call UnitTesting.RunTest("test_charops_ack_ok", test_charops_ack_ok())
    Call UnitTesting.RunTest("test_charops_ack_error", test_charops_ack_error())
    Call UnitTesting.RunTest("test_charops_ack_sanitizes_detail", test_charops_ack_sanitizes_detail())
    test_suite_charops = True
End Function

Private Function test_charops_parse_ok() As Boolean
    Dim opId As String, opType As String, email As String
    Dim charId As Long
    Dim ok As Boolean
    ok = CharOps_ParseLine("12;mark_char_delete;345;player@example.com", opId, opType, charId, email)
    ok = ok And (opId = "12") And (opType = "mark_char_delete") And (charId = 345) And (email = "player@example.com")
    ok = ok And CharOps_ParseLine("13;set_char_deleted;1;a@b.c", opId, opType, charId, email)
    ok = ok And (opType = "set_char_deleted") And (charId = 1)
    test_charops_parse_ok = ok
End Function

Private Function test_charops_parse_normalizes() As Boolean
    Dim opId As String, opType As String, email As String
    Dim charId As Long
    Dim ok As Boolean
    ok = CharOps_ParseLine(" 5 ; MARK_CHAR_DELETE ; 10 ; X@Y.Z ", opId, opType, charId, email)
    ok = ok And (opId = "5") And (opType = "mark_char_delete") And (charId = 10) And (email = "X@Y.Z")
    test_charops_parse_normalizes = ok
End Function

Private Function test_charops_parse_keeps_semicolons_in_email() As Boolean
    Dim opId As String, opType As String, email As String
    Dim charId As Long
    Dim ok As Boolean
    ok = CharOps_ParseLine("1;set_char_deleted;7;a;b@c.d", opId, opType, charId, email)
    ok = ok And (email = "a;b@c.d")
    test_charops_parse_keeps_semicolons_in_email = ok
End Function

Private Function test_charops_parse_rejects_malformed() As Boolean
    Dim opId As String, opType As String, email As String
    Dim charId As Long
    Dim ok As Boolean
    ok = Not CharOps_ParseLine(vbNullString, opId, opType, charId, email)
    ok = ok And Not CharOps_ParseLine("1;mark_char_delete;5", opId, opType, charId, email)
    ok = ok And Not CharOps_ParseLine(";mark_char_delete;5;a@b.c", opId, opType, charId, email)
    ok = ok And Not CharOps_ParseLine("1;mark_char_delete;12a;a@b.c", opId, opType, charId, email)
    ok = ok And Not CharOps_ParseLine("1;mark_char_delete;-5;a@b.c", opId, opType, charId, email)
    ok = ok And Not CharOps_ParseLine("1;mark_char_delete;0;a@b.c", opId, opType, charId, email)
    ok = ok And Not CharOps_ParseLine("1;mark_char_delete;1e3;a@b.c", opId, opType, charId, email)
    ok = ok And Not CharOps_ParseLine("1;mark_char_delete;1234567890;a@b.c", opId, opType, charId, email)
    ok = ok And Not CharOps_ParseLine("1;mark_char_delete;5;", opId, opType, charId, email)
    ok = ok And Not CharOps_ParseLine("1;mark_char_delete;5;   ", opId, opType, charId, email)
    test_charops_parse_rejects_malformed = ok
End Function

Private Function test_charops_parse_keeps_opid_when_malformed() As Boolean
    Dim opId As String, opType As String, email As String
    Dim charId As Long
    Dim ok As Boolean
    ok = Not CharOps_ParseLine("9;set_char_deleted;x;a@b.c", opId, opType, charId, email)
    ok = ok And (opId = "9")
    ok = ok And Not CharOps_ParseLine("77", opId, opType, charId, email)
    ok = ok And (opId = "77")
    test_charops_parse_keeps_opid_when_malformed = ok
End Function

Private Function test_charops_ack_ok() As Boolean
    Dim ok As Boolean
    ok = (CharOps_FormatAck("3", "ok") = "3;ok")
    ok = ok And (CharOps_FormatAck("3", "OK") = "3;ok")
    test_charops_ack_ok = ok
End Function

Private Function test_charops_ack_error() As Boolean
    Dim ok As Boolean
    ok = (CharOps_FormatAck("3", "error;not_owner") = "3;error;not_owner")
    ok = ok And (CharOps_FormatAck("4", "error;unknown_type") = "4;error;unknown_type")
    ' A result that is neither "ok" nor "error;..." is never reported as ok.
    ok = ok And (CharOps_FormatAck("5", "boom") = "5;error;boom")
    test_charops_ack_error = ok
End Function

Private Function test_charops_ack_sanitizes_detail() As Boolean
    Dim ok As Boolean
    ok = (CharOps_FormatAck("3", "error;a;b") = "3;error;a_b")
    ok = ok And (CharOps_FormatAck("3", "error;" & "x" & vbCrLf & "y") = "3;error;xy")
    ok = ok And (CharOps_FormatAck("3", "error;") = "3;error;unknown_error")
    ok = ok And (CharOps_SanitizeDetail("  ") = "unknown_error")
    test_charops_ack_sanitizes_detail = ok
End Function

#End If
