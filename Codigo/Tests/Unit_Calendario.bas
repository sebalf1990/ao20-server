Attribute VB_Name = "Unit_Calendario"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Automatic events calendar (plan 2026-11-24-revision-eventos-automaticos, stage 2).
' Pure functions of ModCalendarioConfig with an INJECTED clock: nothing here reads Now.
' Reference dates: 2026-11-27 is a Friday, 28 a Saturday, 29 a Sunday, 2026-11-23 a Monday.
' ==========================================================================

Public Function test_suite_calendario() As Boolean
    Call UnitTesting.RunTest("test_cal_window_start_and_end", test_cal_window_start_and_end())
    Call UnitTesting.RunTest("test_cal_day_mask", test_cal_day_mask())
    Call UnitTesting.RunTest("test_cal_midnight_window", test_cal_midnight_window())
    Call UnitTesting.RunTest("test_cal_midnight_uses_start_day", test_cal_midnight_uses_start_day())
    Call UnitTesting.RunTest("test_cal_one_fire_per_date", test_cal_one_fire_per_date())
    Call UnitTesting.RunTest("test_cal_parse_dias_hora", test_cal_parse_dias_hora())
    Call UnitTesting.RunTest("test_cal_key_is_stable", test_cal_key_is_stable())
    Call UnitTesting.RunTest("test_cal_validate_ranges_no_byte_overflow", test_cal_validate_ranges_no_byte_overflow())
    Call UnitTesting.RunTest("test_cal_validate_modes", test_cal_validate_modes())
    Call UnitTesting.RunTest("test_cal_next_start", test_cal_next_start())
    test_suite_calendario = True
End Function

Private Function TestFecha(ByVal anio As Long, ByVal mes As Long, ByVal dia As Long, ByVal hh As Long, ByVal mm As Long) As Date
    TestFecha = DateSerial(anio, mes, dia) + TimeSerial(hh, mm, 0)
End Function

Private Function TestVariante() As t_CalVarianteRaw
    Dim v As t_CalVarianteRaw
    v.Modo = e_CalModo.CalModoDeathMatch
    v.MinLevel = 1
    v.MaxLevel = 47
    v.MinPlayers = 2
    v.MaxPlayers = 20
    v.TeamSize = 0
    v.TeamType = 1
    v.RoundNumber = 1
    v.InscriptionFee = 500
    v.MinOnline = 2
    v.Description = vbNullString
    TestVariante = v
End Function

' Friday 21:00, window 30: [21:00, 21:30).
Private Function test_cal_window_start_and_end() As Boolean
    Dim f  As String
    Dim ok As Boolean
    ok = Not CalTurnoEnVentana(CAL_DIA_VIE, 1260, 30, TestFecha(2026, 11, 27, 20, 59), f)
    ok = ok And CalTurnoEnVentana(CAL_DIA_VIE, 1260, 30, TestFecha(2026, 11, 27, 21, 0), f)
    ok = ok And (f = "2026-11-27")
    ok = ok And CalTurnoEnVentana(CAL_DIA_VIE, 1260, 30, TestFecha(2026, 11, 27, 21, 29), f)
    ok = ok And Not CalTurnoEnVentana(CAL_DIA_VIE, 1260, 30, TestFecha(2026, 11, 27, 21, 30), f)
    test_cal_window_start_and_end = ok
End Function

Private Function test_cal_day_mask() As Boolean
    Dim f  As String
    Dim ok As Boolean
    ' Friday-only turno does not fire on Saturday at the same hour.
    ok = Not CalTurnoEnVentana(CAL_DIA_VIE, 1260, 30, TestFecha(2026, 11, 28, 21, 10), f)
    ' TODOS fires on Saturday.
    ok = ok And CalTurnoEnVentana(CAL_DIAS_TODOS, 1260, 30, TestFecha(2026, 11, 28, 21, 10), f)
    ' A list of days.
    ok = ok And CalTurnoEnVentana(CAL_DIA_VIE Or CAL_DIA_SAB, 1260, 30, TestFecha(2026, 11, 28, 21, 10), f)
    ok = ok And Not CalTurnoEnVentana(CAL_DIA_VIE Or CAL_DIA_SAB, 1260, 30, TestFecha(2026, 11, 29, 21, 10), f)
    test_cal_day_mask = ok
End Function

' Friday 23:50, window 30: [23:50, 00:20 of Saturday).
Private Function test_cal_midnight_window() As Boolean
    Dim f  As String
    Dim ok As Boolean
    ok = CalTurnoEnVentana(CAL_DIA_VIE, 1430, 30, TestFecha(2026, 11, 27, 23, 55), f)
    ok = ok And (f = "2026-11-27")
    ok = ok And CalTurnoEnVentana(CAL_DIA_VIE, 1430, 30, TestFecha(2026, 11, 28, 0, 10), f)
    ok = ok And (f = "2026-11-27")
    ok = ok And Not CalTurnoEnVentana(CAL_DIA_VIE, 1430, 30, TestFecha(2026, 11, 28, 0, 20), f)
    test_cal_midnight_window = ok
End Function

' A Saturday-only turno at 23:50: its window after midnight (Sunday 00:05) is still the Saturday occurrence,
' and Friday 23:55 is not in it.
Private Function test_cal_midnight_uses_start_day() As Boolean
    Dim f  As String
    Dim ok As Boolean
    ok = Not CalTurnoEnVentana(CAL_DIA_SAB, 1430, 30, TestFecha(2026, 11, 27, 23, 55), f)
    ok = ok And CalTurnoEnVentana(CAL_DIA_SAB, 1430, 30, TestFecha(2026, 11, 28, 23, 55), f)
    ok = ok And CalTurnoEnVentana(CAL_DIA_SAB, 1430, 30, TestFecha(2026, 11, 29, 0, 5), f)
    ok = ok And (f = "2026-11-28")
    ' The Sunday itself is not a Saturday turno day: 23:55 Sunday is outside.
    ok = ok And Not CalTurnoEnVentana(CAL_DIA_SAB, 1430, 30, TestFecha(2026, 11, 29, 23, 55), f)
    test_cal_midnight_uses_start_day = ok
End Function

Private Function test_cal_one_fire_per_date() As Boolean
    Dim f  As String
    Dim ok As Boolean
    ' Not handled yet: fires inside the window.
    ok = TurnoDebeDisparar(CAL_DIA_VIE, 1260, 30, TestFecha(2026, 11, 27, 21, 5), vbNullString, f)
    ok = ok And (f = "2026-11-27")
    ok = ok And TurnoDebeDisparar(CAL_DIA_VIE, 1260, 30, TestFecha(2026, 11, 27, 21, 5), "2026-11-20", f)
    ' Already handled on that date (a row exists, also after a restart): never again, window or not.
    ok = ok And Not TurnoDebeDisparar(CAL_DIA_VIE, 1260, 30, TestFecha(2026, 11, 27, 21, 5), "2026-11-27", f)
    ok = ok And Not TurnoDebeDisparar(CAL_DIA_VIE, 1260, 30, TestFecha(2026, 11, 27, 21, 29), "2026-11-27", f)
    ' Outside the window it never fires, handled or not.
    ok = ok And Not TurnoDebeDisparar(CAL_DIA_VIE, 1260, 30, TestFecha(2026, 11, 27, 21, 30), vbNullString, f)
    ' The post-midnight tail of a handled occurrence is still handled.
    ok = ok And Not TurnoDebeDisparar(CAL_DIA_VIE, 1430, 30, TestFecha(2026, 11, 28, 0, 5), "2026-11-27", f)
    ok = ok And TurnoDebeDisparar(CAL_DIA_VIE, 1430, 30, TestFecha(2026, 11, 28, 0, 5), "2026-11-20", f)
    test_cal_one_fire_per_date = ok
End Function

Private Function test_cal_parse_dias_hora() As Boolean
    Dim m  As Long
    Dim mn As Long
    Dim ok As Boolean
    ok = CalParseDias("TODOS", m) And (m = CAL_DIAS_TODOS)
    ok = ok And CalParseDias("sab", m) And (m = CAL_DIA_SAB)
    ok = ok And CalParseDias("MIE, VIE", m) And (m = (CAL_DIA_MIE Or CAL_DIA_VIE))
    ok = ok And Not CalParseDias("MIE,XXX", m)
    ok = ok And Not CalParseDias("", m)
    ok = ok And CalParseHora("19:00", mn) And (mn = 1140)
    ok = ok And CalParseHora("0:05", mn) And (mn = 5)
    ok = ok And CalParseHora("23:59", mn) And (mn = 1439)
    ok = ok And Not CalParseHora("24:00", mn)
    ok = ok And Not CalParseHora("19:60", mn)
    ok = ok And Not CalParseHora("19", mn)
    ok = ok And Not CalParseHora("19:5", mn)
    ok = ok And Not CalParseHora("ab:cd", mn)
    test_cal_parse_dias_hora = ok
End Function

' TODOS and the seven days spelled out are the same turno; another minute or another day is another key.
Private Function test_cal_key_is_stable() As Boolean
    Dim todos As Long
    Dim lista As Long
    Dim ok    As Boolean
    Call CalParseDias("TODOS", todos)
    Call CalParseDias("DOM,LUN,MAR,MIE,JUE,VIE,SAB", lista)
    ok = (CalClaveTurno(todos, 1140) = CalClaveTurno(lista, 1140))
    ok = ok And (CalClaveTurno(todos, 1140) <> CalClaveTurno(todos, 1141))
    ok = ok And (CalClaveTurno(CAL_DIA_VIE, 1140) <> CalClaveTurno(CAL_DIA_SAB, 1140))
    test_cal_key_is_stable = ok
End Function

' 300 players would raise error 6 in a Byte cast: it must be rejected by the range check first.
Private Function test_cal_validate_ranges_no_byte_overflow() As Boolean
    Dim v      As t_CalVarianteRaw
    Dim motivo As String
    Dim adv    As String
    Dim ok     As Boolean
    v = TestVariante()
    ok = CalValidarVariante(v, motivo, adv)
    v = TestVariante()
    v.MaxPlayers = 300
    ok = ok And Not CalValidarVariante(v, motivo, adv)
    v = TestVariante()
    v.MinPlayers = 1
    ok = ok And Not CalValidarVariante(v, motivo, adv)
    v = TestVariante()
    v.MinLevel = 30
    v.MaxLevel = 20
    ok = ok And Not CalValidarVariante(v, motivo, adv)
    v = TestVariante()
    v.MaxLevel = 48
    ok = ok And Not CalValidarVariante(v, motivo, adv)
    v = TestVariante()
    v.InscriptionFee = -1
    ok = ok And Not CalValidarVariante(v, motivo, adv)
    v = TestVariante()
    v.InscriptionFee = CAL_MAX_CUOTA + 1
    ok = ok And Not CalValidarVariante(v, motivo, adv)
    v = TestVariante()
    v.Description = "a" & Chr$(172) & "b"
    ok = ok And Not CalValidarVariante(v, motivo, adv)
    test_cal_validate_ranges_no_byte_overflow = ok
End Function

Private Function test_cal_validate_modes() As Boolean
    Dim v      As t_CalVarianteRaw
    Dim motivo As String
    Dim adv    As String
    Dim ok     As Boolean
    ' INVASION is rejected with the "not enabled in this build" reason.
    v = TestVariante()
    v.Modo = e_CalModo.CalModoInvasion
    ok = Not CalValidarVariante(v, motivo, adv)
    ok = ok And (InStr(1, motivo, "not enabled in this build") > 0)
    ' CAPTURA: at most 48 players.
    v = TestVariante()
    v.Modo = e_CalModo.CalModoCaptura
    v.MaxPlayers = 49
    ok = ok And Not CalValidarVariante(v, motivo, adv)
    v.MaxPlayers = 48
    ok = ok And CalValidarVariante(v, motivo, adv)
    ' NAVAL with an odd MaxPlayers is accepted with a warning.
    v = TestVariante()
    v.Modo = e_CalModo.CalModoNaval
    v.MaxPlayers = 9
    ok = ok And CalValidarVariante(v, motivo, adv)
    ok = ok And (LenB(adv) > 0)
    v.MaxPlayers = 10
    ok = ok And CalValidarVariante(v, motivo, adv)
    ok = ok And (LenB(adv) = 0)
    ' Unknown mode text.
    ok = ok And (CalParseModo("DEATHMATCH") = e_CalModo.CalModoDeathMatch)
    ok = ok And (CalParseModo("loquesea") = e_CalModo.CalModoNinguno)
    test_cal_validate_modes = ok
End Function

Private Function test_cal_next_start() As Boolean
    Dim ok As Boolean
    ' From Monday 2026-11-23 10:00 the next Friday 21:00 is 2026-11-27 21:00.
    ok = (Format$(CalProximoInicio(CAL_DIA_VIE, 1260, TestFecha(2026, 11, 23, 10, 0)), "yyyy-mm-dd hh:nn") = "2026-11-27 21:00")
    ' Friday 21:30: today's start is behind, the next one is a week later.
    ok = ok And (Format$(CalProximoInicio(CAL_DIA_VIE, 1260, TestFecha(2026, 11, 27, 21, 30)), "yyyy-mm-dd hh:nn") = "2026-12-04 21:00")
    ' Friday 20:00: today.
    ok = ok And (Format$(CalProximoInicio(CAL_DIA_VIE, 1260, TestFecha(2026, 11, 27, 20, 0)), "yyyy-mm-dd hh:nn") = "2026-11-27 21:00")
    ' Empty mask: none.
    ok = ok And (CalProximoInicio(0, 1260, TestFecha(2026, 11, 27, 20, 0)) = 0)
    test_cal_next_start = ok
End Function

#End If
