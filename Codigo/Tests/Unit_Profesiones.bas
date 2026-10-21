Attribute VB_Name = "Unit_Profesiones"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Professions Test Suite (plan 21.002, part B2)
' Characterization of the learnable-professions rules (modProfesiones) and of
' the crafting catalyst probability (modCrafteos.CalculateCraftProb).
' Chance paths are checked at their deterministic edges (skill 255 clamps to
' 100%), or with a wide statistical tolerance. Rules that depend on
' profesiones.ini read the SAME file the server loaded, so a config edit does
' not turn the suite red. Everything that mutates global state restores it.
' The wire / end-to-end flow stays with the probe (craft_receta_*, profesiones).
' ==========================================================================

Private Const PROF_TOGGLE As String = "professions_learnable"

Private Type t_ProfFixture
    Prof(1 To 7) As Integer
    Forgot As Byte
    Priv As e_PlayerType
    Toggle As Boolean
End Type

Public Function test_suite_profesiones() As Boolean
    Call UnitTesting.RunTest("test_prof_slots_ocupados_counts_nonzero", test_prof_slots_ocupados_counts_nonzero())
    Call UnitTesting.RunTest("test_prof_tiene_toggle_off_always_true", test_prof_tiene_toggle_off_always_true())
    Call UnitTesting.RunTest("test_prof_tiene_toggle_on_requires_slot", test_prof_tiene_toggle_on_requires_slot())
    Call UnitTesting.RunTest("test_prof_tiene_gm_bypass", test_prof_tiene_gm_bypass())
    Call UnitTesting.RunTest("test_prof_puede_aprender_toggle_off_false", test_prof_puede_aprender_toggle_off_false())
    Call UnitTesting.RunTest("test_prof_puede_aprender_out_of_range_false", test_prof_puede_aprender_out_of_range_false())
    Call UnitTesting.RunTest("test_prof_puede_aprender_new_profession_ok", test_prof_puede_aprender_new_profession_ok())
    Call UnitTesting.RunTest("test_prof_puede_aprender_already_learned_false", test_prof_puede_aprender_already_learned_false())
    Call UnitTesting.RunTest("test_prof_puede_aprender_forgot_limit_follows_config", test_prof_puede_aprender_forgot_limit_follows_config())
    Call UnitTesting.RunTest("test_prof_puede_aprender_slot_limit_follows_config", test_prof_puede_aprender_slot_limit_follows_config())
    Call UnitTesting.RunTest("test_prof_puede_aprender_gm_and_staff_exempt", test_prof_puede_aprender_gm_and_staff_exempt())
    Call UnitTesting.RunTest("test_prof_puede_olvidar_rules", test_prof_puede_olvidar_rules())
    Call UnitTesting.RunTest("test_prof_puede_olvidar_forgot_limit_follows_config", test_prof_puede_olvidar_forgot_limit_follows_config())
    Call UnitTesting.RunTest("test_prof_cantidad_fija_artesania", test_prof_cantidad_fija_artesania())
    Call UnitTesting.RunTest("test_prof_item_returns_zero_for_invalid_index", test_prof_item_returns_zero_for_invalid_index())
    Call UnitTesting.RunTest("test_prof_item_manual_and_potion_return_zero", test_prof_item_manual_and_potion_return_zero())
    Call UnitTesting.RunTest("test_prof_item_profesionid_field_is_primary", test_prof_item_profesionid_field_is_primary())
    Call UnitTesting.RunTest("test_prof_item_profesionid_out_of_range_is_ignored", test_prof_item_profesionid_out_of_range_is_ignored())
    Call UnitTesting.RunTest("test_prof_item_tool_subtype_fallback", test_prof_item_tool_subtype_fallback())
    Call UnitTesting.RunTest("test_prof_item_profesionid_beats_subtype", test_prof_item_profesionid_beats_subtype())
    Call UnitTesting.RunTest("test_prof_exito_non_extractive_always_true", test_prof_exito_non_extractive_always_true())
    Call UnitTesting.RunTest("test_prof_exito_max_skill_always_true", test_prof_exito_max_skill_always_true())
    Call UnitTesting.RunTest("test_prof_exito_zero_skill_rate_matches_config", test_prof_exito_zero_skill_rate_matches_config())
    Call UnitTesting.RunTest("test_prof_cantidad_non_extractive_is_one", test_prof_cantidad_non_extractive_is_one())
    Call UnitTesting.RunTest("test_prof_cantidad_caps_at_config_max", test_prof_cantidad_caps_at_config_max())
    Call UnitTesting.RunTest("test_prof_cantidad_matches_formula_and_is_monotonic", test_prof_cantidad_matches_formula_and_is_monotonic())
    Call UnitTesting.RunTest("test_prof_cantidad_safe_zone_never_exceeds_normal", test_prof_cantidad_safe_zone_never_exceeds_normal())
    Call UnitTesting.RunTest("test_craft_prob_no_catalyst_is_base", test_craft_prob_no_catalyst_is_base())
    Call UnitTesting.RunTest("test_craft_prob_matching_catalyst_scales", test_craft_prob_matching_catalyst_scales())
    Call UnitTesting.RunTest("test_craft_prob_matching_catalyst_clamps", test_craft_prob_matching_catalyst_clamps())
    Call UnitTesting.RunTest("test_craft_prob_other_type_catalyst_ignored", test_craft_prob_other_type_catalyst_ignored())
    test_suite_profesiones = True
End Function

' --- fixtures -----------------------------------------------------------------

' Saves user 1's profession state, then leaves it clean: no professions, no forgets, plain user, toggle untouched.
Private Sub prof_save(ByRef fx As t_ProfFixture)
    Dim i As Long
    For i = 1 To PROF_MAX_SLOTS
        fx.Prof(i) = UserList(1).Professions(i)
        UserList(1).Professions(i) = 0
    Next i
    fx.Forgot = UserList(1).ProfessionForgotCount
    fx.Priv = UserList(1).flags.Privilegios
    fx.Toggle = IsFeatureEnabled(PROF_TOGGLE)
    UserList(1).ProfessionForgotCount = 0
    UserList(1).flags.Privilegios = e_PlayerType.User
End Sub

Private Sub prof_restore(ByRef fx As t_ProfFixture)
    Dim i As Long
    For i = 1 To PROF_MAX_SLOTS
        UserList(1).Professions(i) = fx.Prof(i)
    Next i
    UserList(1).ProfessionForgotCount = fx.Forgot
    UserList(1).flags.Privilegios = fx.Priv
    Call SetFeatureToggle(PROF_TOGGLE, fx.Toggle)
End Sub

' The same profesiones.ini the server loaded at boot (App.Path).
Private Function prof_live_ini() As clsIniManager
    Dim ini As clsIniManager
    Set ini = New clsIniManager
    Call ini.Initialize(App.Path & "\profesiones.ini")
    Set prof_live_ini = ini
End Function

' Same defaulting LoadProfesionesConfig applies to MaxOlvidosPorPersonaje (0 = unlimited).
Private Function prof_max_olvidos() As Long
    prof_max_olvidos = CLng(CByte(val(prof_live_ini().GetValue("General", "MaxOlvidosPorPersonaje"))))
End Function

' Same defaulting LoadProfesionesConfig applies to MaxProfesionesAprendidas (0 -> 2, cap at PROF_MAX_SLOTS).
Private Function prof_max_aprendidas() As Long
    Dim n As Long
    n = CLng(CByte(val(prof_live_ini().GetValue("General", "MaxProfesionesAprendidas"))))
    If n = 0 Then n = 2
    If n > PROF_MAX_SLOTS Then n = PROF_MAX_SLOTS
    prof_max_aprendidas = n
End Function

Private Function prof_section(ByVal profId As Integer) As String
    Select Case profId
        Case e_Skill.Pescar: prof_section = "Pesca"
        Case e_Skill.Talar: prof_section = "Tala"
        Case e_Skill.Mineria: prof_section = "Mineria"
    End Select
End Function

' First map whose Seguro flag equals wantSafe (0 = normal, 1 = safe zone); 0 if none.
Private Function prof_find_map(ByVal wantSafe As Byte) As Long
    Dim m As Long
    For m = 1 To NumMaps
        If MapInfo(m).Seguro = wantSafe Then
            prof_find_map = m
            Exit Function
        End If
    Next m
End Function

' --- SlotsOcupados / TieneProfesionAprendida ------------------------------------------

' SlotsOcupados counts the non-zero slots, wherever they are.
Private Function test_prof_slots_ocupados_counts_nonzero() As Boolean
    On Error GoTo e
    Dim fx As t_ProfFixture, ok As Boolean, i As Long
    Call prof_save(fx)
    ok = (SlotsOcupados(1) = 0)
    UserList(1).Professions(1) = 17
    UserList(1).Professions(4) = 20
    ok = ok And (SlotsOcupados(1) = 2)
    For i = 1 To PROF_MAX_SLOTS
        UserList(1).Professions(i) = 17 + i - 1
    Next i
    ok = ok And (SlotsOcupados(1) = PROF_MAX_SLOTS)
    Call prof_restore(fx)
    test_prof_slots_ocupados_counts_nonzero = ok
    Exit Function
e:
    Call prof_restore(fx)
    test_prof_slots_ocupados_counts_nonzero = False
End Function

' Toggle OFF: everyone "has" every profession (the legacy game, nothing gated).
Private Function test_prof_tiene_toggle_off_always_true() As Boolean
    On Error GoTo e
    Dim fx As t_ProfFixture, ok As Boolean
    Call prof_save(fx)
    Call SetFeatureToggle(PROF_TOGGLE, False)
    ok = TieneProfesionAprendida(1, 17) And TieneProfesionAprendida(1, 23) And TieneProfesionAprendida(1, 99)
    Call prof_restore(fx)
    test_prof_tiene_toggle_off_always_true = ok
    Exit Function
e:
    Call prof_restore(fx)
    test_prof_tiene_toggle_off_always_true = False
End Function

' Toggle ON, regular player: only the professions in the slots.
Private Function test_prof_tiene_toggle_on_requires_slot() As Boolean
    On Error GoTo e
    Dim fx As t_ProfFixture, ok As Boolean
    Call prof_save(fx)
    Call SetFeatureToggle(PROF_TOGGLE, True)
    ok = (Not TieneProfesionAprendida(1, 17))
    UserList(1).Professions(3) = 17
    ok = ok And TieneProfesionAprendida(1, 17) And (Not TieneProfesionAprendida(1, 18))
    Call prof_restore(fx)
    test_prof_tiene_toggle_on_requires_slot = ok
    Exit Function
e:
    Call prof_restore(fx)
    test_prof_tiene_toggle_on_requires_slot = False
End Function

' Toggle ON, staff (any GM privilege) bypasses the slots completely.
Private Function test_prof_tiene_gm_bypass() As Boolean
    On Error GoTo e
    Dim fx As t_ProfFixture, ok As Boolean
    Call prof_save(fx)
    Call SetFeatureToggle(PROF_TOGGLE, True)
    UserList(1).flags.Privilegios = e_PlayerType.Admin
    ok = TieneProfesionAprendida(1, 17) And TieneProfesionAprendida(1, 23)
    UserList(1).flags.Privilegios = e_PlayerType.Consejero
    ok = ok And TieneProfesionAprendida(1, 20)
    Call prof_restore(fx)
    test_prof_tiene_gm_bypass = ok
    Exit Function
e:
    Call prof_restore(fx)
    test_prof_tiene_gm_bypass = False
End Function

' --- PuedeAprenderProfesion ---------------------------------------------------------------

Private Function test_prof_puede_aprender_toggle_off_false() As Boolean
    On Error GoTo e
    Dim fx As t_ProfFixture, ok As Boolean
    Call prof_save(fx)
    Call SetFeatureToggle(PROF_TOGGLE, False)
    ok = (Not PuedeAprenderProfesion(1, 17, True))
    Call prof_restore(fx)
    test_prof_puede_aprender_toggle_off_false = ok
    Exit Function
e:
    Call prof_restore(fx)
    test_prof_puede_aprender_toggle_off_false = False
End Function

' Only ids PROF_MIN_ID..PROF_MAX_ID (17..23) are professions.
Private Function test_prof_puede_aprender_out_of_range_false() As Boolean
    On Error GoTo e
    Dim fx As t_ProfFixture, ok As Boolean
    Call prof_save(fx)
    Call SetFeatureToggle(PROF_TOGGLE, True)
    ok = (Not PuedeAprenderProfesion(1, 0, True)) And (Not PuedeAprenderProfesion(1, 16, True)) _
        And (Not PuedeAprenderProfesion(1, 24, True)) And (Not PuedeAprenderProfesion(1, 255, True))
    Call prof_restore(fx)
    test_prof_puede_aprender_out_of_range_false = ok
    Exit Function
e:
    Call prof_restore(fx)
    test_prof_puede_aprender_out_of_range_false = False
End Function

' A regular player with a free slot can learn every profession id in range.
Private Function test_prof_puede_aprender_new_profession_ok() As Boolean
    On Error GoTo e
    Dim fx As t_ProfFixture, ok As Boolean, p As Long
    Call prof_save(fx)
    Call SetFeatureToggle(PROF_TOGGLE, True)
    ok = True
    For p = PROF_MIN_ID To PROF_MAX_ID
        ok = ok And PuedeAprenderProfesion(1, p, True)
    Next p
    Call prof_restore(fx)
    test_prof_puede_aprender_new_profession_ok = ok
    Exit Function
e:
    Call prof_restore(fx)
    test_prof_puede_aprender_new_profession_ok = False
End Function

' Already learned: False, for staff actions (no message) and for the player-facing path.
Private Function test_prof_puede_aprender_already_learned_false() As Boolean
    On Error GoTo e
    Dim fx As t_ProfFixture, ok As Boolean
    Call prof_save(fx)
    Call SetFeatureToggle(PROF_TOGGLE, True)
    UserList(1).Professions(2) = 19
    ok = (Not PuedeAprenderProfesion(1, 19, True)) And (Not PuedeAprenderProfesion(1, 19, False))
    ok = ok And PuedeAprenderProfesion(1, 20, True)
    Call prof_restore(fx)
    test_prof_puede_aprender_already_learned_false = ok
    Exit Function
e:
    Call prof_restore(fx)
    test_prof_puede_aprender_already_learned_false = False
End Function

' MaxOlvidosPorPersonaje: 0 = unlimited; N > 0 blocks learning once the forget counter reaches N.
' The expectation is derived from the server's own profesiones.ini.
Private Function test_prof_puede_aprender_forgot_limit_follows_config() As Boolean
    On Error GoTo e
    Dim fx As t_ProfFixture, ok As Boolean, maxOlv As Long
    Call prof_save(fx)
    Call SetFeatureToggle(PROF_TOGGLE, True)
    maxOlv = prof_max_olvidos()
    UserList(1).ProfessionForgotCount = 255
    ok = (PuedeAprenderProfesion(1, 17, False) = (maxOlv = 0))
    If maxOlv > 0 Then
        UserList(1).ProfessionForgotCount = CByte(maxOlv - 1)
        ok = ok And PuedeAprenderProfesion(1, 17, False)
        UserList(1).ProfessionForgotCount = CByte(maxOlv)
        ok = ok And (Not PuedeAprenderProfesion(1, 17, False))
    End If
    Call prof_restore(fx)
    test_prof_puede_aprender_forgot_limit_follows_config = ok
    Exit Function
e:
    Call prof_restore(fx)
    test_prof_puede_aprender_forgot_limit_follows_config = False
End Function

' MaxProfesionesAprendidas: a regular player below the limit may learn, at the limit may not.
' With the shipped value (7) every slot is learnable, so only the "below the limit" side exists.
Private Function test_prof_puede_aprender_slot_limit_follows_config() As Boolean
    On Error GoTo e
    Dim fx As t_ProfFixture, ok As Boolean, n As Long, i As Long, filled As Long
    Call prof_save(fx)
    Call SetFeatureToggle(PROF_TOGGLE, True)
    n = prof_max_aprendidas()
    If n >= PROF_MAX_SLOTS Then
        filled = n - 1
    Else
        filled = n
    End If
    For i = 1 To filled
        UserList(1).Professions(i) = PROF_MIN_ID + i - 1
    Next i
    ' PROF_MAX_ID is never among the filled ids (filled <= 6)
    ok = (PuedeAprenderProfesion(1, PROF_MAX_ID, True) = (n >= PROF_MAX_SLOTS))
    Call prof_restore(fx)
    test_prof_puede_aprender_slot_limit_follows_config = ok
    Exit Function
e:
    Call prof_restore(fx)
    test_prof_puede_aprender_slot_limit_follows_config = False
End Function

' GMs and staff actions skip the config limits (forgot counter at 255, still allowed).
Private Function test_prof_puede_aprender_gm_and_staff_exempt() As Boolean
    On Error GoTo e
    Dim fx As t_ProfFixture, ok As Boolean
    Call prof_save(fx)
    Call SetFeatureToggle(PROF_TOGGLE, True)
    UserList(1).ProfessionForgotCount = 255
    ok = PuedeAprenderProfesion(1, 17, True)
    UserList(1).flags.Privilegios = e_PlayerType.Admin
    ok = ok And PuedeAprenderProfesion(1, 17, False)
    Call prof_restore(fx)
    test_prof_puede_aprender_gm_and_staff_exempt = ok
    Exit Function
e:
    Call prof_restore(fx)
    test_prof_puede_aprender_gm_and_staff_exempt = False
End Function

' --- PuedeOlvidarProfesion -------------------------------------------------------------------

' Toggle OFF -> False; out of range -> False; not learned -> False; learned -> True.
Private Function test_prof_puede_olvidar_rules() As Boolean
    On Error GoTo e
    Dim fx As t_ProfFixture, ok As Boolean
    Call prof_save(fx)
    Call SetFeatureToggle(PROF_TOGGLE, False)
    UserList(1).Professions(1) = 21
    ok = (Not PuedeOlvidarProfesion(1, 21, True))
    Call SetFeatureToggle(PROF_TOGGLE, True)
    ok = ok And PuedeOlvidarProfesion(1, 21, True)
    ok = ok And (Not PuedeOlvidarProfesion(1, 22, True))
    ok = ok And (Not PuedeOlvidarProfesion(1, 0, True)) And (Not PuedeOlvidarProfesion(1, 24, True))
    Call prof_restore(fx)
    test_prof_puede_olvidar_rules = ok
    Exit Function
e:
    Call prof_restore(fx)
    test_prof_puede_olvidar_rules = False
End Function

' Same forget limit as learning: regular player at the limit cannot forget; GM and staff actions can.
Private Function test_prof_puede_olvidar_forgot_limit_follows_config() As Boolean
    On Error GoTo e
    Dim fx As t_ProfFixture, ok As Boolean, maxOlv As Long
    Call prof_save(fx)
    Call SetFeatureToggle(PROF_TOGGLE, True)
    maxOlv = prof_max_olvidos()
    UserList(1).Professions(1) = 21
    UserList(1).ProfessionForgotCount = 255
    ok = (PuedeOlvidarProfesion(1, 21, False) = (maxOlv = 0))
    ok = ok And PuedeOlvidarProfesion(1, 21, True)
    UserList(1).flags.Privilegios = e_PlayerType.Admin
    ok = ok And PuedeOlvidarProfesion(1, 21, False)
    Call prof_restore(fx)
    test_prof_puede_olvidar_forgot_limit_follows_config = ok
    Exit Function
e:
    Call prof_restore(fx)
    test_prof_puede_olvidar_forgot_limit_follows_config = False
End Function

' --- CantidadFijaArtesania -----------------------------------------------------------------------

' Out-of-range ids -> 1; every real id is at least 1; extraction professions never set a fixed amount -> 1.
Private Function test_prof_cantidad_fija_artesania() As Boolean
    On Error GoTo e
    Dim ok As Boolean, p As Long
    ok = (CantidadFijaArtesania(0) = 1) And (CantidadFijaArtesania(16) = 1) And (CantidadFijaArtesania(24) = 1) And (CantidadFijaArtesania(-3) = 1)
    For p = PROF_MIN_ID To PROF_MAX_ID
        ok = ok And (CantidadFijaArtesania(p) >= 1)
    Next p
    ok = ok And (CantidadFijaArtesania(e_Skill.Pescar) = 1) And (CantidadFijaArtesania(e_Skill.Talar) = 1) And (CantidadFijaArtesania(e_Skill.Mineria) = 1)
    test_prof_cantidad_fija_artesania = ok
    Exit Function
e:
    test_prof_cantidad_fija_artesania = False
End Function

' --- ProfesionDelItem (mutates ObjData(1), restored) ---------------------------------------------

Private Function prof_item(ByVal objType As Long, ByVal profId As Integer, ByVal subtipo As Long) As Integer
    ObjData(1).OBJType = objType
    ObjData(1).ProfesionId = profId
    ObjData(1).Subtipo = subtipo
    prof_item = ProfesionDelItem(1)
End Function

Private Function test_prof_item_returns_zero_for_invalid_index() As Boolean
    On Error GoTo e
    test_prof_item_returns_zero_for_invalid_index = (ProfesionDelItem(0) = 0) And (ProfesionDelItem(-1) = 0)
    Exit Function
e:
    test_prof_item_returns_zero_for_invalid_index = False
End Function

' Manuals (otUseOnce) and forget potions (otPotions) reuse ProfesionId as "teaches/forgets": ProfesionDelItem says 0.
Private Function test_prof_item_manual_and_potion_return_zero() As Boolean
    On Error GoTo e
    Dim saved As t_ObjData, ok As Boolean
    saved = ObjData(1)
    ok = (prof_item(e_OBJType.otUseOnce, 21, 0) = 0) And (prof_item(e_OBJType.otPotions, 21, 0) = 0)
    ObjData(1) = saved
    test_prof_item_manual_and_potion_return_zero = ok
    Exit Function
e:
    ObjData(1) = saved
    test_prof_item_manual_and_potion_return_zero = False
End Function

' The obj.dat ProfesionId field is the primary source, for any item type (armor, recipe scroll, etc.).
Private Function test_prof_item_profesionid_field_is_primary() As Boolean
    On Error GoTo e
    Dim saved As t_ObjData, ok As Boolean
    saved = ObjData(1)
    ok = (prof_item(e_OBJType.otArmor, 21, 0) = 21) And (prof_item(e_OBJType.otParchment, 17, 0) = 17) _
        And (prof_item(e_OBJType.otWeapon, 23, 0) = 23)
    ObjData(1) = saved
    test_prof_item_profesionid_field_is_primary = ok
    Exit Function
e:
    ObjData(1) = saved
    test_prof_item_profesionid_field_is_primary = False
End Function

' A ProfesionId outside 17..23 is ignored; a non-tool item then has no profession.
Private Function test_prof_item_profesionid_out_of_range_is_ignored() As Boolean
    On Error GoTo e
    Dim saved As t_ObjData, ok As Boolean
    saved = ObjData(1)
    ok = (prof_item(e_OBJType.otArmor, 0, 0) = 0) And (prof_item(e_OBJType.otArmor, 16, 0) = 0) And (prof_item(e_OBJType.otArmor, 24, 0) = 0)
    ObjData(1) = saved
    test_prof_item_profesionid_out_of_range_is_ignored = ok
    Exit Function
e:
    ObjData(1) = saved
    test_prof_item_profesionid_out_of_range_is_ignored = False
End Function

' Legacy fallback for working tools without ProfesionId: Subtipo -> skill.
Private Function test_prof_item_tool_subtype_fallback() As Boolean
    On Error GoTo e
    Dim saved As t_ObjData, ok As Boolean, t As Long
    saved = ObjData(1)
    t = e_OBJType.otWorkingTools
    ok = (prof_item(t, 0, e_WorkingToolSubType.FishingRod) = e_Skill.Pescar) _
        And (prof_item(t, 0, e_WorkingToolSubType.FishingNet) = e_Skill.Pescar) _
        And (prof_item(t, 0, e_WorkingToolSubType.FishingLine) = e_Skill.Pescar) _
        And (prof_item(t, 0, e_WorkingToolSubType.AlchemyScissors) = e_Skill.Alquimia) _
        And (prof_item(t, 0, e_WorkingToolSubType.AlchemyCauldron) = e_Skill.Alquimia) _
        And (prof_item(t, 0, e_WorkingToolSubType.CarpentryHacksaw) = e_Skill.Carpinteria) _
        And (prof_item(t, 0, e_WorkingToolSubType.FellingAxe) = e_Skill.Talar) _
        And (prof_item(t, 0, e_WorkingToolSubType.SmithHammer) = e_Skill.Herreria) _
        And (prof_item(t, 0, e_WorkingToolSubType.MinerPickaxe) = e_Skill.Mineria) _
        And (prof_item(t, 0, e_WorkingToolSubType.TailorSewingbox) = e_Skill.Sastreria) _
        And (prof_item(t, 0, 0) = 0) And (prof_item(t, 0, 99) = 0)
    ObjData(1) = saved
    test_prof_item_tool_subtype_fallback = ok
    Exit Function
e:
    ObjData(1) = saved
    test_prof_item_tool_subtype_fallback = False
End Function

' When both exist the ProfesionId field wins over the Subtipo mapping (a fishing rod tagged for tailoring is tailoring).
Private Function test_prof_item_profesionid_beats_subtype() As Boolean
    On Error GoTo e
    Dim saved As t_ObjData, ok As Boolean
    saved = ObjData(1)
    ok = (prof_item(e_OBJType.otWorkingTools, 23, e_WorkingToolSubType.FishingRod) = 23)
    ObjData(1) = saved
    test_prof_item_profesionid_beats_subtype = ok
    Exit Function
e:
    ObjData(1) = saved
    test_prof_item_profesionid_beats_subtype = False
End Function

' --- extraction success chance ------------------------------------------------------------------------

' Non-extraction professions and ids outside 17..23 always succeed (the roll does not apply).
Private Function test_prof_exito_non_extractive_always_true() As Boolean
    On Error GoTo e
    Dim ok As Boolean, i As Long
    ok = True
    For i = 1 To 20
        ok = ok And CalcularExitoExtraccion(1, e_Skill.Carpinteria) And CalcularExitoExtraccion(1, e_Skill.Herreria) _
            And CalcularExitoExtraccion(1, e_Skill.Sastreria) And CalcularExitoExtraccion(1, e_Skill.Alquimia) _
            And CalcularExitoExtraccion(1, 16) And CalcularExitoExtraccion(1, 24) And CalcularExitoExtraccion(1, 0)
    Next i
    test_prof_exito_non_extractive_always_true = ok
    Exit Function
e:
    test_prof_exito_non_extractive_always_true = False
End Function

' Skill 255 pushes ProbBase + ProbPorSkill*skill over 100, which is clamped: always succeeds.
Private Function test_prof_exito_max_skill_always_true() As Boolean
    On Error GoTo e
    Dim sp As Byte, ok As Boolean, i As Long, p As Long
    ok = True
    For p = e_Skill.Talar To e_Skill.Mineria
        If p = e_Skill.Talar Or p = e_Skill.Pescar Or p = e_Skill.Mineria Then
            sp = UserList(1).Stats.UserSkills(p)
            UserList(1).Stats.UserSkills(p) = 255
            For i = 1 To 200
                ok = ok And CalcularExitoExtraccion(1, p)
            Next i
            UserList(1).Stats.UserSkills(p) = sp
        End If
    Next p
    test_prof_exito_max_skill_always_true = ok
    Exit Function
e:
    UserList(1).Stats.UserSkills(e_Skill.Pescar) = sp
    test_prof_exito_max_skill_always_true = False
End Function

' At skill 0 the success rate is ProbBase% (from profesiones.ini, [Pesca]); measured over 3000 rolls, +-5 points.
Private Function test_prof_exito_zero_skill_rate_matches_config() As Boolean
    On Error GoTo e
    Dim sp As Byte, hits As Long, i As Long, expectedPct As Long, probBase As Single, probSkill As Single
    probBase = CSng(val(prof_live_ini().GetValue("Pesca", "ProbBase")))
    expectedPct = CLng(probBase)
    If expectedPct > 100 Then expectedPct = 100
    If expectedPct < 0 Then expectedPct = 0
    sp = UserList(1).Stats.UserSkills(e_Skill.Pescar)
    UserList(1).Stats.UserSkills(e_Skill.Pescar) = 0
    For i = 1 To 3000
        If CalcularExitoExtraccion(1, e_Skill.Pescar) Then hits = hits + 1
    Next i
    UserList(1).Stats.UserSkills(e_Skill.Pescar) = sp
    test_prof_exito_zero_skill_rate_matches_config = (Abs((hits / 30#) - expectedPct) <= 5)
    Exit Function
e:
    UserList(1).Stats.UserSkills(e_Skill.Pescar) = sp
    test_prof_exito_zero_skill_rate_matches_config = False
End Function

' --- extraction amount ---------------------------------------------------------------------------------

' Non-extraction professions and ids outside 17..23 give 1.
Private Function test_prof_cantidad_non_extractive_is_one() As Boolean
    On Error GoTo e
    test_prof_cantidad_non_extractive_is_one = (CalcularCantidadExtraccion(1, e_Skill.Carpinteria) = 1) _
        And (CalcularCantidadExtraccion(1, e_Skill.Alquimia) = 1) And (CalcularCantidadExtraccion(1, 16) = 1) _
        And (CalcularCantidadExtraccion(1, 24) = 1) And (CalcularCantidadExtraccion(1, 0) = 1)
    Exit Function
e:
    test_prof_cantidad_non_extractive_is_one = False
End Function

' Skill is capped at SkillParaMax: with skill 255 outside a safe zone the amount is exactly CantidadMax.
Private Function test_prof_cantidad_caps_at_config_max() As Boolean
    On Error GoTo e
    Dim sp As Byte, sm As Integer, ok As Boolean, p As Long, normalMap As Long, cMax As Long
    normalMap = prof_find_map(0)
    If normalMap = 0 Then Exit Function
    sm = UserList(1).pos.Map
    UserList(1).pos.Map = normalMap
    ok = True
    For p = e_Skill.Talar To e_Skill.Mineria
        If p = e_Skill.Talar Or p = e_Skill.Pescar Or p = e_Skill.Mineria Then
            cMax = CLng(val(prof_live_ini().GetValue(prof_section(p), "CantidadMax")))
            sp = UserList(1).Stats.UserSkills(p)
            UserList(1).Stats.UserSkills(p) = 255
            ok = ok And (CalcularCantidadExtraccion(1, p) = cMax)
            UserList(1).Stats.UserSkills(p) = sp
        End If
    Next p
    UserList(1).pos.Map = sm
    test_prof_cantidad_caps_at_config_max = ok
    Exit Function
e:
    UserList(1).pos.Map = sm
    test_prof_cantidad_caps_at_config_max = False
End Function

' Expected amount, recomputed from profesiones.ini with the documented formula:
'   Min + (Max-Min) * min(skill, SkillParaMax) / SkillParaMax, then -PenalizacionZonaSegura% in a safe zone, floor 1.
Private Function prof_expected_cantidad(ByVal profId As Integer, ByVal skill As Integer, ByVal safeMap As Boolean) As Integer
    Dim ini As clsIniManager, sec As String
    Dim skillMax As Integer, cMin As Integer, cMax As Integer, pen As Integer
    Dim rango As Integer, s As Integer, res As Single
    Set ini = prof_live_ini()
    sec = prof_section(profId)
    skillMax = CInt(val(ini.GetValue(sec, "SkillParaMax")))
    cMin = CInt(val(ini.GetValue(sec, "CantidadMin")))
    cMax = CInt(val(ini.GetValue(sec, "CantidadMax")))
    pen = CInt(val(ini.GetValue(sec, "PenalizacionZonaSegura", "0")))
    If skillMax <= 0 Then skillMax = 100
    If pen < 0 Then pen = 0
    If pen > 100 Then pen = 100
    s = skill
    If s > skillMax Then s = skillMax
    rango = cMax - cMin
    If rango < 0 Then rango = 0
    res = CSng(cMin) + (CSng(rango) * CSng(s) / CSng(skillMax))
    If safeMap And pen > 0 Then res = Int(res * CSng(100 - pen) / 100)
    If res < 1 Then res = 1
    prof_expected_cantidad = CInt(res)
End Function

' The amount follows the documented formula at several skills, and never decreases as skill grows.
Private Function test_prof_cantidad_matches_formula_and_is_monotonic() As Boolean
    On Error GoTo e
    Dim sp As Byte, sm As Integer, ok As Boolean, p As Long, k As Long, normalMap As Long
    Dim skills As Variant, prev As Long, cur As Long, s As Long
    normalMap = prof_find_map(0)
    If normalMap = 0 Then Exit Function
    skills = Array(0, 1, 25, 50, 75, 99, 100, 255)
    sm = UserList(1).pos.Map
    UserList(1).pos.Map = normalMap
    ok = True
    For p = e_Skill.Talar To e_Skill.Mineria
        If p = e_Skill.Talar Or p = e_Skill.Pescar Or p = e_Skill.Mineria Then
            sp = UserList(1).Stats.UserSkills(p)
            prev = 0
            For k = LBound(skills) To UBound(skills)
                s = skills(k)
                UserList(1).Stats.UserSkills(p) = CByte(s)
                cur = CalcularCantidadExtraccion(1, p)
                ok = ok And (cur = prof_expected_cantidad(p, CInt(s), False)) And (cur >= 1) And (cur >= prev)
                prev = cur
            Next k
            UserList(1).Stats.UserSkills(p) = sp
        End If
    Next p
    UserList(1).pos.Map = sm
    test_prof_cantidad_matches_formula_and_is_monotonic = ok
    Exit Function
e:
    UserList(1).pos.Map = sm
    test_prof_cantidad_matches_formula_and_is_monotonic = False
End Function

' Safe zone (MapInfo.Seguro = 1) applies PenalizacionZonaSegura: never more than outside, and equal to the formula.
Private Function test_prof_cantidad_safe_zone_never_exceeds_normal() As Boolean
    On Error GoTo e
    Dim sp As Byte, sm As Integer, ok As Boolean, p As Long, k As Long, normalMap As Long, safeMap As Long
    Dim skills As Variant, a As Long, b As Long, s As Long
    normalMap = prof_find_map(0)
    safeMap = prof_find_map(1)
    If normalMap = 0 Or safeMap = 0 Then Exit Function
    skills = Array(0, 50, 100, 255)
    sm = UserList(1).pos.Map
    ok = True
    For p = e_Skill.Talar To e_Skill.Mineria
        If p = e_Skill.Talar Or p = e_Skill.Pescar Or p = e_Skill.Mineria Then
            sp = UserList(1).Stats.UserSkills(p)
            For k = LBound(skills) To UBound(skills)
                s = skills(k)
                UserList(1).Stats.UserSkills(p) = CByte(s)
                UserList(1).pos.Map = normalMap
                a = CalcularCantidadExtraccion(1, p)
                UserList(1).pos.Map = safeMap
                b = CalcularCantidadExtraccion(1, p)
                ok = ok And (b <= a) And (b >= 1) And (b = prof_expected_cantidad(p, CInt(s), True))
            Next k
            UserList(1).Stats.UserSkills(p) = sp
        End If
    Next p
    UserList(1).pos.Map = sm
    test_prof_cantidad_safe_zone_never_exceeds_normal = ok
    Exit Function
e:
    UserList(1).pos.Map = sm
    test_prof_cantidad_safe_zone_never_exceeds_normal = False
End Function

' --- CalculateCraftProb (catalyst) ----------------------------------------------------------------------
' Mutates ObjData(1), user 1's catalyst slot and flags.Crafteando; restored.

Private Function craft_prob(ByVal catObj As Integer, ByVal catType As Byte, ByVal catIncrease As Single, ByVal crafting As Byte, ByVal baseProb As Byte) As Long
    ObjData(1).CatalizadorTipo = catType
    ObjData(1).CatalizadorAumento = catIncrease
    UserList(1).CraftCatalyst.ObjIndex = catObj
    UserList(1).flags.Crafteando = crafting
    craft_prob = CalculateCraftProb(1, baseProb)
End Function

' No catalyst in the slot: the base probability, untouched.
Private Function test_craft_prob_no_catalyst_is_base() As Boolean
    On Error GoTo e
    Dim so As t_ObjData, sc As Integer, sf As Byte, ok As Boolean
    so = ObjData(1)
    sc = UserList(1).CraftCatalyst.ObjIndex
    sf = UserList(1).flags.Crafteando
    ok = (craft_prob(0, 1, 0.5, 1, 40) = 40) And (craft_prob(0, 1, 0.5, 1, 0) = 0) And (craft_prob(0, 1, 0.5, 1, 100) = 100)
    ObjData(1) = so
    UserList(1).CraftCatalyst.ObjIndex = sc
    UserList(1).flags.Crafteando = sf
    test_craft_prob_no_catalyst_is_base = ok
    Exit Function
e:
    ObjData(1) = so
    UserList(1).CraftCatalyst.ObjIndex = sc
    UserList(1).flags.Crafteando = sf
    test_craft_prob_no_catalyst_is_base = False
End Function

' Catalyst of the matching type: Fix(base * (1 + increase)): 50*1.5 = 75, 40*1.25 = 50, 33*1.1 = 36 (truncated).
Private Function test_craft_prob_matching_catalyst_scales() As Boolean
    On Error GoTo e
    Dim so As t_ObjData, sc As Integer, sf As Byte, ok As Boolean
    so = ObjData(1)
    sc = UserList(1).CraftCatalyst.ObjIndex
    sf = UserList(1).flags.Crafteando
    ok = (craft_prob(1, 2, 0.5, 2, 50) = 75) And (craft_prob(1, 2, 0.25, 2, 40) = 50) And (craft_prob(1, 2, 0.1, 2, 33) = 36)
    ObjData(1) = so
    UserList(1).CraftCatalyst.ObjIndex = sc
    UserList(1).flags.Crafteando = sf
    test_craft_prob_matching_catalyst_scales = ok
    Exit Function
e:
    ObjData(1) = so
    UserList(1).CraftCatalyst.ObjIndex = sc
    UserList(1).flags.Crafteando = sf
    test_craft_prob_matching_catalyst_scales = False
End Function

' The boosted probability is clamped to 0..100: 80*1.5 = 120 -> 100; a negative increase below -1 -> 0.
Private Function test_craft_prob_matching_catalyst_clamps() As Boolean
    On Error GoTo e
    Dim so As t_ObjData, sc As Integer, sf As Byte, ok As Boolean
    so = ObjData(1)
    sc = UserList(1).CraftCatalyst.ObjIndex
    sf = UserList(1).flags.Crafteando
    ok = (craft_prob(1, 2, 0.5, 2, 80) = 100) And (craft_prob(1, 2, -2, 2, 50) = 0)
    ObjData(1) = so
    UserList(1).CraftCatalyst.ObjIndex = sc
    UserList(1).flags.Crafteando = sf
    test_craft_prob_matching_catalyst_clamps = ok
    Exit Function
e:
    ObjData(1) = so
    UserList(1).CraftCatalyst.ObjIndex = sc
    UserList(1).flags.Crafteando = sf
    test_craft_prob_matching_catalyst_clamps = False
End Function

' A catalyst made for another crafting type has no effect: base probability.
Private Function test_craft_prob_other_type_catalyst_ignored() As Boolean
    On Error GoTo e
    Dim so As t_ObjData, sc As Integer, sf As Byte, ok As Boolean
    so = ObjData(1)
    sc = UserList(1).CraftCatalyst.ObjIndex
    sf = UserList(1).flags.Crafteando
    ok = (craft_prob(1, 3, 0.5, 2, 50) = 50)
    ObjData(1) = so
    UserList(1).CraftCatalyst.ObjIndex = sc
    UserList(1).flags.Crafteando = sf
    test_craft_prob_other_type_catalyst_ignored = ok
    Exit Function
e:
    ObjData(1) = so
    UserList(1).CraftCatalyst.ObjIndex = sc
    UserList(1).flags.Crafteando = sf
    test_craft_prob_other_type_catalyst_ignored = False
End Function

#End If
