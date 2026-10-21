Attribute VB_Name = "Unit_Elemental"
Option Explicit
#If UNIT_TEST = 1 Then

Private Const ELEM_TEMP_INI As String = "test_temp_elemental.ini"

' ==========================================================================
' Elemental Damage Test Suite (plan 20.002, TOGGLE32 elemental_system)
' Tests the deterministic core of modElementalCombat: resistance cascade
' (immune / chance / flat / pct + cap) and component rolling bounds.
' Chance-based paths use 0 or 100 to stay deterministic.
' ==========================================================================
Public Function test_suite_elemental() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start

    Call InitDamageTypeRegistryDefaults

    Call UnitTesting.RunTest("test_resist_immune", test_resist_immune())
    Call UnitTesting.RunTest("test_resist_flat", test_resist_flat())
    Call UnitTesting.RunTest("test_resist_flat_overkill", test_resist_flat_overkill())
    Call UnitTesting.RunTest("test_resist_pct_half", test_resist_pct_half())
    Call UnitTesting.RunTest("test_resist_pct_capped", test_resist_pct_capped())
    Call UnitTesting.RunTest("test_resist_chance_full", test_resist_chance_full())
    Call UnitTesting.RunTest("test_resist_chance_zero", test_resist_chance_zero())
    Call UnitTesting.RunTest("test_resist_zero_raw", test_resist_zero_raw())
    Call UnitTesting.RunTest("test_roll_fixed", test_roll_fixed())
    Call UnitTesting.RunTest("test_roll_swapped", test_roll_swapped())
    Call UnitTesting.RunTest("test_roll_negative", test_roll_negative())
    Call UnitTesting.RunTest("test_damage_type_name", test_damage_type_name())

    ' Plan 21.002 (B1): characterization of the elemental arithmetic that the probe cannot see.
    Call UnitTesting.RunTest("test_resist_cap_per_type_override", test_resist_cap_per_type_override())
    Call UnitTesting.RunTest("test_resist_cap_other_type_unaffected", test_resist_cap_other_type_unaffected())
    Call UnitTesting.RunTest("test_resist_cap_out_of_range_type_uses_default", test_resist_cap_out_of_range_type_uses_default())
    Call UnitTesting.RunTest("test_resist_pct_negative_clamped_to_zero", test_resist_pct_negative_clamped_to_zero())
    Call UnitTesting.RunTest("test_resist_immune_damage_nullifies", test_resist_immune_damage_nullifies())
    Call UnitTesting.RunTest("test_resist_immune_effect_only_passes_damage", test_resist_immune_effect_only_passes_damage())
    Call UnitTesting.RunTest("test_resist_chance_over_100_clamped", test_resist_chance_over_100_clamped())
    Call UnitTesting.RunTest("test_resist_flat_range_degenerate", test_resist_flat_range_degenerate())
    Call UnitTesting.RunTest("test_resist_flat_range_bounds", test_resist_flat_range_bounds())
    Call UnitTesting.RunTest("test_resist_pct_range_degenerate", test_resist_pct_range_degenerate())
    Call UnitTesting.RunTest("test_resist_fractional_rounding_bounds", test_resist_fractional_rounding_bounds())
    Call UnitTesting.RunTest("test_resist_negative_raw_is_zero", test_resist_negative_raw_is_zero())
    Call UnitTesting.RunTest("test_dot_base_fixed_and_range", test_dot_base_fixed_and_range())
    Call UnitTesting.RunTest("test_dot_base_pct_hp_user", test_dot_base_pct_hp_user())
    Call UnitTesting.RunTest("test_dot_base_pct_hp_npc", test_dot_base_pct_hp_npc())
    Call UnitTesting.RunTest("test_dot_base_edge_cases", test_dot_base_edge_cases())
    Call UnitTesting.RunTest("test_dot_tick_resist_toggle_off_is_raw", test_dot_tick_resist_toggle_off_is_raw())
    Call UnitTesting.RunTest("test_dot_tick_resist_applies_flat", test_dot_tick_resist_applies_flat())
    Call UnitTesting.RunTest("test_dot_tick_resist_immune", test_dot_tick_resist_immune())
    Call UnitTesting.RunTest("test_dot_tick_resist_passthrough_cases", test_dot_tick_resist_passthrough_cases())
    Call UnitTesting.RunTest("test_capped_speed_cap_inactive", test_capped_speed_cap_inactive())
    Call UnitTesting.RunTest("test_capped_speed_cap_active_floor", test_capped_speed_cap_active_floor())
    Call UnitTesting.RunTest("test_capped_speed_positive_and_zero", test_capped_speed_positive_and_zero())
    Call UnitTesting.RunTest("test_damage_type_from_color_roundtrip", test_damage_type_from_color_roundtrip())
    Call UnitTesting.RunTest("test_damage_type_from_color_sentinels", test_damage_type_from_color_sentinels())
    Call UnitTesting.RunTest("test_damage_type_name_from_color", test_damage_type_name_from_color())
    Call UnitTesting.RunTest("test_balance_catalog_membership", test_balance_catalog_membership())
    Call UnitTesting.RunTest("test_balance_catalog_tiers", test_balance_catalog_tiers())
    Call UnitTesting.RunTest("test_parse_source_components_and_clamp", test_parse_source_components_and_clamp())
    Call UnitTesting.RunTest("test_parse_source_proc_fields", test_parse_source_proc_fields())
    Call UnitTesting.RunTest("test_parse_source_missing_section_resets_counts", test_parse_source_missing_section_resets_counts())
    Call UnitTesting.RunTest("test_parse_source_negative_counts_and_typo_keys", test_parse_source_negative_counts_and_typo_keys())
    Call UnitTesting.RunTest("test_parse_resist_fields_and_clamp", test_parse_resist_fields_and_clamp())
    Call UnitTesting.RunTest("test_parse_resist_missing_and_negative", test_parse_resist_missing_and_negative())
    Call elem_delete_temp_ini

    Debug.Print "Elemental suite took " & sw.ElapsedMilliseconds & " ms"
    test_suite_elemental = True
End Function

' Immunity nullifies the whole component and flags outNullified.
Private Function test_resist_immune() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean, d As Long
    r.Immune = 1
    d = ApplyElementalResist(100, r, eDmgFire, nul)
    test_resist_immune = (d = 0 And nul = True)
    Exit Function
e:
    test_resist_immune = False
End Function

' Flat reduction subtracts before pct: 10 - 3 = 7.
Private Function test_resist_flat() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean
    r.ReduceFlat = 3
    test_resist_flat = (ApplyElementalResist(10, r, eDmgFire, nul) = 7)
    Exit Function
e:
    test_resist_flat = False
End Function

' Flat greater than raw clamps to 0.
Private Function test_resist_flat_overkill() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean
    r.ReduceFlat = 10
    test_resist_flat_overkill = (ApplyElementalResist(5, r, eDmgFire, nul) = 0)
    Exit Function
e:
    test_resist_flat_overkill = False
End Function

' 50% pct reduction: 100 -> 50.
Private Function test_resist_pct_half() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean
    r.ReducePct = 0.5
    test_resist_pct_half = (ApplyElementalResist(100, r, eDmgFire, nul) = 50)
    Exit Function
e:
    test_resist_pct_half = False
End Function

' Pct above the default cap (0.75) is clamped: 0.9 -> 0.75, so 100 -> 25.
Private Function test_resist_pct_capped() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean
    r.ReducePct = 0.9
    test_resist_pct_capped = (ApplyElementalResist(100, r, eDmgFire, nul) = 25)
    Exit Function
e:
    test_resist_pct_capped = False
End Function

' ReduceChancePct=100 always nullifies.
Private Function test_resist_chance_full() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean, d As Long
    r.ReduceChancePct = 100
    d = ApplyElementalResist(100, r, eDmgFire, nul)
    test_resist_chance_full = (d = 0 And nul = True)
    Exit Function
e:
    test_resist_chance_full = False
End Function

' ReduceChancePct=0 never nullifies; full damage passes.
Private Function test_resist_chance_zero() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean
    test_resist_chance_zero = (ApplyElementalResist(10, r, eDmgFire, nul) = 10 And nul = False)
    Exit Function
e:
    test_resist_chance_zero = False
End Function

' Zero raw damage stays zero.
Private Function test_resist_zero_raw() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean
    r.ReducePct = 0.5
    test_resist_zero_raw = (ApplyElementalResist(0, r, eDmgFire, nul) = 0)
    Exit Function
e:
    test_resist_zero_raw = False
End Function

' min==max rolls the fixed value.
Private Function test_roll_fixed() As Boolean
    On Error GoTo e
    Dim c As t_DamageComponent
    c.DamageType = eDmgFire: c.MinDamage = 7: c.MaxDamage = 7
    test_roll_fixed = (RollDamageComponent(c) = 7)
    Exit Function
e:
    test_roll_fixed = False
End Function

' min>max is sanitized (hi=lo), returns lo.
Private Function test_roll_swapped() As Boolean
    On Error GoTo e
    Dim c As t_DamageComponent
    c.DamageType = eDmgFire: c.MinDamage = 9: c.MaxDamage = 3
    test_roll_swapped = (RollDamageComponent(c) = 9)
    Exit Function
e:
    test_roll_swapped = False
End Function

' negative bounds clamp to 0.
Private Function test_roll_negative() As Boolean
    On Error GoTo e
    Dim c As t_DamageComponent
    c.DamageType = eDmgFire: c.MinDamage = -5: c.MaxDamage = -1
    test_roll_negative = (RollDamageComponent(c) = 0)
    Exit Function
e:
    test_roll_negative = False
End Function

' Registry default name for fire.
Private Function test_damage_type_name() As Boolean
    On Error GoTo e
    test_damage_type_name = (DamageTypeName(eDmgFire) = "Fuego")
    Exit Function
e:
    test_damage_type_name = False
End Function

' ==========================================================================
' Plan 21.002 (B1): extended characterization of modElementalCombat and
' modElementalBalanceLog. Pure arithmetic only; chance paths use 0 or 100 or
' check bounds. Anything that mutates global state restores it.
' ==========================================================================
Private Sub elem_delete_temp_ini()
    On Error Resume Next
    Kill App.Path & "\" & ELEM_TEMP_INI
    On Error GoTo 0
End Sub

Private Function elem_write_temp_ini(ByVal content As String) As String
    Dim filePath As String
    Dim f As Integer
    filePath = App.Path & "\" & ELEM_TEMP_INI
    f = FreeFile
    Open filePath For Output As #f
    Print #f, content
    Close #f
    elem_write_temp_ini = filePath
End Function

' --- resist cascade: cap per type ---------------------------------------

' A type with its own ResistCapPct uses it instead of the 0.75 default: 0.9 -> 0.5, so 100 -> 50.
Private Function test_resist_cap_per_type_override() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean, savedCap As Single
    savedCap = DamageTypeReg(eDmgFire).ResistCapPct
    DamageTypeReg(eDmgFire).ResistCapPct = 0.5
    r.ReducePct = 0.9
    test_resist_cap_per_type_override = (ApplyElementalResist(100, r, eDmgFire, nul) = 50)
    DamageTypeReg(eDmgFire).ResistCapPct = savedCap
    Exit Function
e:
    DamageTypeReg(eDmgFire).ResistCapPct = savedCap
    test_resist_cap_per_type_override = False
End Function

' The override of one type does not leak into another: frost still caps at the default 0.75.
Private Function test_resist_cap_other_type_unaffected() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean, savedCap As Single
    savedCap = DamageTypeReg(eDmgFire).ResistCapPct
    DamageTypeReg(eDmgFire).ResistCapPct = 0.5
    r.ReducePct = 0.9
    test_resist_cap_other_type_unaffected = (ApplyElementalResist(100, r, eDmgFrost, nul) = 25)
    DamageTypeReg(eDmgFire).ResistCapPct = savedCap
    Exit Function
e:
    DamageTypeReg(eDmgFire).ResistCapPct = savedCap
    test_resist_cap_other_type_unaffected = False
End Function

' Type ids outside 1..MAX_DAMAGE_TYPE_ID fall back to the default cap (no array overrun).
Private Function test_resist_cap_out_of_range_type_uses_default() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean
    r.ReducePct = 0.9
    test_resist_cap_out_of_range_type_uses_default = (ApplyElementalResist(100, r, eDmgNone, nul) = 25) _
        And (ApplyElementalResist(100, r, MAX_DAMAGE_TYPE_ID + 1, nul) = 25) _
        And (ApplyElementalResist(100, r, 99, nul) = 25) _
        And (ApplyElementalResist(100, r, -1, nul) = 25)
    Exit Function
e:
    test_resist_cap_out_of_range_type_uses_default = False
End Function

' A negative ReducePct never amplifies damage: it is floored at 0.
Private Function test_resist_pct_negative_clamped_to_zero() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean
    r.ReducePct = -0.5
    test_resist_pct_negative_clamped_to_zero = (ApplyElementalResist(100, r, eDmgFire, nul) = 100 And nul = False)
    Exit Function
e:
    test_resist_pct_negative_clamped_to_zero = False
End Function

' ImmuneDamage (damage-only immunity) nullifies the component like Immune does.
Private Function test_resist_immune_damage_nullifies() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean, d As Long
    r.ImmuneDamage = 1
    d = ApplyElementalResist(100, r, eDmgFire, nul)
    test_resist_immune_damage_nullifies = (d = 0 And nul = True)
    Exit Function
e:
    test_resist_immune_damage_nullifies = False
End Function

' ImmuneEffect only blocks the effect (checked elsewhere): the damage still goes through untouched.
Private Function test_resist_immune_effect_only_passes_damage() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean, d As Long
    r.ImmuneEffect = 1
    d = ApplyElementalResist(100, r, eDmgFire, nul)
    test_resist_immune_effect_only_passes_damage = (d = 100 And nul = False)
    Exit Function
e:
    test_resist_immune_effect_only_passes_damage = False
End Function

' ReduceChancePct above 100 is clamped to 100: always nullifies, never errors.
Private Function test_resist_chance_over_100_clamped() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean, i As Long
    r.ReduceChancePct = 250
    For i = 1 To 30
        If ApplyElementalResist(100, r, eDmgFire, nul) <> 0 Or nul = False Then
            test_resist_chance_over_100_clamped = False
            Exit Function
        End If
    Next i
    test_resist_chance_over_100_clamped = True
    Exit Function
e:
    test_resist_chance_over_100_clamped = False
End Function

' FlatMax > 0 switches to the range and ignores ReduceFlat; Min>=Max is deterministic: 10 - 4 = 6.
Private Function test_resist_flat_range_degenerate() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean
    r.ReduceFlat = 9
    r.ReduceFlatMin = 4
    r.ReduceFlatMax = 4
    test_resist_flat_range_degenerate = (ApplyElementalResist(10, r, eDmgFire, nul) = 6)
    Exit Function
e:
    test_resist_flat_range_degenerate = False
End Function

' Flat range 2..5 on 10 damage always lands in 5..8.
Private Function test_resist_flat_range_bounds() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean, i As Long, d As Long
    r.ReduceFlatMin = 2
    r.ReduceFlatMax = 5
    For i = 1 To 60
        d = ApplyElementalResist(10, r, eDmgFire, nul)
        If d < 5 Or d > 8 Then
            test_resist_flat_range_bounds = False
            Exit Function
        End If
    Next i
    test_resist_flat_range_bounds = True
    Exit Function
e:
    test_resist_flat_range_bounds = False
End Function

' PctMax > 0 switches to the integer range and ignores ReducePct; Min>=Max is deterministic: 50% of 100 = 50.
Private Function test_resist_pct_range_degenerate() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean
    r.ReducePct = 0.1
    r.ReducePctMin = 50
    r.ReducePctMax = 50
    test_resist_pct_range_degenerate = (ApplyElementalResist(100, r, eDmgFire, nul) = 50)
    Exit Function
e:
    test_resist_pct_range_degenerate = False
End Function

' Probabilistic rounding (plan 10.001 D21): 33.3% off 10 is 6.67, so the result is 6 or 7, never anything else.
Private Function test_resist_fractional_rounding_bounds() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean, i As Long, d As Long
    r.ReducePct = 0.333
    For i = 1 To 100
        d = ApplyElementalResist(10, r, eDmgFire, nul)
        If d < 6 Or d > 7 Then
            test_resist_fractional_rounding_bounds = False
            Exit Function
        End If
    Next i
    test_resist_fractional_rounding_bounds = True
    Exit Function
e:
    test_resist_fractional_rounding_bounds = False
End Function

' Negative raw damage is treated like zero: result 0, not nullified.
Private Function test_resist_negative_raw_is_zero() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean
    r.ReduceFlat = 3
    test_resist_negative_raw_is_zero = (ApplyElementalResist(-5, r, eDmgFire, nul) = 0 And nul = False)
    Exit Function
e:
    test_resist_negative_raw_is_zero = False
End Function

' --- RollDotTickBase ------------------------------------------------------

' Mode 0 = fixed (Min), mode 1 = range; sign is preserved.
Private Function test_dot_base_fixed_and_range() As Boolean
    On Error GoTo e
    Dim i As Long, v As Long
    If RollDotTickBase(1, e_ReferenceType.eUser, 0, 7, 99) <> 7 Then Exit Function
    If RollDotTickBase(1, e_ReferenceType.eUser, 0, -4, -9) <> -4 Then Exit Function
    If RollDotTickBase(1, e_ReferenceType.eUser, 1, 5, 5) <> 5 Then Exit Function
    For i = 1 To 50
        v = RollDotTickBase(1, e_ReferenceType.eUser, 1, 3, 8)
        If v < 3 Or v > 8 Then Exit Function
    Next i
    test_dot_base_fixed_and_range = True
    Exit Function
e:
    test_dot_base_fixed_and_range = False
End Function

' Modes 2/3 = percent of the target's max HP (integer division): 10% of 200 = 20.
Private Function test_dot_base_pct_hp_user() As Boolean
    On Error GoTo e
    Dim savedHp As Long
    savedHp = UserList(1).Stats.MaxHp
    UserList(1).Stats.MaxHp = 200
    test_dot_base_pct_hp_user = (RollDotTickBase(1, e_ReferenceType.eUser, 2, 10, 99) = 20) _
        And (RollDotTickBase(1, e_ReferenceType.eUser, 3, 10, 10) = 20) _
        And (RollDotTickBase(1, e_ReferenceType.eUser, 2, 1, 1) = 2) _
        And (RollDotTickBase(1, e_ReferenceType.eUser, 2, -10, 0) = -20)
    UserList(1).Stats.MaxHp = savedHp
    Exit Function
e:
    UserList(1).Stats.MaxHp = savedHp
    test_dot_base_pct_hp_user = False
End Function

' Same percent-of-max-HP mode against an NPC target.
Private Function test_dot_base_pct_hp_npc() As Boolean
    On Error GoTo e
    Dim savedHp As Long
    savedHp = NpcList(1).Stats.MaxHp
    NpcList(1).Stats.MaxHp = 50
    test_dot_base_pct_hp_npc = (RollDotTickBase(1, e_ReferenceType.eNpc, 2, 20, 0) = 10) _
        And (RollDotTickBase(1, e_ReferenceType.eNpc, 3, 20, 20) = 10)
    NpcList(1).Stats.MaxHp = savedHp
    Exit Function
e:
    NpcList(1).Stats.MaxHp = savedHp
    test_dot_base_pct_hp_npc = False
End Function

' Unknown mode returns 0; percent mode against a target type with no HP source returns 0.
Private Function test_dot_base_edge_cases() As Boolean
    On Error GoTo e
    test_dot_base_edge_cases = (RollDotTickBase(1, e_ReferenceType.eUser, 9, 10, 10) = 0) _
        And (RollDotTickBase(1, e_ReferenceType.eNone, 2, 10, 10) = 0)
    Exit Function
e:
    test_dot_base_edge_cases = False
End Function

' --- ApplyDotTickResist -----------------------------------------------------
' Fixture: user 1 wears ObjData(1) as armor, and that object carries a fire resist.

Private Function dot_resist_setup(ByRef savedObj As t_ObjData, ByRef savedArmor As Integer, ByRef savedSystem As Boolean, ByRef savedLog As Boolean) As Boolean
    savedObj = ObjData(1)
    savedArmor = UserList(1).invent.EquippedArmorObjIndex
    savedSystem = IsFeatureEnabled("elemental_system")
    savedLog = IsFeatureEnabled("elemental_balance_log")
    Dim emptyResist As t_ElementalResist
    ObjData(1).ElementalResist.Resist(1) = emptyResist
    ObjData(1).ElementalResist.Count = 1
    ObjData(1).ElementalResist.Resist(1).DamageType = eDmgFire
    ObjData(1).ElementalResist.Resist(1).ReduceFlat = 3
    ObjData(1).ElementalResist.Resist(1).Immune = 0
    UserList(1).invent.EquippedArmorObjIndex = 1
    Call SetFeatureToggle("elemental_balance_log", False)
    dot_resist_setup = True
End Function

Private Sub dot_resist_teardown(ByRef savedObj As t_ObjData, ByVal savedArmor As Integer, ByVal savedSystem As Boolean, ByVal savedLog As Boolean)
    ObjData(1) = savedObj
    UserList(1).invent.EquippedArmorObjIndex = savedArmor
    Call SetFeatureToggle("elemental_system", savedSystem)
    Call SetFeatureToggle("elemental_balance_log", savedLog)
End Sub

' With the master toggle OFF the tick is returned untouched even if the target has a resist.
Private Function test_dot_tick_resist_toggle_off_is_raw() As Boolean
    On Error GoTo e
    Dim so As t_ObjData, sa As Integer, ss As Boolean, sl As Boolean
    Call dot_resist_setup(so, sa, ss, sl)
    Call SetFeatureToggle("elemental_system", False)
    test_dot_tick_resist_toggle_off_is_raw = (ApplyDotTickResist(False, 1, eDmgFire, 10) = 10)
    Call dot_resist_teardown(so, sa, ss, sl)
    Exit Function
e:
    Call dot_resist_teardown(so, sa, ss, sl)
    test_dot_tick_resist_toggle_off_is_raw = False
End Function

' Toggle ON: the worn item's flat fire resist reduces the tick, 10 - 3 = 7.
Private Function test_dot_tick_resist_applies_flat() As Boolean
    On Error GoTo e
    Dim so As t_ObjData, sa As Integer, ss As Boolean, sl As Boolean
    Call dot_resist_setup(so, sa, ss, sl)
    Call SetFeatureToggle("elemental_system", True)
    test_dot_tick_resist_applies_flat = (ApplyDotTickResist(False, 1, eDmgFire, 10) = 7)
    Call dot_resist_teardown(so, sa, ss, sl)
    Exit Function
e:
    Call dot_resist_teardown(so, sa, ss, sl)
    test_dot_tick_resist_applies_flat = False
End Function

' Toggle ON and an Immune fire resist: the tick becomes 0 (the burn-on-immune bug of plan 02.001 M6).
Private Function test_dot_tick_resist_immune() As Boolean
    On Error GoTo e
    Dim so As t_ObjData, sa As Integer, ss As Boolean, sl As Boolean
    Call dot_resist_setup(so, sa, ss, sl)
    Call SetFeatureToggle("elemental_system", True)
    ObjData(1).ElementalResist.Resist(1).Immune = 1
    test_dot_tick_resist_immune = (ApplyDotTickResist(False, 1, eDmgFire, 10) = 0)
    Call dot_resist_teardown(so, sa, ss, sl)
    Exit Function
e:
    Call dot_resist_teardown(so, sa, ss, sl)
    test_dot_tick_resist_immune = False
End Function

' Passthrough: no damage type, zero/negative raw, a different type, or nothing worn all return the raw value.
Private Function test_dot_tick_resist_passthrough_cases() As Boolean
    On Error GoTo e
    Dim so As t_ObjData, sa As Integer, ss As Boolean, sl As Boolean, ok As Boolean
    Call dot_resist_setup(so, sa, ss, sl)
    Call SetFeatureToggle("elemental_system", True)
    ok = (ApplyDotTickResist(False, 1, 0, 10) = 10)
    ok = ok And (ApplyDotTickResist(False, 1, eDmgFire, 0) = 0)
    ok = ok And (ApplyDotTickResist(False, 1, eDmgFire, -5) = -5)
    ok = ok And (ApplyDotTickResist(False, 1, eDmgFrost, 10) = 10)
    UserList(1).invent.EquippedArmorObjIndex = 0
    ok = ok And (ApplyDotTickResist(False, 1, eDmgFire, 10) = 10)
    test_dot_tick_resist_passthrough_cases = ok
    Call dot_resist_teardown(so, sa, ss, sl)
    Exit Function
e:
    Call dot_resist_teardown(so, sa, ss, sl)
    test_dot_tick_resist_passthrough_cases = False
End Function

' --- CappedSpeedMult ---------------------------------------------------------

' Cap inactive (toggle OFF): vanilla, the multiplier floors at 0 (so -1 gives exactly 0).
Private Function test_capped_speed_cap_inactive() As Boolean
    On Error GoTo e
    Dim saved As Boolean
    saved = ElementalSlowCapActive
    ElementalSlowCapActive = False
    test_capped_speed_cap_inactive = (CappedSpeedMult(-1) = 0) And (CappedSpeedMult(-5) = 0) And (CappedSpeedMult(-0.5) = 0.5)
    ElementalSlowCapActive = saved
    Exit Function
e:
    ElementalSlowCapActive = saved
    test_capped_speed_cap_inactive = False
End Function

' Cap active: slows never freeze. Any deep slow lands on the same positive floor (<= 1), and -1 is not 0.
Private Function test_capped_speed_cap_active_floor() As Boolean
    On Error GoTo e
    Dim saved As Boolean, a As Single, b As Single, c As Single
    saved = ElementalSlowCapActive
    ElementalSlowCapActive = True
    a = CappedSpeedMult(-1)
    b = CappedSpeedMult(-5)
    c = CappedSpeedMult(-100)
    test_capped_speed_cap_active_floor = (a > 0) And (a <= 1) And (a = b) And (b = c)
    ElementalSlowCapActive = saved
    Exit Function
e:
    ElementalSlowCapActive = saved
    test_capped_speed_cap_active_floor = False
End Function

' Zero and positive modifiers are untouched by the cap in both modes: 1+mod.
Private Function test_capped_speed_positive_and_zero() As Boolean
    On Error GoTo e
    Dim saved As Boolean, ok As Boolean
    saved = ElementalSlowCapActive
    ElementalSlowCapActive = True
    ok = (CappedSpeedMult(0) = 1) And (CappedSpeedMult(0.5) = 1.5)
    ElementalSlowCapActive = False
    ok = ok And (CappedSpeedMult(0) = 1) And (CappedSpeedMult(0.5) = 1.5)
    ElementalSlowCapActive = saved
    test_capped_speed_positive_and_zero = ok
    Exit Function
e:
    ElementalSlowCapActive = saved
    test_capped_speed_positive_and_zero = False
End Function

' --- damage type <-> color ----------------------------------------------------

' Every non-physical built-in type round-trips: type -> color -> type.
Private Function test_damage_type_from_color_roundtrip() As Boolean
    On Error GoTo e
    Dim t As Long
    For t = eDmgFire To eDmgBleed
        If DamageTypeFromColor(DamageTypeColor(t)) <> t Then Exit Function
    Next t
    test_damage_type_from_color_roundtrip = True
    Exit Function
e:
    test_damage_type_from_color_roundtrip = False
End Function

' Color 0, white (physical) and a color nobody uses map to eDmgNone.
Private Function test_damage_type_from_color_sentinels() As Boolean
    On Error GoTo e
    test_damage_type_from_color_sentinels = (DamageTypeFromColor(0) = eDmgNone) _
        And (DamageTypeFromColor(vbWhite) = eDmgNone) _
        And (DamageTypeFromColor(RGB(1, 2, 3)) = eDmgNone)
    Exit Function
e:
    test_damage_type_from_color_sentinels = False
End Function

' Name from color: the registry name of the type, or "" for 0/white/unknown.
Private Function test_damage_type_name_from_color() As Boolean
    On Error GoTo e
    Dim t As Long
    For t = eDmgFire To eDmgBleed
        If DamageTypeNameFromColor(DamageTypeColor(t)) <> DamageTypeName(t) Then Exit Function
    Next t
    test_damage_type_name_from_color = (DamageTypeNameFromColor(RGB(255, 140, 0)) = "Fuego") _
        And (LenB(DamageTypeNameFromColor(0)) = 0) _
        And (LenB(DamageTypeNameFromColor(vbWhite)) = 0) _
        And (LenB(DamageTypeNameFromColor(RGB(1, 2, 3))) = 0)
    Exit Function
e:
    test_damage_type_name_from_color = False
End Function

' --- balance catalog ------------------------------------------------------------

' Catalog membership is the inclusive range OBJ9084..OBJ9161, including items with no explicit tier.
Private Function test_balance_catalog_membership() As Boolean
    On Error GoTo e
    test_balance_catalog_membership = (Not modElementalBalanceLog.ElementalBalanceInCatalog(9083)) _
        And modElementalBalanceLog.ElementalBalanceInCatalog(9084) _
        And modElementalBalanceLog.ElementalBalanceInCatalog(9096) _
        And modElementalBalanceLog.ElementalBalanceInCatalog(9161) _
        And (Not modElementalBalanceLog.ElementalBalanceInCatalog(9162)) _
        And (Not modElementalBalanceLog.ElementalBalanceInCatalog(0)) _
        And (Not modElementalBalanceLog.ElementalBalanceInCatalog(-5))
    Exit Function
e:
    test_balance_catalog_membership = False
End Function

' Tier boundaries per documented range (1 or 2); untiered catalog items (arrows, upper oils, orbs,
' thorns gear, poison oils, antidotes, scrolls) and anything outside the catalog return 0.
Private Function test_balance_catalog_tiers() As Boolean
    On Error GoTo e
    Dim ok As Boolean
    ok = True
    ok = ok And elem_tier_is(9083, 0) And elem_tier_is(9084, 1) And elem_tier_is(9089, 1)
    ok = ok And elem_tier_is(9090, 2) And elem_tier_is(9095, 2) And elem_tier_is(9096, 0)
    ok = ok And elem_tier_is(9100, 0) And elem_tier_is(9101, 1) And elem_tier_is(9105, 1)
    ok = ok And elem_tier_is(9106, 0) And elem_tier_is(9115, 0) And elem_tier_is(9116, 1)
    ok = ok And elem_tier_is(9121, 1) And elem_tier_is(9122, 2) And elem_tier_is(9127, 2)
    ok = ok And elem_tier_is(9128, 2) And elem_tier_is(9129, 2) And elem_tier_is(9130, 0)
    ok = ok And elem_tier_is(9142, 0) And elem_tier_is(9143, 1) And elem_tier_is(9148, 1)
    ok = ok And elem_tier_is(9149, 1) And elem_tier_is(9154, 1) And elem_tier_is(9155, 2)
    ok = ok And elem_tier_is(9157, 2) And elem_tier_is(9158, 2) And elem_tier_is(9160, 2)
    ok = ok And elem_tier_is(9161, 0) And elem_tier_is(9162, 0) And elem_tier_is(0, 0) And elem_tier_is(-1, 0)
    test_balance_catalog_tiers = ok
    Exit Function
e:
    test_balance_catalog_tiers = False
End Function

Private Function elem_tier_is(ByVal objIndex As Long, ByVal expected As Long) As Boolean
    elem_tier_is = (modElementalBalanceLog.ElementalBalanceCatalogTier(objIndex) = expected)
End Function

' --- INI parsers --------------------------------------------------------------------

Private Function elem_parse_ini_text() As String
    Dim s As String
    s = "[OBJ1]" & vbCrLf & _
        "ElemCompCount=6" & vbCrLf & _
        "ElemComp1Type=2" & vbCrLf & "ElemComp1Min=3" & vbCrLf & "ElemComp1Max=9" & vbCrLf & _
        "ElemComp2Type=3" & vbCrLf & "ElemComp2Min=1" & vbCrLf & "ElemComp2Max=1" & vbCrLf & _
        "ElemComp4Type=5" & vbCrLf & "ElemComp4Min=7" & vbCrLf & "ElemComp4Max=8" & vbCrLf & _
        "ElemProcCount=2" & vbCrLf & _
        "ElemProc1Chance=25" & vbCrLf & "ElemProc1Kind=1" & vbCrLf & "ElemProc1Trigger=1" & vbCrLf & _
        "ElemProc1Type=2" & vbCrLf & "ElemProc1Min=4" & vbCrLf & "ElemProc1Max=6" & vbCrLf & _
        "ElemProc1Eot=11" & vbCrLf & "ElemProc1ReflectPct=30" & vbCrLf & "ElemProc1Physical=1" & vbCrLf & _
        "ElemProc1BypassResist=1" & vbCrLf & "ElemProc1Lethal=1" & vbCrLf & "ElemProc1ReflectMirror=1" & vbCrLf & _
        "ElemProc2Chance=100" & vbCrLf & _
        "[EMPTY]" & vbCrLf & _
        "ElemCompCount=-3" & vbCrLf & "ElemProcCount=-1" & vbCrLf & _
        "[TYPO]" & vbCrLf & _
        "ElemCompCount=1" & vbCrLf & "ElemComp1Typ=2" & vbCrLf & "ElemComp1Min=5" & vbCrLf & _
        "ElemProcCount=9" & vbCrLf & _
        "[RES]" & vbCrLf & _
        "ElemResistCount=12" & vbCrLf & _
        "ElemResist1Type=2" & vbCrLf & "ElemResist1Chance=10" & vbCrLf & "ElemResist1Flat=3" & vbCrLf & _
        "ElemResist1Pct=0.5" & vbCrLf & "ElemResist1FlatMin=1" & vbCrLf & "ElemResist1FlatMax=2" & vbCrLf & _
        "ElemResist1PctMin=5" & vbCrLf & "ElemResist1PctMax=15" & vbCrLf & "ElemResist1Immune=1" & vbCrLf & _
        "ElemResist1ImmuneDmg=1" & vbCrLf & "ElemResist1ImmuneEff=1" & vbCrLf & "ElemResist1EffChance=35" & vbCrLf & _
        "ElemResist8Type=7" & vbCrLf & "ElemResist9Type=3" & vbCrLf & _
        "[RESEMPTY]" & vbCrLf & "ElemResistCount=-2"
    elem_parse_ini_text = s
End Function

' ElemCompCount above MAX_DAMAGE_COMPONENTS (4) is clamped; fields are read per index; gaps are zero.
Private Function test_parse_source_components_and_clamp() As Boolean
    On Error GoTo e
    Dim ini As New clsIniManager
    Dim src As t_ElementalSource
    Call ini.Initialize(elem_write_temp_ini(elem_parse_ini_text()))
    Call ParseElementalSourceFromIni(ini, "OBJ1", src)
    test_parse_source_components_and_clamp = (src.CompCount = MAX_DAMAGE_COMPONENTS) _
        And (src.Comp(1).DamageType = 2) And (src.Comp(1).MinDamage = 3) And (src.Comp(1).MaxDamage = 9) _
        And (src.Comp(2).DamageType = 3) And (src.Comp(2).MinDamage = 1) And (src.Comp(2).MaxDamage = 1) _
        And (src.Comp(3).DamageType = 0) And (src.Comp(3).MinDamage = 0) And (src.Comp(3).MaxDamage = 0) _
        And (src.Comp(4).DamageType = 5) And (src.Comp(4).MinDamage = 7) And (src.Comp(4).MaxDamage = 8)
    Exit Function
e:
    test_parse_source_components_and_clamp = False
End Function

' All twelve proc keys map to their field; a proc with only Chance leaves the rest at 0.
Private Function test_parse_source_proc_fields() As Boolean
    On Error GoTo e
    Dim ini As New clsIniManager
    Dim src As t_ElementalSource
    Call ini.Initialize(elem_write_temp_ini(elem_parse_ini_text()))
    Call ParseElementalSourceFromIni(ini, "OBJ1", src)
    test_parse_source_proc_fields = (src.ProcCount = 2) _
        And (src.Proc(1).ChancePct = 25) And (src.Proc(1).Kind = 1) And (src.Proc(1).Trigger = 1) _
        And (src.Proc(1).DamageType = 2) And (src.Proc(1).MinDamage = 4) And (src.Proc(1).MaxDamage = 6) _
        And (src.Proc(1).EotId = 11) And (src.Proc(1).ReflectPct = 30) And (src.Proc(1).Physical = 1) _
        And (src.Proc(1).BypassResist = 1) And (src.Proc(1).Lethal = 1) And (src.Proc(1).ReflectMirror = 1) _
        And (src.Proc(2).ChancePct = 100) And (src.Proc(2).Kind = 0) And (src.Proc(2).Trigger = 0) _
        And (src.Proc(2).EotId = 0) And (src.Proc(2).MinDamage = 0)
    Exit Function
e:
    test_parse_source_proc_fields = False
End Function

' A section that does not exist resets both counts to 0 even if the target had stale counts.
Private Function test_parse_source_missing_section_resets_counts() As Boolean
    On Error GoTo e
    Dim ini As New clsIniManager
    Dim src As t_ElementalSource
    Call ini.Initialize(elem_write_temp_ini(elem_parse_ini_text()))
    src.CompCount = 3
    src.ProcCount = 3
    Call ParseElementalSourceFromIni(ini, "NO_SUCH_SECTION", src)
    test_parse_source_missing_section_resets_counts = (src.CompCount = 0 And src.ProcCount = 0)
    Exit Function
e:
    test_parse_source_missing_section_resets_counts = False
End Function

' Negative counts clamp to 0; a misspelled key is ignored (field stays 0); ElemProcCount above 4 clamps to 4.
Private Function test_parse_source_negative_counts_and_typo_keys() As Boolean
    On Error GoTo e
    Dim ini As New clsIniManager
    Dim a As t_ElementalSource, b As t_ElementalSource
    Call ini.Initialize(elem_write_temp_ini(elem_parse_ini_text()))
    Call ParseElementalSourceFromIni(ini, "EMPTY", a)
    Call ParseElementalSourceFromIni(ini, "TYPO", b)
    test_parse_source_negative_counts_and_typo_keys = (a.CompCount = 0) And (a.ProcCount = 0) _
        And (b.CompCount = 1) And (b.Comp(1).DamageType = 0) And (b.Comp(1).MinDamage = 5) _
        And (b.ProcCount = MAX_ELEMENTAL_PROCS)
    Exit Function
e:
    test_parse_source_negative_counts_and_typo_keys = False
End Function

' All resist keys map to their field; ElemResistCount above MAX_ELEMENTAL_RESISTS (8) clamps; entry 9 is not read.
Private Function test_parse_resist_fields_and_clamp() As Boolean
    On Error GoTo e
    Dim ini As New clsIniManager
    Dim rs As t_ElementalResistSet
    Call ini.Initialize(elem_write_temp_ini(elem_parse_ini_text()))
    Call ParseElementalResistFromIni(ini, "RES", rs)
    test_parse_resist_fields_and_clamp = (rs.Count = MAX_ELEMENTAL_RESISTS) _
        And (rs.Resist(1).DamageType = 2) And (rs.Resist(1).ReduceChancePct = 10) And (rs.Resist(1).ReduceFlat = 3) _
        And (rs.Resist(1).ReducePct = 0.5) And (rs.Resist(1).ReduceFlatMin = 1) And (rs.Resist(1).ReduceFlatMax = 2) _
        And (rs.Resist(1).ReducePctMin = 5) And (rs.Resist(1).ReducePctMax = 15) And (rs.Resist(1).Immune = 1) _
        And (rs.Resist(1).ImmuneDamage = 1) And (rs.Resist(1).ImmuneEffect = 1) And (rs.Resist(1).ReduceEffectChancePct = 35) _
        And (rs.Resist(8).DamageType = 7) And (rs.Resist(2).DamageType = 0)
    Exit Function
e:
    test_parse_resist_fields_and_clamp = False
End Function

' Missing section and negative count both give Count = 0 (even over a stale count).
Private Function test_parse_resist_missing_and_negative() As Boolean
    On Error GoTo e
    Dim ini As New clsIniManager
    Dim a As t_ElementalResistSet, b As t_ElementalResistSet
    Call ini.Initialize(elem_write_temp_ini(elem_parse_ini_text()))
    a.Count = 5
    b.Count = 5
    Call ParseElementalResistFromIni(ini, "NO_SUCH_SECTION", a)
    Call ParseElementalResistFromIni(ini, "RESEMPTY", b)
    test_parse_resist_missing_and_negative = (a.Count = 0 And b.Count = 0)
    Exit Function
e:
    test_parse_resist_missing_and_negative = False
End Function

#End If
