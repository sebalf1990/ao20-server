Attribute VB_Name = "Unit_Venenos"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Poison Test Suite (plan 21.002, part B4)
' Characterization of the poison arithmetic that needs no network or map:
'   - PoisonHemoEffect: stack threshold / cap, timer refresh, partial cure,
'     expiry. Built on NPC references (no player I/O) with a huge tick time, so
'     no damage tick ever fires; stacks are read back through ReduceStacks(0).
'   - modPoisonResist adapters (effect resist read from the elemental engine).
'   - FindPoisonProfile lookup on a synthetic profile table.
' Not covered (the probe does it end to end): ApplyPoisonProfileToUser,
' NpcEnvenenarUser, TryPoisonNpcWithWeapon, ThrowArrowToTile, PerformTick.
' ValidatePoisonArrowData is Private to FileIO and only runs while loading
' obj.dat, so it cannot be called from a suite without touching production.
' ==========================================================================

Private SavedProfiles() As t_PoisonProfile
Private SavedProfileCount As Integer
Private mSavedObj1 As t_ObjData
Private mSavedObj2 As t_ObjData
Private mSavedSlots(1 To 7) As Integer
Private mSavedNpcNumero As Integer

Public Function test_suite_venenos() As Boolean
    Call UnitTesting.RunTest("test_hemo_initial_stacks_clamped", test_hemo_initial_stacks_clamped())
    Call UnitTesting.RunTest("test_hemo_hit_threshold_gates_stacks", test_hemo_hit_threshold_gates_stacks())
    Call UnitTesting.RunTest("test_hemo_nonpositive_threshold_means_every_hit", test_hemo_nonpositive_threshold_means_every_hit())
    Call UnitTesting.RunTest("test_hemo_cap_never_exceeded", test_hemo_cap_never_exceeded())
    Call UnitTesting.RunTest("test_hemo_new_stack_refreshes_timer", test_hemo_new_stack_refreshes_timer())
    Call UnitTesting.RunTest("test_hemo_hits_below_threshold_keep_timer", test_hemo_hits_below_threshold_keep_timer())
    Call UnitTesting.RunTest("test_hemo_cap_refresh_flag_on", test_hemo_cap_refresh_flag_on())
    Call UnitTesting.RunTest("test_hemo_cap_refresh_flag_off", test_hemo_cap_refresh_flag_off())
    Call UnitTesting.RunTest("test_hemo_reduce_stacks_partial", test_hemo_reduce_stacks_partial())
    Call UnitTesting.RunTest("test_hemo_reduce_stacks_to_zero_marks_remove", test_hemo_reduce_stacks_to_zero_marks_remove())
    Call UnitTesting.RunTest("test_hemo_reset_acts_as_hit", test_hemo_reset_acts_as_hit())
    Call UnitTesting.RunTest("test_hemo_expires_when_duration_elapses", test_hemo_expires_when_duration_elapses())
    Call UnitTesting.RunTest("test_poison_resist_user_invalid_index_zero", test_poison_resist_user_invalid_index_zero())
    Call UnitTesting.RunTest("test_poison_resist_user_no_gear_zero", test_poison_resist_user_no_gear_zero())
    Call UnitTesting.RunTest("test_poison_resist_user_immunity_variants", test_poison_resist_user_immunity_variants())
    Call UnitTesting.RunTest("test_poison_resist_user_effect_chance_adds_up", test_poison_resist_user_effect_chance_adds_up())
    Call UnitTesting.RunTest("test_poison_resist_user_ignores_other_types", test_poison_resist_user_ignores_other_types())
    Call UnitTesting.RunTest("test_poison_resist_user_family_is_ignored", test_poison_resist_user_family_is_ignored())
    Call UnitTesting.RunTest("test_poison_resist_npc_invalid_index_zero", test_poison_resist_npc_invalid_index_zero())
    Call UnitTesting.RunTest("test_poison_resist_npc_immunity_and_chance", test_poison_resist_npc_immunity_and_chance())
    Call UnitTesting.RunTest("test_find_poison_profile_exact_and_case_insensitive", test_find_poison_profile_exact_and_case_insensitive())
    Call UnitTesting.RunTest("test_find_poison_profile_first_match_wins", test_find_poison_profile_first_match_wins())
    Call UnitTesting.RunTest("test_find_poison_profile_not_found_cases", test_find_poison_profile_not_found_cases())
    test_suite_venenos = True
End Function

' --- Hemotoxin fixtures ------------------------------------------------------

' Source and target are NPC 1. TickTime is huge so Update never fires a damage tick.
Private Function hemo_new(ByVal stacksMax As Integer, ByVal golpes As Integer, ByVal refresca As Byte, ByVal stacksInicial As Integer, ByVal duration As Long) As PoisonHemoEffect
    Dim eff As PoisonHemoEffect
    Set eff = New PoisonHemoEffect
    Call eff.Setup(1, e_ReferenceType.eNpc, 1, e_ReferenceType.eNpc, 1, 1, 100000000, duration, 0, 1, 1, 0, 1, 1, stacksMax, golpes, 0, refresca, 1, 1, stacksInicial)
    Set hemo_new = eff
End Function

' Current stacks (ReduceStacks(0) removes nothing and returns what is left).
Private Function hemo_stacks(ByVal eff As PoisonHemoEffect) As Integer
    hemo_stacks = eff.ReduceStacks(0)
End Function

Private Sub hemo_hits(ByVal eff As PoisonHemoEffect, ByVal n As Long)
    Dim i As Long
    For i = 1 To n
        Call eff.TryAddStack
    Next i
End Sub

Private Function hemo_iface(ByVal eff As PoisonHemoEffect) As IBaseEffectOverTime
    Dim ifc As IBaseEffectOverTime
    Set ifc = eff
    Set hemo_iface = ifc
End Function

' --- Hemotoxin ---------------------------------------------------------------------

' Initial stacks clamp into 1..StacksMax; a non-positive StacksMax becomes 1.
Private Function test_hemo_initial_stacks_clamped() As Boolean
    On Error GoTo e
    test_hemo_initial_stacks_clamped = (hemo_stacks(hemo_new(3, 1, 0, 2, 1000)) = 2) _
        And (hemo_stacks(hemo_new(3, 1, 0, 0, 1000)) = 1) _
        And (hemo_stacks(hemo_new(3, 1, 0, -4, 1000)) = 1) _
        And (hemo_stacks(hemo_new(3, 1, 0, 9, 1000)) = 3) _
        And (hemo_stacks(hemo_new(0, 1, 0, 5, 1000)) = 1) _
        And (hemo_stacks(hemo_new(-2, 1, 0, 5, 1000)) = 1)
    Exit Function
e:
    test_hemo_initial_stacks_clamped = False
End Function

' GolpesQueSumanStacks = 2: only every second hit adds a stack, up to StacksMax = 3 (peak recorded).
Private Function test_hemo_hit_threshold_gates_stacks() As Boolean
    On Error GoTo e
    Dim eff As PoisonHemoEffect, ok As Boolean
    Set eff = hemo_new(3, 2, 0, 1, 1000)
    Call eff.TryAddStack
    ok = (hemo_stacks(eff) = 1)
    Call eff.TryAddStack
    ok = ok And (hemo_stacks(eff) = 2)
    Call eff.TryAddStack
    ok = ok And (hemo_stacks(eff) = 2)
    Call eff.TryAddStack
    ok = ok And (hemo_stacks(eff) = 3)
    Call hemo_hits(eff, 4)
    ok = ok And (hemo_stacks(eff) = 3) And (eff.TelemetryPeakStacks = 3)
    test_hemo_hit_threshold_gates_stacks = ok
    Exit Function
e:
    test_hemo_hit_threshold_gates_stacks = False
End Function

' A threshold of 0 or less is treated as 1: every hit adds a stack.
Private Function test_hemo_nonpositive_threshold_means_every_hit() As Boolean
    On Error GoTo e
    Dim a As PoisonHemoEffect, b As PoisonHemoEffect, ok As Boolean
    Set a = hemo_new(5, 0, 0, 1, 1000)
    Set b = hemo_new(5, -3, 0, 1, 1000)
    Call hemo_hits(a, 2)
    Call hemo_hits(b, 2)
    test_hemo_nonpositive_threshold_means_every_hit = (hemo_stacks(a) = 3) And (hemo_stacks(b) = 3)
    Exit Function
e:
    test_hemo_nonpositive_threshold_means_every_hit = False
End Function

' However many hits arrive, stacks never exceed StacksMax.
Private Function test_hemo_cap_never_exceeded() As Boolean
    On Error GoTo e
    Dim eff As PoisonHemoEffect
    Set eff = hemo_new(4, 1, 0, 1, 1000)
    Call hemo_hits(eff, 50)
    test_hemo_cap_never_exceeded = (hemo_stacks(eff) = 4) And (eff.TelemetryPeakStacks = 4)
    Exit Function
e:
    test_hemo_cap_never_exceeded = False
End Function

' Adding a stack always refreshes the global timer: 400 ms served -> 0 after the new stack.
Private Function test_hemo_new_stack_refreshes_timer() As Boolean
    On Error GoTo e
    Dim eff As PoisonHemoEffect, ok As Boolean
    Set eff = hemo_new(3, 1, 0, 1, 1000)
    Call hemo_iface(eff).Update(400)
    ok = (eff.TelemetryServedMs = 400)
    Call eff.TryAddStack
    ok = ok And (hemo_stacks(eff) = 2) And (eff.TelemetryServedMs = 0)
    test_hemo_new_stack_refreshes_timer = ok
    Exit Function
e:
    test_hemo_new_stack_refreshes_timer = False
End Function

' Hits that do not reach the threshold touch neither stacks nor timer.
Private Function test_hemo_hits_below_threshold_keep_timer() As Boolean
    On Error GoTo e
    Dim eff As PoisonHemoEffect, ok As Boolean
    Set eff = hemo_new(3, 3, 1, 1, 1000)
    Call hemo_iface(eff).Update(400)
    Call hemo_hits(eff, 2)
    ok = (hemo_stacks(eff) = 1) And (eff.TelemetryServedMs = 400)
    Call eff.TryAddStack
    ok = ok And (hemo_stacks(eff) = 2) And (eff.TelemetryServedMs = 0)
    test_hemo_hits_below_threshold_keep_timer = ok
    Exit Function
e:
    test_hemo_hits_below_threshold_keep_timer = False
End Function

' At the cap with RefrescaTimerAlStackear <> 0 a qualifying hit refreshes the timer but adds no stack.
Private Function test_hemo_cap_refresh_flag_on() As Boolean
    On Error GoTo e
    Dim eff As PoisonHemoEffect, ok As Boolean
    Set eff = hemo_new(2, 1, 1, 2, 1000)
    Call hemo_iface(eff).Update(400)
    Call eff.TryAddStack
    ok = (hemo_stacks(eff) = 2) And (eff.TelemetryServedMs = 0)
    test_hemo_cap_refresh_flag_on = ok
    Exit Function
e:
    test_hemo_cap_refresh_flag_on = False
End Function

' At the cap with RefrescaTimerAlStackear = 0 nothing changes: same stacks, timer keeps running.
Private Function test_hemo_cap_refresh_flag_off() As Boolean
    On Error GoTo e
    Dim eff As PoisonHemoEffect, ok As Boolean
    Set eff = hemo_new(2, 1, 0, 2, 1000)
    Call hemo_iface(eff).Update(400)
    Call eff.TryAddStack
    ok = (hemo_stacks(eff) = 2) And (eff.TelemetryServedMs = 400)
    test_hemo_cap_refresh_flag_off = ok
    Exit Function
e:
    test_hemo_cap_refresh_flag_off = False
End Function

' A partial cure removes N stacks and returns what is left; the effect stays alive.
Private Function test_hemo_reduce_stacks_partial() As Boolean
    On Error GoTo e
    Dim eff As PoisonHemoEffect, ok As Boolean
    Set eff = hemo_new(5, 1, 0, 4, 1000)
    ok = (eff.ReduceStacks(1) = 3) And (Not hemo_iface(eff).RemoveMe)
    ok = ok And (eff.ReduceStacks(2) = 1) And (Not hemo_iface(eff).RemoveMe)
    test_hemo_reduce_stacks_partial = ok
    Exit Function
e:
    test_hemo_reduce_stacks_partial = False
End Function

' Curing every stack (or more than there are) returns 0 and flags the effect for removal.
Private Function test_hemo_reduce_stacks_to_zero_marks_remove() As Boolean
    On Error GoTo e
    Dim a As PoisonHemoEffect, b As PoisonHemoEffect, ok As Boolean
    Set a = hemo_new(5, 1, 0, 2, 1000)
    Set b = hemo_new(5, 1, 0, 1, 1000)
    ok = (a.ReduceStacks(2) = 0) And hemo_iface(a).RemoveMe
    ok = ok And (b.ReduceStacks(5) = 0) And hemo_iface(b).RemoveMe
    test_hemo_reduce_stacks_to_zero_marks_remove = ok
    Exit Function
e:
    test_hemo_reduce_stacks_to_zero_marks_remove = False
End Function

' Re-applying the same poison (IBaseEffectOverTime.Reset) is a hit: it returns True and can add a stack.
Private Function test_hemo_reset_acts_as_hit() As Boolean
    On Error GoTo e
    Dim eff As PoisonHemoEffect, ok As Boolean
    Set eff = hemo_new(3, 1, 0, 1, 1000)
    ok = hemo_iface(eff).Reset(0, e_ReferenceType.eNpc, 0)
    ok = ok And (hemo_stacks(eff) = 2)
    test_hemo_reset_acts_as_hit = ok
    Exit Function
e:
    test_hemo_reset_acts_as_hit = False
End Function

' The global timer: 999 of 1000 ms keeps the effect alive, the last millisecond removes it.
Private Function test_hemo_expires_when_duration_elapses() As Boolean
    On Error GoTo e
    Dim eff As PoisonHemoEffect, ok As Boolean
    Set eff = hemo_new(3, 1, 0, 1, 1000)
    Call hemo_iface(eff).Update(999)
    ok = (Not hemo_iface(eff).RemoveMe)
    Call hemo_iface(eff).Update(1)
    ok = ok And hemo_iface(eff).RemoveMe
    test_hemo_expires_when_duration_elapses = ok
    Exit Function
e:
    test_hemo_expires_when_duration_elapses = False
End Function

' --- modPoisonResist adapters ------------------------------------------------------
' Fixture: user 1 wears ObjData(1) as armor and ObjData(2) as helmet, every other slot empty.


Private Sub resist_save()
    mSavedObj1 = ObjData(1)
    mSavedObj2 = ObjData(2)
    With UserList(1).invent
        mSavedSlots(1) = .EquippedArmorObjIndex
        mSavedSlots(2) = .EquippedHelmetObjIndex
        mSavedSlots(3) = .EquippedShieldObjIndex
        mSavedSlots(4) = .EquippedRingAccesoryObjIndex
        mSavedSlots(5) = .EquippedAmuletAccesoryObjIndex
        mSavedSlots(6) = .EquippedBackpackObjIndex
        mSavedSlots(7) = .EquippedSaddleObjIndex
        .EquippedArmorObjIndex = 0
        .EquippedHelmetObjIndex = 0
        .EquippedShieldObjIndex = 0
        .EquippedRingAccesoryObjIndex = 0
        .EquippedAmuletAccesoryObjIndex = 0
        .EquippedBackpackObjIndex = 0
        .EquippedSaddleObjIndex = 0
    End With
    mSavedNpcNumero = NpcList(1).Numero
    ObjData(1).ElementalResist.Count = 0
    ObjData(2).ElementalResist.Count = 0
End Sub

Private Sub resist_restore()
    ObjData(1) = mSavedObj1
    ObjData(2) = mSavedObj2
    With UserList(1).invent
        .EquippedArmorObjIndex = mSavedSlots(1)
        .EquippedHelmetObjIndex = mSavedSlots(2)
        .EquippedShieldObjIndex = mSavedSlots(3)
        .EquippedRingAccesoryObjIndex = mSavedSlots(4)
        .EquippedAmuletAccesoryObjIndex = mSavedSlots(5)
        .EquippedBackpackObjIndex = mSavedSlots(6)
        .EquippedSaddleObjIndex = mSavedSlots(7)
    End With
    NpcList(1).Numero = mSavedNpcNumero
End Sub

Private Sub resist_set(ByRef rs As t_ElementalResistSet, ByVal dmgType As Long, ByVal immune As Byte, ByVal immDmg As Byte, ByVal immEff As Byte, ByVal effChance As Long)
    rs.Count = 1
    rs.Resist(1).DamageType = dmgType
    rs.Resist(1).Immune = immune
    rs.Resist(1).ImmuneDamage = immDmg
    rs.Resist(1).ImmuneEffect = immEff
    rs.Resist(1).ReduceEffectChancePct = effChance
End Sub

Private Function test_poison_resist_user_invalid_index_zero() As Boolean
    On Error GoTo e
    Dim r As t_PoisonResist
    r = GetUserPoisonResist(0, 1)
    test_poison_resist_user_invalid_index_zero = (r.Inmune = 0 And r.ChancePct = 0)
    r = GetUserPoisonResist(-1, 1)
    test_poison_resist_user_invalid_index_zero = test_poison_resist_user_invalid_index_zero And (r.Inmune = 0 And r.ChancePct = 0)
    Exit Function
e:
    test_poison_resist_user_invalid_index_zero = False
End Function

' Nothing equipped: no resist at all.
Private Function test_poison_resist_user_no_gear_zero() As Boolean
    On Error GoTo e
    Dim r As t_PoisonResist
    Call resist_save
    r = GetUserPoisonResist(1, 1)
    test_poison_resist_user_no_gear_zero = (r.Inmune = 0 And r.ChancePct = 0)
    Call resist_restore
    Exit Function
e:
    Call resist_restore
    test_poison_resist_user_no_gear_zero = False
End Function

' Immune and ImmuneEffect block the effect (Inmune = 1); ImmuneDamage alone only blocks tick damage, so Inmune = 0.
Private Function test_poison_resist_user_immunity_variants() As Boolean
    On Error GoTo e
    Dim ok As Boolean
    Call resist_save
    UserList(1).invent.EquippedArmorObjIndex = 1
    Call resist_set(ObjData(1).ElementalResist, eDmgPoison, 1, 0, 0, 0)
    ok = (GetUserPoisonResist(1, 1).Inmune = 1)
    Call resist_set(ObjData(1).ElementalResist, eDmgPoison, 0, 0, 1, 0)
    ok = ok And (GetUserPoisonResist(1, 1).Inmune = 1)
    Call resist_set(ObjData(1).ElementalResist, eDmgPoison, 0, 1, 0, 0)
    ok = ok And (GetUserPoisonResist(1, 1).Inmune = 0)
    Call resist_restore
    test_poison_resist_user_immunity_variants = ok
    Exit Function
e:
    Call resist_restore
    test_poison_resist_user_immunity_variants = False
End Function

' ReduceEffectChancePct adds up across worn items (armor 30 + helmet 40 = 70).
Private Function test_poison_resist_user_effect_chance_adds_up() As Boolean
    On Error GoTo e
    Dim ok As Boolean
    Call resist_save
    UserList(1).invent.EquippedArmorObjIndex = 1
    UserList(1).invent.EquippedHelmetObjIndex = 2
    Call resist_set(ObjData(1).ElementalResist, eDmgPoison, 0, 0, 0, 30)
    Call resist_set(ObjData(2).ElementalResist, eDmgPoison, 0, 0, 0, 40)
    ok = (GetUserPoisonResist(1, 1).ChancePct = 70)
    Call resist_restore
    test_poison_resist_user_effect_chance_adds_up = ok
    Exit Function
e:
    Call resist_restore
    test_poison_resist_user_effect_chance_adds_up = False
End Function

' Only the Poison damage type counts: a fire resist (even immune) does not protect against the poison effect.
Private Function test_poison_resist_user_ignores_other_types() As Boolean
    On Error GoTo e
    Dim ok As Boolean, r As t_PoisonResist
    Call resist_save
    UserList(1).invent.EquippedArmorObjIndex = 1
    Call resist_set(ObjData(1).ElementalResist, eDmgFire, 1, 1, 1, 50)
    r = GetUserPoisonResist(1, 1)
    ok = (r.Inmune = 0 And r.ChancePct = 0)
    Call resist_restore
    test_poison_resist_user_ignores_other_types = ok
    Exit Function
e:
    Call resist_restore
    test_poison_resist_user_ignores_other_types = False
End Function

' The three poison families share one resist type: the family argument does not change the result.
Private Function test_poison_resist_user_family_is_ignored() As Boolean
    On Error GoTo e
    Dim ok As Boolean, f As Long
    Call resist_save
    UserList(1).invent.EquippedArmorObjIndex = 1
    Call resist_set(ObjData(1).ElementalResist, eDmgPoison, 0, 0, 0, 25)
    ok = True
    For f = 0 To 3
        ok = ok And (GetUserPoisonResist(1, CByte(f)).ChancePct = 25)
    Next f
    Call resist_restore
    test_poison_resist_user_family_is_ignored = ok
    Exit Function
e:
    Call resist_restore
    test_poison_resist_user_family_is_ignored = False
End Function

Private Function test_poison_resist_npc_invalid_index_zero() As Boolean
    On Error GoTo e
    Dim r As t_PoisonResist
    r = GetNpcPoisonResist(0, 1)
    test_poison_resist_npc_invalid_index_zero = (r.Inmune = 0 And r.ChancePct = 0)
    r = GetNpcPoisonResist(-1, 1)
    test_poison_resist_npc_invalid_index_zero = test_poison_resist_npc_invalid_index_zero And (r.Inmune = 0 And r.ChancePct = 0)
    Exit Function
e:
    test_poison_resist_npc_invalid_index_zero = False
End Function

' NPC resist comes from NpcInfoCache(NpcList(i).Numero): Immune/ImmuneEffect -> Inmune, ReduceEffectChancePct -> ChancePct;
' an NPC without a template number (Numero 0) has no resist.
Private Function test_poison_resist_npc_immunity_and_chance() As Boolean
    On Error GoTo e
    Dim ok As Boolean, r As t_PoisonResist, savedSet As t_ElementalResistSet
    If Not NpcInfoCacheInitialized Then Exit Function
    If UBound(NpcInfoCache) < 1 Then Exit Function
    Call resist_save
    savedSet = NpcInfoCache(1).ElementalResist
    NpcList(1).Numero = 1
    Call resist_set(NpcInfoCache(1).ElementalResist, eDmgPoison, 0, 0, 1, 60)
    r = GetNpcPoisonResist(1, 2)
    ok = (r.Inmune = 1 And r.ChancePct = 60)
    Call resist_set(NpcInfoCache(1).ElementalResist, eDmgPoison, 0, 1, 0, 0)
    r = GetNpcPoisonResist(1, 2)
    ok = ok And (r.Inmune = 0 And r.ChancePct = 0)
    NpcList(1).Numero = 0
    r = GetNpcPoisonResist(1, 2)
    ok = ok And (r.Inmune = 0 And r.ChancePct = 0)
    NpcInfoCache(1).ElementalResist = savedSet
    Call resist_restore
    test_poison_resist_npc_immunity_and_chance = ok
    Exit Function
e:
    NpcInfoCache(1).ElementalResist = savedSet
    Call resist_restore
    test_poison_resist_npc_immunity_and_chance = False
End Function

' --- FindPoisonProfile (synthetic table: Alpha, Beta, alpha) ----------------------------

Private Sub profiles_install()
    Dim i As Long
    SavedProfileCount = PoisonProfileCount
    If SavedProfileCount > 0 Then
        ReDim SavedProfiles(1 To SavedProfileCount)
        For i = 1 To SavedProfileCount
            SavedProfiles(i) = PoisonProfiles(i)
        Next i
    End If
    ReDim PoisonProfiles(1 To 3)
    PoisonProfiles(1).nombre = "Alpha"
    PoisonProfiles(2).nombre = "Beta"
    PoisonProfiles(3).nombre = "alpha"
    PoisonProfileCount = 3
End Sub

Private Sub profiles_restore()
    Dim i As Long
    If SavedProfileCount > 0 Then
        ReDim PoisonProfiles(1 To SavedProfileCount)
        For i = 1 To SavedProfileCount
            PoisonProfiles(i) = SavedProfiles(i)
        Next i
    Else
        ReDim PoisonProfiles(0 To 0)
    End If
    PoisonProfileCount = SavedProfileCount
End Sub

' Lookup by name returns the 1-based index; the comparison ignores case.
Private Function test_find_poison_profile_exact_and_case_insensitive() As Boolean
    On Error GoTo e
    Dim ok As Boolean
    Call profiles_install
    ok = (FindPoisonProfile("Beta") = 2) And (FindPoisonProfile("beta") = 2) And (FindPoisonProfile("BETA") = 2)
    Call profiles_restore
    test_find_poison_profile_exact_and_case_insensitive = ok
    Exit Function
e:
    Call profiles_restore
    test_find_poison_profile_exact_and_case_insensitive = False
End Function

' With two names equal ignoring case, the first one in the table wins.
Private Function test_find_poison_profile_first_match_wins() As Boolean
    On Error GoTo e
    Dim ok As Boolean
    Call profiles_install
    ok = (FindPoisonProfile("ALPHA") = 1) And (FindPoisonProfile("alpha") = 1)
    Call profiles_restore
    test_find_poison_profile_first_match_wins = ok
    Exit Function
e:
    Call profiles_restore
    test_find_poison_profile_first_match_wins = False
End Function

' Empty name, unknown name, name with a stray space, and an empty table all give 0.
Private Function test_find_poison_profile_not_found_cases() As Boolean
    On Error GoTo e
    Dim ok As Boolean
    Call profiles_install
    ok = (FindPoisonProfile("") = 0) And (FindPoisonProfile("Gamma") = 0) And (FindPoisonProfile(" Alpha") = 0)
    PoisonProfileCount = 0
    ok = ok And (FindPoisonProfile("Alpha") = 0)
    Call profiles_restore
    test_find_poison_profile_not_found_cases = ok
    Exit Function
e:
    Call profiles_restore
    test_find_poison_profile_not_found_cases = False
End Function

#End If
