Attribute VB_Name = "Unit_ClanExp"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Clan Experience Test Suite (plan 21.002, part B3)
' Characterization of the pure clan-level helpers in modGuilds, from the traps
' documented in docs/experiencia-de-clan.md:
'   1. no intermediate level may require 0 (0 is the client's "max level" sentinel)
'   2. the table is indexed by the CURRENT level, not the destination level
'   3. thresholds above 32767 must not overflow (GuildExpLevel5 was 35000 once)
' CheckClanExp itself is NOT covered here: it needs a live clan (clsClan reads
' and writes database.db on every call) and a fixture would have to touch the
' database. It stays with the probe (fundar_clan_oro and clan experience).
' ==========================================================================

Public Function test_suite_clanexp() As Boolean
    Call UnitTesting.RunTest("test_clan_required_exp_max_level_is_zero", test_clan_required_exp_max_level_is_zero())
    Call UnitTesting.RunTest("test_clan_required_exp_intermediate_levels_never_zero", test_clan_required_exp_intermediate_levels_never_zero())
    Call UnitTesting.RunTest("test_clan_required_exp_index_is_current_level", test_clan_required_exp_index_is_current_level())
    Call UnitTesting.RunTest("test_clan_required_exp_invalid_levels_zero", test_clan_required_exp_invalid_levels_zero())
    Call UnitTesting.RunTest("test_clan_required_exp_above_integer_range", test_clan_required_exp_above_integer_range())
    Call UnitTesting.RunTest("test_clan_required_exp_loaded_thresholds_increase", test_clan_required_exp_loaded_thresholds_increase())
    Call UnitTesting.RunTest("test_clan_members_by_level_loaded_positive", test_clan_members_by_level_loaded_positive())
    Call UnitTesting.RunTest("test_clan_miembros_permite_invalid_guild_zero", test_clan_miembros_permite_invalid_guild_zero())
    test_suite_clanexp = True
End Function

' The maximum level (and anything above it) reports 0: the "no more levels" sentinel of the client.
Private Function test_clan_required_exp_max_level_is_zero() As Boolean
    On Error GoTo e
    test_clan_required_exp_max_level_is_zero = (GetRequiredExpForGuildLevel(MAX_LEVEL_GUILD) = 0) _
        And (GetRequiredExpForGuildLevel(MAX_LEVEL_GUILD + 1) = 0) _
        And (GetRequiredExpForGuildLevel(255) = 0) _
        And (GetRequiredExpForGuildLevel(32767) = 0)
    Exit Function
e:
    test_clan_required_exp_max_level_is_zero = False
End Function

' Trap 1: levels 1..MAX-1 must each need a positive amount (a 0 would make the client think the clan is maxed out)
' and return the loaded table entry.
Private Function test_clan_required_exp_intermediate_levels_never_zero() As Boolean
    On Error GoTo e
    Dim lvl As Long, ok As Boolean
    ok = True
    For lvl = 1 To MAX_LEVEL_GUILD - 1
        ok = ok And (GetRequiredExpForGuildLevel(lvl) > 0) And (GetRequiredExpForGuildLevel(lvl) = ExpLevelUpGuild(lvl))
    Next lvl
    test_clan_required_exp_intermediate_levels_never_zero = ok
    Exit Function
e:
    test_clan_required_exp_intermediate_levels_never_zero = False
End Function

' Trap 2: the argument is the clan's current level, and the value is the entry of THAT level (not level+1).
Private Function test_clan_required_exp_index_is_current_level() As Boolean
    On Error GoTo e
    Dim s1 As Long, s2 As Long, ok As Boolean
    s1 = ExpLevelUpGuild(1)
    s2 = ExpLevelUpGuild(2)
    ExpLevelUpGuild(1) = 111
    ExpLevelUpGuild(2) = 222
    ok = (GetRequiredExpForGuildLevel(1) = 111) And (GetRequiredExpForGuildLevel(2) = 222)
    ExpLevelUpGuild(1) = s1
    ExpLevelUpGuild(2) = s2
    test_clan_required_exp_index_is_current_level = ok
    Exit Function
e:
    ExpLevelUpGuild(1) = s1
    ExpLevelUpGuild(2) = s2
    test_clan_required_exp_index_is_current_level = False
End Function

' Level 0 and negative levels are outside the table: 0, no array overrun.
Private Function test_clan_required_exp_invalid_levels_zero() As Boolean
    On Error GoTo e
    test_clan_required_exp_invalid_levels_zero = (GetRequiredExpForGuildLevel(0) = 0) _
        And (GetRequiredExpForGuildLevel(-1) = 0) And (GetRequiredExpForGuildLevel(-32768) = 0)
    Exit Function
e:
    test_clan_required_exp_invalid_levels_zero = False
End Function

' Trap 3: thresholds above 32767 come back intact (Long all the way): the historical 35000 and 200000.
Private Function test_clan_required_exp_above_integer_range() As Boolean
    On Error GoTo e
    Dim s5 As Long, s6 As Long, ok As Boolean
    s5 = ExpLevelUpGuild(5)
    s6 = ExpLevelUpGuild(6)
    ExpLevelUpGuild(5) = 35000
    ExpLevelUpGuild(6) = 200000
    ok = (GetRequiredExpForGuildLevel(5) = 35000) And (GetRequiredExpForGuildLevel(6) = 200000)
    ExpLevelUpGuild(5) = s5
    ExpLevelUpGuild(6) = s6
    test_clan_required_exp_above_integer_range = ok
    Exit Function
e:
    ExpLevelUpGuild(5) = s5
    ExpLevelUpGuild(6) = s6
    test_clan_required_exp_above_integer_range = False
End Function

' Data check on Clanes.dat as loaded: every level asks for more experience than the previous one.
Private Function test_clan_required_exp_loaded_thresholds_increase() As Boolean
    On Error GoTo e
    Dim lvl As Long, ok As Boolean
    ok = True
    For lvl = 2 To MAX_LEVEL_GUILD - 1
        ok = ok And (GetRequiredExpForGuildLevel(lvl) > GetRequiredExpForGuildLevel(lvl - 1))
    Next lvl
    test_clan_required_exp_loaded_thresholds_increase = ok
    Exit Function
e:
    test_clan_required_exp_loaded_thresholds_increase = False
End Function

' Data check on Clanes.dat as loaded: every level admits at least one member.
Private Function test_clan_members_by_level_loaded_positive() As Boolean
    On Error GoTo e
    Dim lvl As Long, ok As Boolean
    ok = True
    For lvl = 1 To MAX_LEVEL_GUILD
        ok = ok And (MembersByLevel(lvl) > 0)
    Next lvl
    test_clan_members_by_level_loaded_positive = ok
    Exit Function
e:
    test_clan_members_by_level_loaded_positive = False
End Function

' A guild index that does not exist (0, negative, past the loaded count) allows 0 members, never an error.
Private Function test_clan_miembros_permite_invalid_guild_zero() As Boolean
    On Error GoTo e
    test_clan_miembros_permite_invalid_guild_zero = (MiembrosPermite(0) = 0) And (MiembrosPermite(-1) = 0) _
        And (MiembrosPermite(CANTIDADDECLANES + 1) = 0) And (MiembrosPermite(32767) = 0)
    Exit Function
e:
    test_clan_miembros_permite_invalid_guild_zero = False
End Function

#End If
