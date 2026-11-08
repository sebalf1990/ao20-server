Attribute VB_Name = "modSpellCooldowns"
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
' Spell cooldowns that survive a relog (plan 2026-11-08-reemplazo-hambre-sed-energia, E3).
'
' The cooldown of a spell is a tick stamp per spell SLOT (Counters.UserHechizosInterval).
' Before this module a relog kept whatever the last occupant of the UserIndex left in
' that array: the cooldown of a chained resurrection could be skipped by relogging.
' On logout the stamps of the spells still cooling down are stored in server memory by
' character name and spell id; on login they are put back (the slots of every other spell
' are cleared). Lost when the server restarts (no database column).
' ==========================================================================
Private m_Stamps As Dictionary
Public Const SPELL_COOLDOWN_MAX_AGE_MINUTES As Long = 60

Private Function StampKey(ByVal charName As String, ByVal spellId As Integer) As String
    StampKey = UCase$(charName) & "#" & CStr(spellId)
End Function

' Pure rule: is a stamp taken at stampTick still cooling down for a cooldown of cooldownSeconds?
Public Function SpellCooldown_IsActive(ByVal stampTick As Long, ByVal nowTick As Long, ByVal cooldownSeconds As Long) As Boolean
    If stampTick = 0 Then Exit Function
    If cooldownSeconds <= 0 Then Exit Function
    SpellCooldown_IsActive = (TicksElapsed(stampTick, nowTick) < CDbl(cooldownSeconds) * 1000#)
End Function

' Cooldown in seconds that applies to the caster. Spells flagged FixedCooldown ignore the
' elven wood reduction (each equipped elven wood item halves it otherwise).
Public Function SpellCooldown_EffectiveSeconds(ByVal baseSeconds As Long, ByVal fixedCooldown As Boolean, ByVal elvenWoodItems As Long) As Double
    SpellCooldown_EffectiveSeconds = baseSeconds
    If fixedCooldown Then Exit Function
    Dim i As Long
    For i = 1 To elvenWoodItems
        SpellCooldown_EffectiveSeconds = SpellCooldown_EffectiveSeconds / 2
    Next i
End Function

Public Sub SpellCooldown_Store(ByVal charName As String, ByVal spellId As Integer, ByVal stampTick As Long, ByVal savedAt As Date)
    If m_Stamps Is Nothing Then Set m_Stamps = New Dictionary
    Dim key As String
    key = StampKey(charName, spellId)
    If m_Stamps.Exists(key) Then m_Stamps.Remove key
    Call m_Stamps.Add(key, CStr(stampTick) & "|" & Trim$(Str$(CDbl(savedAt))))
End Sub

' Returns the stored stamp and forgets it (0 when there is none, or when it was saved more
' than SPELL_COOLDOWN_MAX_AGE_MINUTES ago: a long break counts as rest).
Public Function SpellCooldown_Take(ByVal charName As String, ByVal spellId As Integer, ByVal nowDate As Date) As Long
    On Error GoTo SpellCooldown_Take_Err
    If m_Stamps Is Nothing Then Exit Function
    Dim key   As String
    Dim parts() As String
    key = StampKey(charName, spellId)
    If Not m_Stamps.Exists(key) Then Exit Function
    parts = Split(CStr(m_Stamps.Item(key)), "|")
    m_Stamps.Remove key
    If UBound(parts) <> 1 Then Exit Function
    If (CDbl(nowDate) - val(parts(1))) * 1440# > SPELL_COOLDOWN_MAX_AGE_MINUTES Then Exit Function
    SpellCooldown_Take = CLng(parts(0))
    Exit Function
SpellCooldown_Take_Err:
    SpellCooldown_Take = 0
End Function

Public Sub SpellCooldown_OnLogout(ByVal UserIndex As Integer)
    On Error GoTo SpellCooldown_OnLogout_Err
    Dim slot    As Long
    Dim spellId As Integer
    Dim nowTick As Long
    If LenB(UserList(UserIndex).name) = 0 Then Exit Sub
    nowTick = GetTickCountRaw()
    With UserList(UserIndex)
        For slot = 1 To MAXUSERHECHIZOS
            spellId = .Stats.UserHechizos(slot)
            If spellId > 0 Then
                If Hechizos(spellId).Cooldown > 0 Then
                    If SpellCooldown_IsActive(.Counters.UserHechizosInterval(slot), nowTick, Hechizos(spellId).Cooldown) Then
                        Call SpellCooldown_Store(.name, spellId, .Counters.UserHechizosInterval(slot), Now)
                    End If
                End If
            End If
        Next slot
    End With
    Exit Sub
SpellCooldown_OnLogout_Err:
    Call TraceError(Err.Number, Err.Description, "modSpellCooldowns.SpellCooldown_OnLogout", Erl)
End Sub

' Must run after the spells of the character are loaded.
Public Sub SpellCooldown_OnLogin(ByVal UserIndex As Integer)
    On Error GoTo SpellCooldown_OnLogin_Err
    Dim slot    As Long
    Dim spellId As Integer
    Dim stamp   As Long
    Dim nowTick As Long
    nowTick = GetTickCountRaw()
    With UserList(UserIndex)
        For slot = 1 To MAXUSERHECHIZOS
            .Counters.UserHechizosInterval(slot) = 0
            spellId = .Stats.UserHechizos(slot)
            If spellId > 0 Then
                stamp = SpellCooldown_Take(.name, spellId, Now)
                If stamp <> 0 Then
                    If SpellCooldown_IsActive(stamp, nowTick, Hechizos(spellId).Cooldown) Then
                        .Counters.UserHechizosInterval(slot) = stamp
                    End If
                End If
            End If
        Next slot
    End With
    Exit Sub
SpellCooldown_OnLogin_Err:
    Call TraceError(Err.Number, Err.Description, "modSpellCooldowns.SpellCooldown_OnLogin", Erl)
End Sub
