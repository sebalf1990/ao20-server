Attribute VB_Name = "EffectsOverTime"
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
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
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit
Private LastUpdateTime        As Long
Private UniqueIdCounter       As Long

Private ActiveEffects         As t_EffectOverTimeList
Const UnequipEffectId = 23
Const INITIAL_POOL_SIZE = 200
Private EffectPools() As t_EffectOverTimeList

Public Enum e_EffectCallbackMask
    eTargetUseMagic = 1
    eTargetWillAttack = 2
    eTargetDidHit = 4
    eTargetFailedAttack = 8
    eTargetWasDamaged = 16
    eTargetWillAttackPosition = 32
    eTargetApplyDamageReduction = 64
    eTargetChangeTerrain = 128
End Enum

Public Sub InitializePools()
    On Error GoTo InitializePools_Err
    Dim i           As Integer
    Dim j           As Integer
    Dim InitialSize As Integer
    If RunningInVB() Then
        InitialSize = 2
    Else
        InitialSize = INITIAL_POOL_SIZE
    End If
    ReDim EffectPools(1 To e_EffectOverTimeType.EffectTypeCount - 1) As t_EffectOverTimeList
    For i = 1 To e_EffectOverTimeType.EffectTypeCount - 1
        ReDim EffectPools(i).EffectList(InitialSize) As IBaseEffectOverTime
        For j = 0 To InitialSize
            Call AddEffect(EffectPools(i), InstantiateEOT(i))
        Next j
    Next i
    LastUpdateTime = GetTickCountRaw()
    Exit Sub
InitializePools_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.InitializePools", Erl)
End Sub


Public Sub UpdateEffectOverTime()
    On Error GoTo Update_Err

    Dim CurrTime As Long
    Dim ElapsedTime As Long

    CurrTime = GetTickCountRaw()

    ' First frame (or after reset): prime and bail to avoid huge delta
    If LastUpdateTime = 0 Then
        LastUpdateTime = CurrTime
        Exit Sub
    End If

    ' Wrap-safe delta; guard against negative/insane values
    ElapsedTime = CLng(TicksElapsed(LastUpdateTime, CurrTime))
    If ElapsedTime < 0 Then ElapsedTime = 0

    ' Clamp to avoid massive catch-up after a hitch (e.g., 2s+ pause)
    ' 200 ms is a good default for server-side EOT; tune if needed.
    If ElapsedTime > 200 Then ElapsedTime = 200

    LastUpdateTime = CurrTime

    Dim i As Integer
    Do While i < ActiveEffects.EffectCount
        If UpdateEffect(i, ElapsedTime) Then
            i = i + 1
        End If
    Loop
    Exit Sub
Update_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.Update", Erl)
End Sub

Private Function UpdateEffect(ByVal Index As Integer, ByVal ElapsedTime As Long) As Boolean
    On Error GoTo UpdateEffect_Err
    'this should never happend but it covers us for breaking all effects if something goes wrong
    If ActiveEffects.EffectList(Index) Is Nothing Then
        UpdateEffect = True
        Exit Function
    End If
    Dim CurrentEffect As IBaseEffectOverTime
    Set CurrentEffect = ActiveEffects.EffectList(Index)
    ' 06.002 Ola 1: si ya esta marcado para remover (p.ej. ClearEffectList por una muerte en esta
    ' misma pasada), no ejecutar su Update: el tick golpearia a un target ya muerto.
    If Not CurrentEffect.RemoveMe Then CurrentEffect.Update (ElapsedTime)
    If CurrentEffect.RemoveMe Then
        If CurrentEffect.TargetIsValid Then
            If CurrentEffect.TargetRefType = eUser Then
                Call RemoveEffect(UserList(CurrentEffect.TargetArrayIndex).EffectOverTime, CurrentEffect)
            ElseIf CurrentEffect.TargetRefType = eNpc Then
                Call RemoveEffect(NpcList(CurrentEffect.TargetArrayIndex).EffectOverTime, CurrentEffect)
            End If
        End If
        Call RemoveEffectAtPos(ActiveEffects, Index)
        Call RecycleEffect(CurrentEffect)
        UpdateEffect = False
    Else
        UpdateEffect = True
    End If
    Exit Function
UpdateEffect_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.UpdateEffect", Erl)
    Set ActiveEffects.EffectList(Index) = Nothing
    UpdateEffect = True
End Function

Private Function GetNextId() As Long
    UniqueIdCounter = (UniqueIdCounter + 1) And &H7FFFFFFF
    GetNextId = UniqueIdCounter
End Function

Public Sub CreateEffect(ByVal SourceIndex As Integer, _
                        ByVal SourceType As e_ReferenceType, _
                        ByVal TargetIndex As Integer, _
                        ByVal TargetType As e_ReferenceType, _
                        ByVal EffectIndex As Integer)
    On Error GoTo CreateEffect_Err
    
    If EffectIndex < LBound(EffectOverTime) Or EffectIndex > UBound(EffectOverTime) Then
        Call TraceError(9, "Invalid EffectIndex=" & EffectIndex, "EffectsOverTime.CreateEffect", Erl)
        Exit Sub
    End If

    Dim EffectType As e_EffectOverTimeType
    EffectType = EffectOverTime(EffectIndex).Type
    Select Case EffectType
        Case e_EffectOverTimeType.eHealthModifier
            Dim Dot As UpdateHpOverTime
            Set Dot = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call Dot.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(Dot)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, Dot)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, Dot)
            End If
        Case e_EffectOverTimeType.eApplyModifiers
            Dim StatDot As StatModifier
            Set StatDot = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call StatDot.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(StatDot)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, StatDot)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, StatDot)
            End If
        Case e_EffectOverTimeType.eProvoke
            Dim Provoke As EffectProvoke
            Set Provoke = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call Provoke.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(Provoke)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, Provoke)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, Provoke)
            End If
        Case e_EffectOverTimeType.eProvoked
            Dim StatProvoked As EffectProvoked
            Set StatProvoked = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call StatProvoked.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(StatProvoked)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, StatProvoked)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, StatProvoked)
            End If
        Case e_EffectOverTimeType.eDrunk
            Dim Drunk As DrunkEffect
            Set Drunk = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call Drunk.Setup(SourceIndex, SourceType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(Drunk)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, Drunk)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, Drunk)
            End If
        Case e_EffectOverTimeType.eTranslation
            Dim TE As TranslationEffect
            Set TE = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call TE.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(TE)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, TE)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, TE)
            End If
        Case e_EffectOverTimeType.eApplyEffectOnHit
            Dim EOH As ApplyEffectOnHit
            Set EOH = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call EOH.Setup(SourceIndex, SourceType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(EOH)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, EOH)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, EOH)
            End If
        Case e_EffectOverTimeType.eManaModifier
            Dim Mot As UpdateManaOverTime
            Set Mot = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call Mot.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(Mot)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, Mot)
                'npc doesn't have mana
            End If
        Case e_EffectOverTimeType.ePartyBonus
            Dim PartyEffect As ApplyEffectToParty
            Set PartyEffect = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call PartyEffect.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(PartyEffect)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, PartyEffect)
                'npc doesn't have groups
            End If
        Case e_EffectOverTimeType.ePullTarget
            Dim PullEffect As AttrackEffect
            Set PullEffect = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call PullEffect.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(PullEffect)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, PullEffect)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, PullEffect)
            End If
        Case e_EffectOverTimeType.eMultipleAttacks
            Dim MultiAttacks As MultipleAttacks
            Set MultiAttacks = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call MultiAttacks.Setup(SourceIndex, SourceType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(MultiAttacks)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, MultiAttacks)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, MultiAttacks)
            End If
        Case e_EffectOverTimeType.eProtection
            Dim Protect As ProtectEffect
            Set Protect = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call Protect.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(Protect)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, Protect)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, Protect)
            End If
        Case e_EffectOverTimeType.eTransform
            Dim Transform As TransformEffect
            Set Transform = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call Transform.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(Transform)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, Transform)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, Transform)
            End If
        Case e_EffectOverTimeType.eBonusDamage
            Dim BonusDamage As BonusDamageEffect
            Set BonusDamage = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call BonusDamage.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(BonusDamage)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, BonusDamage)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, BonusDamage)
            End If
        Case e_EffectOverTimeType.eMinimapRadar
            Dim Radar As EffectMinimapRadar
            Set Radar = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call Radar.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(Radar)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, Radar)
            End If
        Case e_EffectOverTimeType.eMinimapUserDetect
            Dim UserDetect As EffectMinimapUDetect
            Set UserDetect = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call UserDetect.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(UserDetect)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, UserDetect)
            End If
        Case Else
            Debug.Assert False
    End Select
    Exit Sub
CreateEffect_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.CreateEffect EffectIndex:" & EffectIndex, Erl)
End Sub

Public Sub CreateTrap(ByVal SourceIndex As Integer, _
                      ByVal SourceType As e_ReferenceType, _
                      ByVal Map As Integer, _
                      ByVal TileX As Integer, _
                      ByVal TileY As Integer, _
                      ByVal EffectTypeId As Integer)
    On Error GoTo CreateTrap_Err
    Dim EffectType As e_EffectOverTimeType
    EffectType = e_EffectOverTimeType.eTrap
    Dim Trap As clsTrap
    Set Trap = GetEOT(EffectType)
    UniqueIdCounter = GetNextId()
    Call Trap.Setup(SourceIndex, SourceType, EffectTypeId, UniqueIdCounter, Map, TileX, TileY)
    Call AddEffectToUpdate(Trap)
    If SourceType = eUser Then
        Call AddEffect(UserList(SourceIndex).EffectOverTime, Trap)
    ElseIf SourceType = eNpc Then
        Call AddEffect(NpcList(SourceIndex).EffectOverTime, Trap)
    End If
    Exit Sub
CreateTrap_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.CreateTrap", Erl)
End Sub

Public Sub CreateDelayedBlast(ByVal SourceIndex As Integer, _
                              ByVal SourceType As e_ReferenceType, _
                              ByVal Map As Integer, _
                              ByVal TileX As Integer, _
                              ByVal TileY As Integer, _
                              ByVal EffectTypeId As Integer, _
                              ByVal SourceObjIndex As Integer)
    On Error GoTo CreateDelayedBlast_Err
    Dim EffectType As e_EffectOverTimeType
    EffectType = e_EffectOverTimeType.eDelayedBlast
    Dim Blast As DelayedBlast
    Set Blast = GetEOT(EffectType)
    UniqueIdCounter = GetNextId()
    Call Blast.Setup(SourceIndex, SourceType, EffectTypeId, UniqueIdCounter, Map, TileX, TileY, SourceObjIndex)
    Call AddEffectToUpdate(Blast)
    If SourceType = eUser Then
        Call AddEffect(UserList(SourceIndex).EffectOverTime, Blast)
    ElseIf SourceType = eNpc Then
        Call AddEffect(NpcList(SourceIndex).EffectOverTime, Blast)
    End If
    Exit Sub
CreateDelayedBlast_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.CreateTrap", Erl)
End Sub

Public Sub CreateUnequip(ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, ByVal ItemSlotType As Long)
    On Error GoTo CreateDelayedBlast_Err
    If Not IsFeatureEnabled("bandit_unequip_bonus") Then Exit Sub
    Dim EffectType As e_EffectOverTimeType
    EffectType = e_EffectOverTimeType.eUnequip
    Dim Unequip As UnequipItem
    Set Unequip = GetEOT(EffectType)
    UniqueIdCounter = GetNextId()
    Call Unequip.Setup(TargetIndex, TargetType, UnequipEffectId, UniqueIdCounter, ItemSlotType)
    Call AddEffectToUpdate(Unequip)
    If TargetType = eUser Then
        Call AddEffect(UserList(TargetIndex).EffectOverTime, Unequip)
    End If
    Exit Sub
CreateDelayedBlast_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.CreateTrap", Erl)
End Sub

' === Sistema de venenos nuevo (TOGGLE26 new_poison_system) ===
' Helper que crea un PoisonMinorEffect con parametros desde la fuente (arma, hechizo, perfil NPC).
' EffectId apunta a uno de los presets en EffectsOverTime.dat (EOT63 para Menor por default).
Public Sub CreatePoisonMinor(ByVal SourceIndex As Integer, _
                             ByVal SourceType As e_ReferenceType, _
                             ByVal TargetIndex As Integer, _
                             ByVal TargetType As e_ReferenceType, _
                             ByVal EffectId As Integer, _
                             ByVal TickTime As Long, _
                             ByVal Duration As Long, _
                             ByVal DanoModo As Byte, _
                             ByVal DanoMin As Long, _
                             ByVal DanoMax As Long, _
                             ByVal FactorPvP As Single, _
                             ByVal FactorPvE As Single)
    On Error GoTo CreatePoisonMinor_Err
    Dim Existing As IBaseEffectOverTime
    Dim eotIdRef As Integer
    eotIdRef = EffectId
    ' Buscar si ya existe un Veneno Menor en el target. Si si: Reset (refresh timer).
    If TargetType = eUser Then
        Set Existing = FindEffectOnTarget(TargetIndex, UserList(TargetIndex).EffectOverTime, eotIdRef)
    ElseIf TargetType = eNpc Then
        Set Existing = FindEffectOnTarget(TargetIndex, NpcList(TargetIndex).EffectOverTime, eotIdRef)
    End If
    If Not (Existing Is Nothing) Then
        Call Existing.Reset(SourceIndex, SourceType, eotIdRef)
        Exit Sub
    End If
    Dim Effect As PoisonMinorEffect
    Set Effect = GetEOT(e_EffectOverTimeType.ePoisonMinor)
    UniqueIdCounter = GetNextId()
    Call Effect.Setup(SourceIndex, SourceType, TargetIndex, TargetType, eotIdRef, UniqueIdCounter, _
                       TickTime, Duration, DanoModo, DanoMin, DanoMax, FactorPvP, FactorPvE)
    Call AddEffectToUpdate(Effect)
    If TargetType = eUser Then
        Call AddEffect(UserList(TargetIndex).EffectOverTime, Effect)
        Call LogPoisonEvent("apply", "", UserList(TargetIndex).name, eotIdRef, 1, 0, 0, 0, 0)
    ElseIf TargetType = eNpc Then
        Call AddEffect(NpcList(TargetIndex).EffectOverTime, Effect)
    End If
    Exit Sub
CreatePoisonMinor_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.CreatePoisonMinor", Erl)
End Sub

' Limpia explicitamente Veneno Menor del target (cura/muerte). Pone flag.PoisonMinorActive=0.
Public Sub RemovePoisonMinor(ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType)
    On Error GoTo RemovePoisonMinor_Err
    Dim ListRef As Integer
    Dim i As Integer
    If TargetType = eUser Then
        UserList(TargetIndex).flags.PoisonMinorActive = 0
        Do While i < UserList(TargetIndex).EffectOverTime.EffectCount
            If UserList(TargetIndex).EffectOverTime.EffectList(i).TypeId = e_EffectOverTimeType.ePoisonMinor Then
                ' Grupo 7 (plan 10.001, punto 7b): effect_end reason=curado (motivo 1).
                If modElementalBalanceLog.ElementalPlayerTelemetryEnabled() Then
                    Dim ebMinorCured As PoisonMinorEffect
                    Set ebMinorCured = UserList(TargetIndex).EffectOverTime.EffectList(i)
                    Call modElementalBalanceLog.LogPoisonEffectEnd(TargetIndex, UserList(TargetIndex).EffectOverTime.EffectList(i).CasterArrayIndex, UserList(TargetIndex).EffectOverTime.EffectList(i).CasterRefType = eNpc, UserList(TargetIndex).EffectOverTime.EffectList(i).CasterIsValid, UserList(TargetIndex).EffectOverTime.EffectList(i).EotId, ebMinorCured.TelemetryServedMs, 0, 1)
                End If
                UserList(TargetIndex).EffectOverTime.EffectList(i).RemoveMe = True
                Call RemoveEffectAtPos(UserList(TargetIndex).EffectOverTime, i)
            Else
                i = i + 1
            End If
        Loop
    ElseIf TargetType = eNpc Then
        Do While i < NpcList(TargetIndex).EffectOverTime.EffectCount
            If NpcList(TargetIndex).EffectOverTime.EffectList(i).TypeId = e_EffectOverTimeType.ePoisonMinor Then
                NpcList(TargetIndex).EffectOverTime.EffectList(i).RemoveMe = True
                Call RemoveEffectAtPos(NpcList(TargetIndex).EffectOverTime, i)
            Else
                i = i + 1
            End If
        Loop
    End If
    Exit Sub
RemovePoisonMinor_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.RemovePoisonMinor", Erl)
End Sub

' === Hemotoxina (Fase 3) ===
' Crea o stackea un PoisonHemoEffect en el target.
' Si ya existe, llama Reset (que internamente hace TryAddStack -> +1 stack y refresca timer).
Public Sub CreatePoisonHemo(ByVal SourceIndex As Integer, _
                            ByVal SourceType As e_ReferenceType, _
                            ByVal TargetIndex As Integer, _
                            ByVal TargetType As e_ReferenceType, _
                            ByVal EffectId As Integer, _
                            ByVal TickTime As Long, _
                            ByVal Duration As Long, _
                            ByVal DanoModo As Byte, _
                            ByVal DanoMin As Long, _
                            ByVal DanoMax As Long, _
                            ByVal DanoPorStackModo As Byte, _
                            ByVal DanoPorStackMin As Long, _
                            ByVal DanoPorStackMax As Long, _
                            ByVal StacksMax As Integer, _
                            ByVal GolpesQueSumanStacks As Integer, _
                            ByVal IntervaloDecayStackMs As Long, _
                            ByVal RefrescaTimerAlStackear As Byte, _
                            ByVal FactorPvP As Single, _
                            ByVal FactorPvE As Single, _
                            ByVal StacksInicial As Integer)
    On Error GoTo CreatePoisonHemo_Err
    Dim Existing As IBaseEffectOverTime
    Dim eotIdRef As Integer
    eotIdRef = EffectId
    If TargetType = eUser Then
        Set Existing = FindEffectOnTarget(TargetIndex, UserList(TargetIndex).EffectOverTime, eotIdRef)
    ElseIf TargetType = eNpc Then
        Set Existing = FindEffectOnTarget(TargetIndex, NpcList(TargetIndex).EffectOverTime, eotIdRef)
    End If
    If Not (Existing Is Nothing) Then
        ' Reset Hemo = +1 stack si no esta en cap
        Call Existing.Reset(SourceIndex, SourceType, eotIdRef)
        Exit Sub
    End If
    Dim Effect As PoisonHemoEffect
    Set Effect = GetEOT(e_EffectOverTimeType.ePoisonHemo)
    UniqueIdCounter = GetNextId()
    Call Effect.Setup(SourceIndex, SourceType, TargetIndex, TargetType, eotIdRef, UniqueIdCounter, _
                       TickTime, Duration, DanoModo, DanoMin, DanoMax, _
                       DanoPorStackModo, DanoPorStackMin, DanoPorStackMax, _
                       StacksMax, GolpesQueSumanStacks, IntervaloDecayStackMs, RefrescaTimerAlStackear, _
                       FactorPvP, FactorPvE, StacksInicial)
    Call AddEffectToUpdate(Effect)
    If TargetType = eUser Then
        Call AddEffect(UserList(TargetIndex).EffectOverTime, Effect)
        Call LogPoisonEvent("apply", "", UserList(TargetIndex).name, eotIdRef, 2, 0, StacksInicial, 0, 0)
    ElseIf TargetType = eNpc Then
        Call AddEffect(NpcList(TargetIndex).EffectOverTime, Effect)
    End If
    Exit Sub
CreatePoisonHemo_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.CreatePoisonHemo", Erl)
End Sub

Public Sub RemovePoisonHemo(ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType)
    On Error GoTo RemovePoisonHemo_Err
    Dim i As Integer
    If TargetType = eUser Then
        UserList(TargetIndex).flags.PoisonHemoStacks = 0
        Do While i < UserList(TargetIndex).EffectOverTime.EffectCount
            If UserList(TargetIndex).EffectOverTime.EffectList(i).TypeId = e_EffectOverTimeType.ePoisonHemo Then
                ' Grupo 7 (plan 10.001, punto 7b): effect_end reason=curado (motivo 1).
                If modElementalBalanceLog.ElementalPlayerTelemetryEnabled() Then
                    Dim ebHemoCured As PoisonHemoEffect
                    Set ebHemoCured = UserList(TargetIndex).EffectOverTime.EffectList(i)
                    Call modElementalBalanceLog.LogPoisonEffectEnd(TargetIndex, UserList(TargetIndex).EffectOverTime.EffectList(i).CasterArrayIndex, UserList(TargetIndex).EffectOverTime.EffectList(i).CasterRefType = eNpc, UserList(TargetIndex).EffectOverTime.EffectList(i).CasterIsValid, UserList(TargetIndex).EffectOverTime.EffectList(i).EotId, ebHemoCured.TelemetryServedMs, ebHemoCured.TelemetryPeakStacks, 1)
                End If
                UserList(TargetIndex).EffectOverTime.EffectList(i).RemoveMe = True
                Call RemoveEffectAtPos(UserList(TargetIndex).EffectOverTime, i)
            Else
                i = i + 1
            End If
        Loop
        ' Apaga contador de stacks en cliente
        Call WriteUpdatePoisonStacks(TargetIndex, 0, 0)
    ElseIf TargetType = eNpc Then
        Do While i < NpcList(TargetIndex).EffectOverTime.EffectCount
            If NpcList(TargetIndex).EffectOverTime.EffectList(i).TypeId = e_EffectOverTimeType.ePoisonHemo Then
                NpcList(TargetIndex).EffectOverTime.EffectList(i).RemoveMe = True
                Call RemoveEffectAtPos(NpcList(TargetIndex).EffectOverTime, i)
            Else
                i = i + 1
            End If
        Loop
    End If
    Exit Sub
RemovePoisonHemo_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.RemovePoisonHemo", Erl)
End Sub

' === Neurotoxina (Fase 4) ===
' Crea o refresca un PoisonNeuroEffect. No stackea.
Public Sub CreatePoisonNeuro(ByVal SourceIndex As Integer, _
                             ByVal SourceType As e_ReferenceType, _
                             ByVal TargetIndex As Integer, _
                             ByVal TargetType As e_ReferenceType, _
                             ByVal EffectId As Integer, _
                             ByVal TickTime As Long, _
                             ByVal Duration As Long, _
                             ByVal PenalidadPunteriaPct As Long, _
                             ByVal PenalidadEvasionPct As Long, _
                             ByVal PenalidadBloqueoEscudoPct As Long, _
                             ByVal ChancePifiaHechizoPct As Long, _
                             ByVal RegenManaReduccionPct As Long, _
                             ByVal RegenManaReduccionFija As Long, _
                             ByVal BloqueaRegenManaTotal As Byte)
    On Error GoTo CreatePoisonNeuro_Err
    Dim Existing As IBaseEffectOverTime
    Dim eotIdRef As Integer
    eotIdRef = EffectId
    If TargetType = eUser Then
        Set Existing = FindEffectOnTarget(TargetIndex, UserList(TargetIndex).EffectOverTime, eotIdRef)
    ElseIf TargetType = eNpc Then
        Set Existing = FindEffectOnTarget(TargetIndex, NpcList(TargetIndex).EffectOverTime, eotIdRef)
    End If
    If Not (Existing Is Nothing) Then
        Call Existing.Reset(SourceIndex, SourceType, eotIdRef)
        Exit Sub
    End If
    Dim Effect As PoisonNeuroEffect
    Set Effect = GetEOT(e_EffectOverTimeType.ePoisonNeuro)
    UniqueIdCounter = GetNextId()
    Call Effect.Setup(SourceIndex, SourceType, TargetIndex, TargetType, eotIdRef, UniqueIdCounter, _
                       TickTime, Duration, _
                       PenalidadPunteriaPct, PenalidadEvasionPct, PenalidadBloqueoEscudoPct, _
                       ChancePifiaHechizoPct, RegenManaReduccionPct, RegenManaReduccionFija, BloqueaRegenManaTotal)
    Call AddEffectToUpdate(Effect)
    If TargetType = eUser Then
        Call AddEffect(UserList(TargetIndex).EffectOverTime, Effect)
        Call LogPoisonEvent("apply", "", UserList(TargetIndex).name, eotIdRef, 3, 0, 0, 0, 0)
    ElseIf TargetType = eNpc Then
        Call AddEffect(NpcList(TargetIndex).EffectOverTime, Effect)
    End If
    Exit Sub
CreatePoisonNeuro_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.CreatePoisonNeuro", Erl)
End Sub

Public Sub RemovePoisonNeuro(ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType)
    On Error GoTo RemovePoisonNeuro_Err
    Dim i As Integer
    If TargetType = eUser Then
        ' Limpia cache de penalidades
        With UserList(TargetIndex).flags
            .PoisonNeuroActive = 0
            .PoisonNeuroPenalidadPunteriaPct = 0
            .PoisonNeuroPenalidadEvasionPct = 0
            .PoisonNeuroPenalidadBloqueoEscudoPct = 0
            .PoisonNeuroChancePifiaHechizoPct = 0
            .PoisonNeuroRegenManaReduccionPct = 0
            .PoisonNeuroRegenManaReduccionFija = 0
            .PoisonNeuroBloqueaRegenManaTotal = 0
        End With
        Do While i < UserList(TargetIndex).EffectOverTime.EffectCount
            If UserList(TargetIndex).EffectOverTime.EffectList(i).TypeId = e_EffectOverTimeType.ePoisonNeuro Then
                ' Grupo 7 (plan 10.001, punto 7b): effect_end reason=curado (motivo 1).
                If modElementalBalanceLog.ElementalPlayerTelemetryEnabled() Then
                    Dim ebNeuroCured As PoisonNeuroEffect
                    Set ebNeuroCured = UserList(TargetIndex).EffectOverTime.EffectList(i)
                    Call modElementalBalanceLog.LogPoisonEffectEnd(TargetIndex, UserList(TargetIndex).EffectOverTime.EffectList(i).CasterArrayIndex, UserList(TargetIndex).EffectOverTime.EffectList(i).CasterRefType = eNpc, UserList(TargetIndex).EffectOverTime.EffectList(i).CasterIsValid, UserList(TargetIndex).EffectOverTime.EffectList(i).EotId, ebNeuroCured.TelemetryServedMs, 0, 1)
                End If
                UserList(TargetIndex).EffectOverTime.EffectList(i).RemoveMe = True
                Call RemoveEffectAtPos(UserList(TargetIndex).EffectOverTime, i)
            Else
                i = i + 1
            End If
        Loop
    ElseIf TargetType = eNpc Then
        Do While i < NpcList(TargetIndex).EffectOverTime.EffectCount
            If NpcList(TargetIndex).EffectOverTime.EffectList(i).TypeId = e_EffectOverTimeType.ePoisonNeuro Then
                NpcList(TargetIndex).EffectOverTime.EffectList(i).RemoveMe = True
                Call RemoveEffectAtPos(NpcList(TargetIndex).EffectOverTime, i)
            Else
                i = i + 1
            End If
        Loop
    End If
    Exit Sub
RemovePoisonNeuro_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.RemovePoisonNeuro", Erl)
End Sub

' 06.002 Ola 1: resetea los caches de flags de los venenos por efecto (Minor/Hemo/Neuro).
' OnRemove NO limpia estos caches (decision de los callers), asi que muerte/desconexion/reto
' deben resetearlos aca o quedan como penalidad fantasma sin efecto ni icono.
Public Sub ClearPoisonEffectCaches(ByVal UserIndex As Integer)
    On Error GoTo ClearPoisonEffectCaches_Err
    If UserIndex <= 0 Then Exit Sub
    With UserList(UserIndex).flags
        .PoisonMinorActive = 0
        If .PoisonHemoStacks > 0 Then Call WriteUpdatePoisonStacks(UserIndex, 0, 0)
        .PoisonHemoStacks = 0
        .PoisonNeuroActive = 0
        .PoisonNeuroPenalidadPunteriaPct = 0
        .PoisonNeuroPenalidadEvasionPct = 0
        .PoisonNeuroPenalidadBloqueoEscudoPct = 0
        .PoisonNeuroChancePifiaHechizoPct = 0
        .PoisonNeuroRegenManaReduccionPct = 0
        .PoisonNeuroRegenManaReduccionFija = 0
        .PoisonNeuroBloqueaRegenManaTotal = 0
    End With
    Exit Sub
ClearPoisonEffectCaches_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.ClearPoisonEffectCaches", Erl)
End Sub

' Aplica veneno al NPC cuando un user le pega con arma envenenada.
' Soporta Subtipo=10 (veneno fijo en el arma) y Subtipo=11 (arma untada con vial).
' Respeta resistencia/inmunidad del NPC via GetNpcPoisonResist.
Public Sub TryPoisonNpcWithWeapon(ByVal AtacanteIndex As Integer, ByVal NpcIndex As Integer, ByVal ObjInd As Integer)
    On Error GoTo TryPoisonNpcWithWeapon_Err
    If ObjInd <= 0 Then Exit Sub
    Dim familia As Byte
    Dim chanceAplicar As Long
    Dim tickMs As Long, durMs As Long
    Dim danoModo As Byte, danoMin As Long, danoMax As Long
    Dim factorPvP As Single, factorPvE As Single
    ' Hemo extras
    Dim danoStackModo As Byte, danoStackMin As Long, danoStackMax As Long
    Dim stacksMax As Integer, golpesStack As Integer, decayMs As Long, refreshStack As Byte
    ' Neuro extras
    Dim penPunteria As Long, penEvasion As Long, penEscudo As Long
    Dim chPifia As Long, regenManaPct As Long, regenManaFija As Long, bloqRegenTotal As Byte
    Dim isPoisonArrow As Boolean
    Dim isPoisonedAmmo As Boolean

    If ObjData(ObjInd).OBJType = e_OBJType.otArrows Then
        familia = ObjData(ObjInd).FlechaVenenoFamilia
        If familia > 0 Then
            isPoisonArrow = True
            chanceAplicar = ObjData(ObjInd).ChanceAplicarPct
            tickMs = ObjData(ObjInd).TickIntervaloMs
            durMs = ObjData(ObjInd).DuracionMs
            danoModo = ObjData(ObjInd).DanoModo
            danoMin = ObjData(ObjInd).DanoMin
            danoMax = ObjData(ObjInd).DanoMax
            factorPvP = ObjData(ObjInd).FactorPvP
            factorPvE = ObjData(ObjInd).FactorPvE
            danoStackModo = ObjData(ObjInd).DanoPorStackModo
            danoStackMin = ObjData(ObjInd).DanoPorStackMin
            danoStackMax = ObjData(ObjInd).DanoPorStackMax
            stacksMax = ObjData(ObjInd).StacksMax
            golpesStack = ObjData(ObjInd).GolpesQueSumanStacks
            decayMs = ObjData(ObjInd).IntervaloDecayStackMs
            refreshStack = ObjData(ObjInd).RefrescaTimerAlStackear
            penPunteria = ObjData(ObjInd).PenalidadPunteriaPct
            penEvasion = ObjData(ObjInd).PenalidadEvasionPct
            penEscudo = ObjData(ObjInd).PenalidadBloqueoEscudoPct
            chPifia = ObjData(ObjInd).ChancePifiaHechizoPct
            regenManaPct = ObjData(ObjInd).RegenManaReduccionPct
            regenManaFija = ObjData(ObjInd).RegenManaReduccionFija
            bloqRegenTotal = ObjData(ObjInd).BloqueaRegenManaTotal
        Else
            With UserList(AtacanteIndex).flags
                If .PoisonedAmmoObjIndex <> ObjInd Then Exit Sub
                If .PoisonedAmmoCargas <= 0 Then Exit Sub
                If .PoisonedAmmoDuracionMaxMs > 0 Then
                    If (GetTickCountRaw() - .PoisonedAmmoAppliedTick) > .PoisonedAmmoDuracionMaxMs Then
                        Call ClearPoisonedAmmo(AtacanteIndex, "El veneno de tus flechas se ha disipado.", "duracion_expirada")
                        Exit Sub
                    End If
                End If
                familia = .PoisonedAmmoFamilia
                If familia = 0 Then Exit Sub
                isPoisonArrow = True
                isPoisonedAmmo = True
                chanceAplicar = .PoisonedAmmoChanceAplicarPct
                tickMs = .PoisonedAmmoTickIntervaloMs
                durMs = .PoisonedAmmoDuracionEfectoMs
                danoModo = .PoisonedAmmoDanoModo
                danoMin = .PoisonedAmmoDanoMin
                danoMax = .PoisonedAmmoDanoMax
                factorPvP = .PoisonedAmmoFactorPvP
                factorPvE = .PoisonedAmmoFactorPvE
                danoStackModo = .PoisonedAmmoDanoPorStackModo
                danoStackMin = .PoisonedAmmoDanoPorStackMin
                danoStackMax = .PoisonedAmmoDanoPorStackMax
                stacksMax = .PoisonedAmmoStacksMax
                golpesStack = .PoisonedAmmoGolpesQueSumanStacks
                decayMs = .PoisonedAmmoIntervaloDecayStackMs
                refreshStack = .PoisonedAmmoRefrescaTimerAlStackear
                penPunteria = .PoisonedAmmoPenalidadPunteriaPct
                penEvasion = .PoisonedAmmoPenalidadEvasionPct
                penEscudo = .PoisonedAmmoPenalidadBloqueoEscudoPct
                chPifia = .PoisonedAmmoChancePifiaHechizoPct
                regenManaPct = .PoisonedAmmoRegenManaReduccionPct
                regenManaFija = .PoisonedAmmoRegenManaReduccionFija
                bloqRegenTotal = .PoisonedAmmoBloqueaRegenManaTotal
                If .PoisonedAmmoChancePorGolpePct > 0 Then
                    If RandomNumber(1, 100) > .PoisonedAmmoChancePorGolpePct Then Exit Sub
                End If
            End With
        End If
    Else
        Select Case ObjData(ObjInd).Subtipo
        Case 10
            familia = ObjData(ObjInd).FamiliaVeneno
            If familia = 0 Then Exit Sub
            chanceAplicar = ObjData(ObjInd).ChanceAplicarPct
            tickMs = ObjData(ObjInd).TickIntervaloMs
            durMs = ObjData(ObjInd).DuracionMs
            danoModo = ObjData(ObjInd).DanoModo
            danoMin = ObjData(ObjInd).DanoMin
            danoMax = ObjData(ObjInd).DanoMax
            factorPvP = ObjData(ObjInd).FactorPvP
            factorPvE = ObjData(ObjInd).FactorPvE
            danoStackModo = ObjData(ObjInd).DanoPorStackModo
            danoStackMin = ObjData(ObjInd).DanoPorStackMin
            danoStackMax = ObjData(ObjInd).DanoPorStackMax
            stacksMax = ObjData(ObjInd).StacksMax
            golpesStack = ObjData(ObjInd).GolpesQueSumanStacks
            decayMs = ObjData(ObjInd).IntervaloDecayStackMs
            refreshStack = ObjData(ObjInd).RefrescaTimerAlStackear
            penPunteria = ObjData(ObjInd).PenalidadPunteriaPct
            penEvasion = ObjData(ObjInd).PenalidadEvasionPct
            penEscudo = ObjData(ObjInd).PenalidadBloqueoEscudoPct
            chPifia = ObjData(ObjInd).ChancePifiaHechizoPct
            regenManaPct = ObjData(ObjInd).RegenManaReduccionPct
            regenManaFija = ObjData(ObjInd).RegenManaReduccionFija
            bloqRegenTotal = ObjData(ObjInd).BloqueaRegenManaTotal
        Case 11
            ' Arma envenenable: requiere untado activo, mismo ObjIndex equipado, cargas > 0
            With UserList(AtacanteIndex).flags
                If .PoisonedWeaponObjIndex <> ObjInd Then Exit Sub
                If .PoisonedWeaponCargas < 0 Then Exit Sub
                familia = .PoisonedWeaponFamilia
                If familia = 0 Then Exit Sub
                chanceAplicar = .PoisonedWeaponChanceAplicarPct
                tickMs = .PoisonedWeaponTickIntervaloMs
                durMs = .PoisonedWeaponDuracionEfectoMs
                danoModo = .PoisonedWeaponDanoModo
                danoMin = .PoisonedWeaponDanoMin
                danoMax = .PoisonedWeaponDanoMax
                factorPvP = .PoisonedWeaponFactorPvP
                factorPvE = .PoisonedWeaponFactorPvE
                ' Hemo extras (solo se usan si familia=2, pero los leemos siempre)
                danoStackModo = .PoisonedWeaponDanoPorStackModo
                danoStackMin = .PoisonedWeaponDanoPorStackMin
                danoStackMax = .PoisonedWeaponDanoPorStackMax
                stacksMax = .PoisonedWeaponStacksMax
                golpesStack = .PoisonedWeaponGolpesQueSumanStacks
                decayMs = .PoisonedWeaponIntervaloDecayStackMs
                refreshStack = .PoisonedWeaponRefrescaTimerAlStackear
                ' Neuro extras (solo se usan si familia=3)
                penPunteria = .PoisonedWeaponPenalidadPunteriaPct
                penEvasion = .PoisonedWeaponPenalidadEvasionPct
                penEscudo = .PoisonedWeaponPenalidadBloqueoEscudoPct
                chPifia = .PoisonedWeaponChancePifiaHechizoPct
                regenManaPct = .PoisonedWeaponRegenManaReduccionPct
                regenManaFija = .PoisonedWeaponRegenManaReduccionFija
                bloqRegenTotal = .PoisonedWeaponBloqueaRegenManaTotal
                ' Chequear chance por golpe del vial
                If .PoisonedWeaponChancePorGolpePct > 0 Then
                    If RandomNumber(1, 100) > .PoisonedWeaponChancePorGolpePct Then Exit Sub
                End If
            End With
        Case Else
            Exit Sub
        End Select
    End If
    ' Resistencia / inmunidad del NPC
    Dim r As t_PoisonResist
    r = GetNpcPoisonResist(NpcIndex, familia)
    If r.Inmune <> 0 Then Exit Sub
    Dim chFin As Long
    chFin = chanceAplicar - r.ChancePct
    If chFin <= 0 Then Exit Sub
    If RandomNumber(1, 100) > chFin Then Exit Sub
    ' Aplicar segun familia
    Select Case familia
        Case 1
            Call CreatePoisonMinor(AtacanteIndex, eUser, NpcIndex, eNpc, 63, _
                tickMs, durMs, danoModo, danoMin, danoMax, factorPvP, factorPvE)
        Case 2
            Call CreatePoisonHemo(AtacanteIndex, eUser, NpcIndex, eNpc, 64, _
                tickMs, durMs, danoModo, danoMin, danoMax, _
                danoStackModo, danoStackMin, danoStackMax, _
                stacksMax, golpesStack, decayMs, refreshStack, _
                factorPvP, factorPvE, 1)
        Case 3
            Call CreatePoisonNeuro(AtacanteIndex, eUser, NpcIndex, eNpc, 65, _
                tickMs, durMs, penPunteria, penEvasion, penEscudo, _
                chPifia, regenManaPct, regenManaFija, bloqRegenTotal)
    End Select
    If isPoisonArrow Then
        Call WriteConsoleMsg(AtacanteIndex, "Has envenenado a " & NpcList(NpcIndex).name & " con tus flechas.", e_FontTypeNames.FONTTYPE_FIGHT)
        If isPoisonedAmmo Then
            Call LogPoisonEvent("apply_arrow_vial_npc", UserList(AtacanteIndex).name, NpcList(NpcIndex).name, ObjInd, familia, 0, 0, 0, 0)
        Else
            Call LogPoisonEvent("apply_arrow_fixed_npc", UserList(AtacanteIndex).name, NpcList(NpcIndex).name, ObjInd, familia, 0, 0, 0, 0)
        End If
    End If
    Exit Sub
TryPoisonNpcWithWeapon_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.TryPoisonNpcWithWeapon", Erl)
End Sub

' === Vial / arma envenenable (Fase 6) ===
' Limpia el estado de untado del arma del usuario y le manda un mensaje opcional.
' No persiste (decision del plan): se invoca al agotarse cargas, expirar duracion maxima,
' desequipar arma, o desconectar.
Public Sub ClearPoisonedWeapon(ByVal UserIndex As Integer, Optional ByVal msg As String = "")
    On Error GoTo ClearPoisonedWeapon_Err
    If UserIndex <= 0 Then Exit Sub
    With UserList(UserIndex).flags
        If .PoisonedWeaponObjIndex = 0 And .PoisonedWeaponCargas = 0 Then Exit Sub
        ' Apagar icono UI del arma untada en el cliente (antes de borrar los flags)
        Call ClearPoisonedWeaponIcon(UserIndex, .PoisonedWeaponFamilia, .PoisonedWeaponObjIndex)
        .PoisonedWeaponObjIndex = 0
        .PoisonedWeaponFamilia = 0
        .PoisonedWeaponCargas = 0
        .PoisonedWeaponAppliedTick = 0
        .PoisonedWeaponDuracionMaxMs = 0
        .PoisonedWeaponChanceAplicarPct = 0
        .PoisonedWeaponChancePorGolpePct = 0
        .PoisonedWeaponTickIntervaloMs = 0
        .PoisonedWeaponDuracionEfectoMs = 0
        .PoisonedWeaponDanoModo = 0
        .PoisonedWeaponDanoMin = 0
        .PoisonedWeaponDanoMax = 0
        .PoisonedWeaponFactorPvP = 0
        .PoisonedWeaponFactorPvE = 0
        ' Hemo extras
        .PoisonedWeaponDanoPorStackModo = 0
        .PoisonedWeaponDanoPorStackMin = 0
        .PoisonedWeaponDanoPorStackMax = 0
        .PoisonedWeaponStacksMax = 0
        .PoisonedWeaponGolpesQueSumanStacks = 0
        .PoisonedWeaponIntervaloDecayStackMs = 0
        .PoisonedWeaponRefrescaTimerAlStackear = 0
        ' Neuro extras
        .PoisonedWeaponPenalidadPunteriaPct = 0
        .PoisonedWeaponPenalidadEvasionPct = 0
        .PoisonedWeaponPenalidadBloqueoEscudoPct = 0
        .PoisonedWeaponChancePifiaHechizoPct = 0
        .PoisonedWeaponRegenManaReduccionPct = 0
        .PoisonedWeaponRegenManaReduccionFija = 0
        .PoisonedWeaponBloqueaRegenManaTotal = 0
    End With
    If LenB(msg) > 0 Then
        Call WriteConsoleMsg(UserIndex, msg, e_FontTypeNames.FONTTYPE_INFO)
    End If
    Exit Sub
ClearPoisonedWeapon_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.ClearPoisonedWeapon", Erl)
End Sub

' Mapea familia 1/2/3 al ClientEffectTypeId del icono 'arma untada'.
Private Function GetPoisonedWeaponEffectTypeId(ByVal familia As Byte) As Integer
    Select Case familia
        Case 1: GetPoisonedWeaponEffectTypeId = 46
        Case 2: GetPoisonedWeaponEffectTypeId = 47
        Case 3: GetPoisonedWeaponEffectTypeId = 48
        Case Else: GetPoisonedWeaponEffectTypeId = 0
    End Select
End Function

' Envia el icono UI al cliente cuando se unta un arma. typeId 46=Menor 47=Hemo 48=Neuro.
' UniqueId = ObjIndex del arma (para distinguir si se reutilizara para multiples armas en el futuro).
Public Sub WritePoisonedWeaponIcon(ByVal UserIndex As Integer)
    On Error GoTo WritePoisonedWeaponIcon_Err
    If UserIndex <= 0 Then Exit Sub
    With UserList(UserIndex).flags
        If .PoisonedWeaponObjIndex <= 0 Or .PoisonedWeaponCargas <= 0 Then Exit Sub
        Dim typeId As Integer
        typeId = GetPoisonedWeaponEffectTypeId(.PoisonedWeaponFamilia)
        If typeId = 0 Then Exit Sub
        Call WriteSendSkillCdUpdate(UserIndex, typeId, CLng(.PoisonedWeaponObjIndex), .PoisonedWeaponDuracionMaxMs, .PoisonedWeaponDuracionMaxMs, eDebuff, .PoisonedWeaponCargas)
    End With
    Exit Sub
WritePoisonedWeaponIcon_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.WritePoisonedWeaponIcon", Erl)
End Sub

' Apaga el icono UI del arma untada en el cliente.
Public Sub ClearPoisonedWeaponIcon(ByVal UserIndex As Integer, ByVal familia As Byte, ByVal objIndex As Integer)
    On Error GoTo ClearPoisonedWeaponIcon_Err
    If UserIndex <= 0 Then Exit Sub
    Dim typeId As Integer
    typeId = GetPoisonedWeaponEffectTypeId(familia)
    If typeId = 0 Then Exit Sub
    Call WriteSendSkillCdUpdate(UserIndex, typeId, CLng(objIndex), 0, 0, eDebuff, 0)
    Exit Sub
ClearPoisonedWeaponIcon_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.ClearPoisonedWeaponIcon", Erl)
End Sub

' Llamado al hacer swing (acierte o no): decrementa cargas y chequea expiracion por duracion maxima.
' Si el untado se agoto/expiro, lo limpia.
Public Sub OnPoisonedWeaponSwing(ByVal AtacanteIndex As Integer)
    On Error GoTo OnPoisonedWeaponSwing_Err
    If Not IsFeatureEnabled("new_poison_system") Then Exit Sub
    With UserList(AtacanteIndex).flags
        If .PoisonedWeaponObjIndex <= 0 Or .PoisonedWeaponCargas <= 0 Then Exit Sub
        ' Verificar que el arma equipada coincida (si se cambio sin desequipar, no aplica)
        If UserList(AtacanteIndex).invent.EquippedWeaponObjIndex <> .PoisonedWeaponObjIndex Then Exit Sub
        ' Chequear expiracion por duracion maxima de untado
        If .PoisonedWeaponDuracionMaxMs > 0 Then
            If (GetTickCountRaw() - .PoisonedWeaponAppliedTick) > .PoisonedWeaponDuracionMaxMs Then
                Call ClearPoisonedWeapon(AtacanteIndex, "El veneno de tu arma se ha disipado.")
                Exit Sub
            End If
        End If
        ' Consumir 1 carga
        .PoisonedWeaponCargas = .PoisonedWeaponCargas - 1
        If .PoisonedWeaponCargas <= 0 Then
            Call ClearPoisonedWeapon(AtacanteIndex, "El veneno de tu arma se ha agotado.")
        Else
            ' Actualizar contador en cliente
            Call WriteUpdatePoisonStacks(AtacanteIndex, CLng(.PoisonedWeaponObjIndex), .PoisonedWeaponCargas)
        End If
    End With
    Exit Sub
OnPoisonedWeaponSwing_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.OnPoisonedWeaponSwing", Erl)
End Sub

' Tick periodico llamado desde el game loop. Chequea expiracion por tiempo y muestra timer visible.
' El TickInterval del game loop es ~250ms, asi que la frecuencia visible es razonable.
Public Sub CheckPoisonedWeaponTick(ByVal UserIndex As Integer)
    On Error GoTo CheckPoisonedWeaponTick_Err
    If Not IsFeatureEnabled("new_poison_system") Then Exit Sub
    With UserList(UserIndex).flags
        If .PoisonedWeaponObjIndex <= 0 Or .PoisonedWeaponCargas <= 0 Then Exit Sub
        ' Expiracion por duracion maxima sin pegar
        If .PoisonedWeaponDuracionMaxMs > 0 Then
            Dim elapsed As Long
            elapsed = GetTickCountRaw() - .PoisonedWeaponAppliedTick
            If elapsed > .PoisonedWeaponDuracionMaxMs Then
                Call ClearPoisonedWeapon(UserIndex, "El veneno de tu arma se ha disipado.")
                Exit Sub
            End If
        End If
    End With
    ' Mensaje de consola periodico eliminado: el icono UI con stacks+timer lo reemplaza.
    Exit Sub
CheckPoisonedWeaponTick_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.CheckPoisonedWeaponTick", Erl)
End Sub

' === Flechas untadas (extension venenos flechas) ===
Public Sub ClearPoisonedAmmo(ByVal UserIndex As Integer, Optional ByVal msg As String = "", Optional ByVal motivo As String = "")
    On Error GoTo ClearPoisonedAmmo_Err
    If UserIndex <= 0 Then Exit Sub
    Dim oldObjIndex As Integer
    Dim oldFamilia As Byte
    Dim oldCargas As Integer
    Dim logMotivo As String
    With UserList(UserIndex).flags
        If .PoisonedAmmoObjIndex = 0 And .PoisonedAmmoCargas = 0 Then Exit Sub
        oldObjIndex = .PoisonedAmmoObjIndex
        oldFamilia = .PoisonedAmmoFamilia
        oldCargas = .PoisonedAmmoCargas
        Call ClearPoisonedAmmoIcon(UserIndex, oldFamilia, oldObjIndex)
        .PoisonedAmmoObjIndex = 0
        .PoisonedAmmoFamilia = 0
        .PoisonedAmmoCargas = 0
        .PoisonedAmmoAppliedTick = 0
        .PoisonedAmmoDuracionMaxMs = 0
        .PoisonedAmmoDuracionEfectoMs = 0
        .PoisonedAmmoTickIntervaloMs = 0
        .PoisonedAmmoChanceAplicarPct = 0
        .PoisonedAmmoChancePorGolpePct = 0
        .PoisonedAmmoDanoModo = 0
        .PoisonedAmmoDanoMin = 0
        .PoisonedAmmoDanoMax = 0
        .PoisonedAmmoFactorPvP = 0
        .PoisonedAmmoFactorPvE = 0
        .PoisonedAmmoDanoPorStackModo = 0
        .PoisonedAmmoDanoPorStackMin = 0
        .PoisonedAmmoDanoPorStackMax = 0
        .PoisonedAmmoStacksMax = 0
        .PoisonedAmmoGolpesQueSumanStacks = 0
        .PoisonedAmmoIntervaloDecayStackMs = 0
        .PoisonedAmmoRefrescaTimerAlStackear = 0
        .PoisonedAmmoPenalidadPunteriaPct = 0
        .PoisonedAmmoPenalidadEvasionPct = 0
        .PoisonedAmmoPenalidadBloqueoEscudoPct = 0
        .PoisonedAmmoChancePifiaHechizoPct = 0
        .PoisonedAmmoRegenManaReduccionPct = 0
        .PoisonedAmmoRegenManaReduccionFija = 0
        .PoisonedAmmoBloqueaRegenManaTotal = 0
    End With
    If LenB(motivo) > 0 Then
        logMotivo = motivo
    Else
        logMotivo = "limpieza"
    End If
    Call LogPoisonEvent("clear_poisoned_ammo_" & logMotivo, UserList(UserIndex).name, "", oldObjIndex, oldFamilia, 0, oldCargas, 0, 0)
    If LenB(msg) > 0 Then
        Call WriteConsoleMsg(UserIndex, msg, e_FontTypeNames.FONTTYPE_INFO)
    End If
    Exit Sub
ClearPoisonedAmmo_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.ClearPoisonedAmmo", Erl)
End Sub

Private Function GetPoisonedAmmoEffectTypeId(ByVal familia As Byte) As Integer
    Select Case familia
        Case 1: GetPoisonedAmmoEffectTypeId = 50
        Case 2: GetPoisonedAmmoEffectTypeId = 51
        Case 3: GetPoisonedAmmoEffectTypeId = 52
        Case Else: GetPoisonedAmmoEffectTypeId = 0
    End Select
End Function

Public Sub WritePoisonedAmmoIcon(ByVal UserIndex As Integer)
    On Error GoTo WritePoisonedAmmoIcon_Err
    If UserIndex <= 0 Then Exit Sub
    With UserList(UserIndex).flags
        If .PoisonedAmmoObjIndex <= 0 Or .PoisonedAmmoCargas <= 0 Then Exit Sub
        Dim typeId As Integer
        typeId = GetPoisonedAmmoEffectTypeId(.PoisonedAmmoFamilia)
        If typeId = 0 Then Exit Sub
        Call WriteSendSkillCdUpdate(UserIndex, typeId, CLng(.PoisonedAmmoObjIndex), .PoisonedAmmoDuracionMaxMs, .PoisonedAmmoDuracionMaxMs, eDebuff, .PoisonedAmmoCargas)
    End With
    Exit Sub
WritePoisonedAmmoIcon_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.WritePoisonedAmmoIcon", Erl)
End Sub

Public Sub ClearPoisonedAmmoIcon(ByVal UserIndex As Integer, ByVal familia As Byte, ByVal objIndex As Integer)
    On Error GoTo ClearPoisonedAmmoIcon_Err
    If UserIndex <= 0 Then Exit Sub
    Dim typeId As Integer
    typeId = GetPoisonedAmmoEffectTypeId(familia)
    If typeId = 0 Then Exit Sub
    Call WriteSendSkillCdUpdate(UserIndex, typeId, CLng(objIndex), 0, 0, eDebuff, 0)
    Exit Sub
ClearPoisonedAmmoIcon_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.ClearPoisonedAmmoIcon", Erl)
End Sub

Public Sub OnPoisonedAmmoSwing(ByVal UserIndex As Integer, ByVal AmmoObjIndex As Integer)
    On Error GoTo OnPoisonedAmmoSwing_Err
    If Not IsFeatureEnabled("new_poison_system") Then Exit Sub
    If UserIndex <= 0 Or AmmoObjIndex <= 0 Then Exit Sub
    With UserList(UserIndex).flags
        If .PoisonedAmmoObjIndex <= 0 Or .PoisonedAmmoCargas <= 0 Then Exit Sub
        If .PoisonedAmmoObjIndex <> AmmoObjIndex Then Exit Sub
        If UserList(UserIndex).invent.EquippedMunitionObjIndex <> AmmoObjIndex Then Exit Sub
        If .PoisonedAmmoDuracionMaxMs > 0 Then
            If (GetTickCountRaw() - .PoisonedAmmoAppliedTick) > .PoisonedAmmoDuracionMaxMs Then
                Call ClearPoisonedAmmo(UserIndex, "El veneno de tus flechas se ha disipado.", "duracion_expirada")
                Exit Sub
            End If
        End If
        .PoisonedAmmoCargas = .PoisonedAmmoCargas - 1
        If .PoisonedAmmoCargas <= 0 Then
            Call ClearPoisonedAmmo(UserIndex, "El veneno de tus flechas se ha agotado.", "cargas_agotadas")
        Else
            Call WriteUpdatePoisonStacks(UserIndex, CLng(.PoisonedAmmoObjIndex), .PoisonedAmmoCargas)
        End If
    End With
    Exit Sub
OnPoisonedAmmoSwing_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.OnPoisonedAmmoSwing", Erl)
End Sub

' === NPCs y perfiles de veneno (Fase 8) ===
' Busca un perfil por nombre. Devuelve el indice (1-based) en PoisonProfiles, o 0 si no existe.
Public Function FindPoisonProfile(ByVal profileName As String) As Integer
    On Error GoTo FindPoisonProfile_Err
    If LenB(profileName) = 0 Or PoisonProfileCount <= 0 Then
        FindPoisonProfile = 0
        Exit Function
    End If
    Dim i As Integer
    For i = 1 To PoisonProfileCount
        If StrComp(PoisonProfiles(i).nombre, profileName, vbTextCompare) = 0 Then
            FindPoisonProfile = i
            Exit Function
        End If
    Next i
    FindPoisonProfile = 0
    Exit Function
FindPoisonProfile_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.FindPoisonProfile", Erl)
    FindPoisonProfile = 0
End Function

' Aplica el perfil del NPC al user envenenado. Dispatcher segun familia.
Public Sub ApplyPoisonProfileToUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal profileIdx As Integer)
    On Error GoTo ApplyPoisonProfileToUser_Err
    If profileIdx <= 0 Or profileIdx > PoisonProfileCount Then Exit Sub
    Dim p As t_PoisonProfile
    p = PoisonProfiles(profileIdx)
    ' Chance del perfil (resistencia del user)
    Dim resistU As t_PoisonResist
    resistU = GetUserPoisonResist(UserIndex, p.FamiliaVeneno)
    If resistU.Inmune <> 0 Then
        Call WriteConsoleMsg(UserIndex, "Resististe el veneno.", e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    Dim chFin As Long
    chFin = p.ChanceAplicarPct - resistU.ChancePct
    If chFin <= 0 Or RandomNumber(1, 100) > chFin Then
        Call WriteConsoleMsg(UserIndex, "Resististe el veneno.", e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    Select Case p.FamiliaVeneno
        Case 1
            Call CreatePoisonMinor(NpcIndex, eNpc, UserIndex, eUser, 63, _
                p.TickIntervaloMs, p.DuracionMs, _
                p.DanoModo, p.DanoMin, p.DanoMax, _
                p.FactorPvP, p.FactorPvE)
        Case 2
            Call CreatePoisonHemo(NpcIndex, eNpc, UserIndex, eUser, 64, _
                p.TickIntervaloMs, p.DuracionMs, _
                p.DanoModo, p.DanoMin, p.DanoMax, _
                p.DanoPorStackModo, p.DanoPorStackMin, p.DanoPorStackMax, _
                p.StacksMax, p.GolpesQueSumanStacks, p.IntervaloDecayStackMs, _
                p.RefrescaTimerAlStackear, p.FactorPvP, p.FactorPvE, 1)
        Case 3
            Call CreatePoisonNeuro(NpcIndex, eNpc, UserIndex, eUser, 65, _
                p.TickIntervaloMs, p.DuracionMs, _
                p.PenalidadPunteriaPct, p.PenalidadEvasionPct, _
                p.PenalidadBloqueoEscudoPct, p.ChancePifiaHechizoPct, _
                p.RegenManaReduccionPct, p.RegenManaReduccionFija, _
                p.BloqueaRegenManaTotal)
    End Select
    Call WriteConsoleMsg(UserIndex, "¡La criatura te ha envenenado!", e_FontTypeNames.FONTTYPE_FIGHT)
    Exit Sub
ApplyPoisonProfileToUser_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.ApplyPoisonProfileToUser", Erl)
End Sub

Private Function InstantiateEOT(ByVal EffectType As e_EffectOverTimeType) As IBaseEffectOverTime
    Select Case EffectType
        Case e_EffectOverTimeType.eHealthModifier
            Set InstantiateEOT = New UpdateHpOverTime
        Case e_EffectOverTimeType.eApplyModifiers
            Set InstantiateEOT = New StatModifier
        Case e_EffectOverTimeType.eProvoke
            Set InstantiateEOT = New EffectProvoke
        Case e_EffectOverTimeType.eProvoked
            Set InstantiateEOT = New EffectProvoked
        Case e_EffectOverTimeType.eTrap
            Set InstantiateEOT = New clsTrap
        Case e_EffectOverTimeType.eDrunk
            Set InstantiateEOT = New DrunkEffect
        Case e_EffectOverTimeType.eTranslation
            Set InstantiateEOT = New TranslationEffect
        Case e_EffectOverTimeType.eApplyEffectOnHit
            Set InstantiateEOT = New ApplyEffectOnHit
        Case e_EffectOverTimeType.eManaModifier
            Set InstantiateEOT = New UpdateManaOverTime
        Case e_EffectOverTimeType.ePartyBonus
            Set InstantiateEOT = New ApplyEffectToParty
        Case e_EffectOverTimeType.ePullTarget
            Set InstantiateEOT = New AttrackEffect
        Case e_EffectOverTimeType.eDelayedBlast
            Set InstantiateEOT = New DelayedBlast
        Case e_EffectOverTimeType.eUnequip
            Set InstantiateEOT = New UnequipItem
        Case e_EffectOverTimeType.eMultipleAttacks
            Set InstantiateEOT = New MultipleAttacks
        Case e_EffectOverTimeType.eProtection
            Set InstantiateEOT = New ProtectEffect
        Case e_EffectOverTimeType.eTransform
            Set InstantiateEOT = New TransformEffect
        Case e_EffectOverTimeType.eBonusDamage
            Set InstantiateEOT = New BonusDamageEffect
        Case e_EffectOverTimeType.eBuffPotenciado
            Set InstantiateEOT = New BuffPotenciadoEffect
        Case e_EffectOverTimeType.eMinimapRadar
            Set InstantiateEOT = New EffectMinimapRadar
        Case e_EffectOverTimeType.eMinimapUserDetect
            Set InstantiateEOT = New EffectMinimapUDetect
        ' --- Sistema de venenos nuevo (TOGGLE26) ---
        Case e_EffectOverTimeType.ePoisonMinor
            Set InstantiateEOT = New PoisonMinorEffect
        Case e_EffectOverTimeType.ePoisonHemo
            Set InstantiateEOT = New PoisonHemoEffect
        Case e_EffectOverTimeType.ePoisonNeuro
            Set InstantiateEOT = New PoisonNeuroEffect
        Case Else
            Debug.Assert False
    End Select
End Function

Private Function GetEOT(ByVal EffectType As e_EffectOverTimeType) As IBaseEffectOverTime
    On Error GoTo GetEOT_Err
    Set GetEOT = Nothing
    If EffectPools(EffectType).EffectCount = 0 Then
        Set GetEOT = InstantiateEOT(EffectType)
        Exit Function
    End If
    Set GetEOT = EffectPools(EffectType).EffectList(EffectPools(EffectType).EffectCount - 1)
    Set EffectPools(EffectType).EffectList(EffectPools(EffectType).EffectCount - 1) = Nothing
    EffectPools(EffectType).EffectCount = EffectPools(EffectType).EffectCount - 1
    Exit Function
GetEOT_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.GetEOT", Erl)
End Function

Private Sub RecycleEffect(ByRef Effect As IBaseEffectOverTime)
    Call AddEffect(EffectPools(Effect.TypeId), Effect)
End Sub

Public Sub AddEffectToUpdate(ByRef Effect As IBaseEffectOverTime)
    On Error GoTo AddEffectToUpdate_Err
    Call AddEffect(ActiveEffects, Effect)
    Exit Sub
AddEffectToUpdate_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.AddEffectToUpdate", Erl)
End Sub

Public Sub AddEffect(ByRef EffectList As t_EffectOverTimeList, ByRef Effect As IBaseEffectOverTime)
    On Error GoTo AddEffect_Err
    If Not IsArrayInitialized(EffectList.EffectList) Then
        ReDim EffectList.EffectList(ACTIVE_EFFECT_LIST_SIZE) As IBaseEffectOverTime
    ElseIf EffectList.EffectCount >= UBound(EffectList.EffectList) Then
        Dim newCap As Integer
        newCap = EffectList.EffectCount + (EffectList.EffectCount \ 5) + 8  ' +20% + 8
        If newCap < EffectList.EffectCount + 1 Then newCap = EffectList.EffectCount + 1
        ReDim Preserve EffectList.EffectList(newCap) As IBaseEffectOverTime
    End If
    Set EffectList.EffectList(EffectList.EffectCount) = Effect
    Call SetMask(EffectList.CallbaclMask, Effect.CallBacksMask)
    EffectList.EffectCount = EffectList.EffectCount + 1
    Exit Sub
AddEffect_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.AddEffect", Erl)
End Sub

Public Sub RemoveEffect(ByRef EffectList As t_EffectOverTimeList, ByRef Effect As IBaseEffectOverTime, Optional ByVal CallRemove As Boolean = True)
    On Error GoTo RemoveEffect_Err
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        If EffectList.EffectList(i).UniqueId() = Effect.UniqueId() Then
            Call RemoveEffectAtPos(EffectList, i, CallRemove)
            Exit Sub
        End If
    Next i
    Exit Sub
RemoveEffect_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.RemoveEffect", Erl)
End Sub

Public Function FindEffectOfTypeOnTarget(ByRef EffectList As t_EffectOverTimeList, ByVal TargetType As e_EffectType) As IBaseEffectOverTime
    On Error GoTo FindEffectOfTypeOnTarget_Err
    Set FindEffectOfTypeOnTarget = Nothing
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        If EffectList.EffectList(i).EffectType = TargetType Then
            Set FindEffectOfTypeOnTarget = EffectList.EffectList(i)
            Exit Function
        End If
    Next i
    Exit Function
FindEffectOfTypeOnTarget_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.FindEffectOnTarget", Erl)
End Function

Public Function FindEffectOnTarget(ByVal CasterIndex As Integer, ByRef EffectList As t_EffectOverTimeList, ByVal EffectId As Integer) As IBaseEffectOverTime
    On Error GoTo FindEffectOnTarget_Err
    Set FindEffectOnTarget = Nothing
    Dim EffectLimit As e_EOTTargetLimit
    EffectLimit = EffectOverTime(EffectId).Limit
    Dim i As Integer
    If EffectLimit = e_EOTTargetLimit.eAny Then
        Exit Function
    End If
    For i = 0 To EffectList.EffectCount - 1
        If EffectLimit = eSingle Or EffectLimit = eSingleByCaster Then
            If EffectList.EffectList(i).EotId = EffectId Then
                If EffectLimit = eSingle Then
                    Set FindEffectOnTarget = EffectList.EffectList(i)
                    Exit Function
                Else
                    If EffectList.EffectList(i).CasterRefType = eUser Then
                        If EffectList.EffectList(i).CasterUserId = UserList(CasterIndex).Id Then
                            Set FindEffectOnTarget = EffectList.EffectList(i)
                            Exit Function
                        End If
                    ElseIf EffectList.EffectList(i).CasterRefType = eNpc Then
                        If EffectList.EffectList(i).CasterIsValid And EffectList.EffectList(i).CasterArrayIndex = CasterIndex Then
                            Set FindEffectOnTarget = EffectList.EffectList(i)
                            Exit Function
                        End If
                    End If
                End If
            End If
        ElseIf EffectLimit = eSingleByType Then
            If EffectList.EffectList(i).TypeId = EffectOverTime(EffectId).Type Then
                Set FindEffectOnTarget = EffectList.EffectList(i)
                Exit Function
            End If
        ElseIf EffectLimit = eSingleByTypeId Then
            If EffectList.EffectList(i).SharedTypeId = EffectOverTime(EffectId).SharedTypeId Then
                Set FindEffectOnTarget = EffectList.EffectList(i)
                Exit Function
            End If
        End If
    Next i
    Exit Function
FindEffectOnTarget_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.FindEffectOnTarget", Erl)
End Function

Public Sub ClearEffectList(ByRef EffectList As t_EffectOverTimeList, Optional ByVal Filter As e_EffectType = e_EffectType.eAny, Optional ByVal ClearForDeath As Boolean = False)
    On Error GoTo ClearEffectList_Err
    Dim i As Integer
    Do While i < EffectList.EffectCount
        If (Filter = e_EffectType.eAny Or Filter = EffectList.EffectList(i).EffectType) And Not (ClearForDeath And EffectList.EffectList(i).KeepAfterDead()) Then
            EffectList.EffectList(i).RemoveMe = True
            Call RemoveEffectAtPos(EffectList, i)
        Else
            i = i + 1
        End If
    Loop
    Exit Sub
ClearEffectList_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.ClearEffectList", Erl)
End Sub

Public Sub RemoveEffectAtPos(ByRef EffectList As t_EffectOverTimeList, ByVal Position As Integer, Optional ByVal CallRemove As Boolean = True)
    On Error GoTo RemoveEffectAtPos_Err
    Dim RegenerateMask As Boolean
    RegenerateMask = EffectList.EffectList(Position).CallBacksMask > 0
    If CallRemove Then Call EffectList.EffectList(Position).OnRemove
    Dim i As Integer
    For i = Position To EffectList.EffectCount - 1
        Set EffectList.EffectList(i) = EffectList.EffectList(i + 1)
    Next i
    Set EffectList.EffectList(EffectList.EffectCount - 1) = Nothing
    EffectList.EffectCount = EffectList.EffectCount - 1
    If RegenerateMask Then
        EffectList.CallbaclMask = 0
        For i = 0 To EffectList.EffectCount - 1
            Call SetMask(EffectList.CallbaclMask, EffectList.EffectList(i).CallBacksMask)
        Next i
    End If
    Exit Sub
RemoveEffectAtPos_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.RemoveEffectAtPos", Erl)
End Sub

' Remueve todos los EOTs de la lista cuyo TypeId coincide con el tipo dado.
' Llama OnRemove de cada uno (limpia flags, envia packets, etc).
Public Sub RemoveEffectsByType(ByRef EffectList As t_EffectOverTimeList, ByVal TypeId As e_EffectOverTimeType)
    On Error GoTo RemoveEffectsByType_Err
    Dim i As Integer
    i = EffectList.EffectCount - 1
    Do While i >= 0
        If EffectList.EffectList(i).TypeId = TypeId Then
            Call RemoveEffectAtPos(EffectList, i, True)
        End If
        i = i - 1
    Loop
    Exit Sub
RemoveEffectsByType_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.RemoveEffectsByType", Erl)
End Sub

Public Sub TargetUseMagic(ByRef EffectList As t_EffectOverTimeList, ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal MagicId As Integer)
    If Not IsSet(EffectList.CallbaclMask, e_EffectCallbackMask.eTargetUseMagic) Then Exit Sub
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        Call EffectList.EffectList(i).TargetUseMagic(TargetUserId, SourceType, MagicId)
    Next i
End Sub

Public Sub TargetWillAttack(ByRef EffectList As t_EffectOverTimeList, ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    If Not IsSet(EffectList.CallbaclMask, e_EffectCallbackMask.eTargetWillAttack) Then Exit Sub
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        Call EffectList.EffectList(i).TargetWillAttack(TargetUserId, SourceType, AttackType)
    Next i
End Sub

Public Sub TargetDidHit(ByRef EffectList As t_EffectOverTimeList, ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    If Not IsSet(EffectList.CallbaclMask, e_EffectCallbackMask.eTargetDidHit) Then Exit Sub
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        Call EffectList.EffectList(i).TargetDidHit(TargetUserId, SourceType, AttackType)
    Next i
End Sub

Public Sub TargetFailedAttack(ByRef EffectList As t_EffectOverTimeList, ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    If Not IsSet(EffectList.CallbaclMask, e_EffectCallbackMask.eTargetFailedAttack) Then Exit Sub
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        Call EffectList.EffectList(i).TargetFailedAttack(TargetUserId, SourceType, AttackType)
    Next i
End Sub

Public Function TargetApplyDamageReduction(ByRef EffectList As t_EffectOverTimeList, _
                                           ByVal Damage As Long, _
                                           ByVal SourceUserId As Integer, _
                                           ByVal SourceType As e_ReferenceType, _
                                           ByVal AttackType As e_DamageSourceType) As Long
    If Not IsSet(EffectList.CallbaclMask, e_EffectCallbackMask.eTargetApplyDamageReduction) Then
        TargetApplyDamageReduction = Damage
        Exit Function
    End If
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        Damage = EffectList.EffectList(i).ApplyDamageReduction(Damage, SourceUserId, SourceType, AttackType)
        If Damage >= 0 Then
            Exit Function
        End If
    Next i
    TargetApplyDamageReduction = Damage
End Function

Public Sub TargetWasDamaged(ByRef EffectList As t_EffectOverTimeList, ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    If Not IsSet(EffectList.CallbaclMask, e_EffectCallbackMask.eTargetWasDamaged) Then Exit Sub
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        Call EffectList.EffectList(i).TargetWasDamaged(SourceUserId, SourceType, AttackType)
    Next i
End Sub

Public Sub TargetWillAttackPosition(ByRef EffectList As t_EffectOverTimeList, ByRef Position As t_WorldPos)
    If Not IsSet(EffectList.CallbaclMask, e_EffectCallbackMask.eTargetWillAttackPosition) Then Exit Sub
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        Call EffectList.EffectList(i).TargetWillAttackPosition(Position.Map, Position.x, Position.y)
    Next i
End Sub

Public Sub TargetUpdateTerrain(ByRef EffectList As t_EffectOverTimeList)
    If Not IsSet(EffectList.CallbaclMask, e_EffectCallbackMask.eTargetChangeTerrain) Then Exit Sub
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        Call EffectList.EffectList(i).TargetChangeTerrain
    Next i
End Sub

Public Sub ChangeOwner(ByVal CurrentOwner As Integer, _
                       ByVal CurrentOwnerType As e_ReferenceType, _
                       ByVal NewOwner As Integer, _
                       ByVal NewOwnerType As e_ReferenceType, _
                       ByRef Effect As IBaseEffectOverTime)
    If CurrentOwnerType = eUser Then
        Call RemoveEffect(UserList(CurrentOwner).EffectOverTime, Effect, False)
    Else
        Call RemoveEffect(NpcList(CurrentOwner).EffectOverTime, Effect, False)
    End If
    Dim PrevEffect As IBaseEffectOverTime
    If NewOwnerType = eUser Then
        Set PrevEffect = FindEffectOnTarget(Effect.CasterArrayIndex, UserList(NewOwner).EffectOverTime, Effect.EotId)
        If Not PrevEffect Is Nothing Then
            PrevEffect.RemoveMe = True
        End If
        If Effect.ChangeTarget(NewOwner, NewOwnerType) Then
            Call AddEffect(UserList(NewOwner).EffectOverTime, Effect)
        Else
            Effect.RemoveMe = True
        End If
    Else
        Set PrevEffect = FindEffectOnTarget(Effect.CasterArrayIndex, NpcList(NewOwner).EffectOverTime, Effect.EotId)
        If Not PrevEffect Is Nothing Then
            If Not EffectOverTime(Effect.EotId).Override Then
                Effect.RemoveMe = True
                Exit Sub
            End If
        End If
        If Effect.ChangeTarget(NewOwner, NewOwnerType) Then
            Call AddEffect(NpcList(NewOwner).EffectOverTime, Effect)
        Else
            Effect.RemoveMe = True
        End If
    End If
End Sub

Public Function ConvertToClientBuff(ByVal buffType As e_EffectType) As e_EffectType
    Select Case buffType
        Case e_EffectType.eInformativeBuff
            ConvertToClientBuff = eBuff
        Case e_EffectType.eInformativeDebuff
            ConvertToClientBuff = eDebuff
        Case Else
            ConvertToClientBuff = buffType
    End Select
End Function


' === Buff Potenciado (plan 25.003) ===
' Unifica toda alteracion temporal de Fuerza/Agilidad (pociones, hechizos) en un
' unico icono "Potenciado" en el panel de buffs del cliente.
' Los atributos ya fueron modificados por el caller; este effect maneja el timer
' y revierte al expirar (UserAtributos = UserAtributosBackUP).
Public Sub CreateBuffPotenciado(ByVal UserIndex As Integer, ByVal DurationSeconds As Integer)
    On Error GoTo CreateBuffPotenciado_Err
    If DurationSeconds <= 0 Then Exit Sub
    Dim DurationMs As Long
    DurationMs = CLng(DurationSeconds) * 1000
    Dim Existing As IBaseEffectOverTime
    Dim i As Integer
    For i = 0 To UserList(UserIndex).EffectOverTime.EffectCount - 1
        If UserList(UserIndex).EffectOverTime.EffectList(i).TypeId = e_EffectOverTimeType.eBuffPotenciado Then
            Set Existing = UserList(UserIndex).EffectOverTime.EffectList(i)
            Exit For
        End If
    Next i
    If Not (Existing Is Nothing) Then
        Dim Buff As BuffPotenciadoEffect
        Set Buff = Existing
        Call Buff.RefreshTimer(UserIndex, eUser, DurationMs)
        Exit Sub
    End If
    Dim NewEffect As BuffPotenciadoEffect
    Set NewEffect = GetEOT(e_EffectOverTimeType.eBuffPotenciado)
    UniqueIdCounter = GetNextId()
    Call NewEffect.Setup(UserIndex, eUser, UserIndex, eUser, UniqueIdCounter, DurationMs)
    Call AddEffectToUpdate(NewEffect)
    Call AddEffect(UserList(UserIndex).EffectOverTime, NewEffect)
    Exit Sub
CreateBuffPotenciado_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.CreateBuffPotenciado", Erl)
End Sub

Public Sub RemoveBuffPotenciado(ByVal UserIndex As Integer)
    On Error GoTo RemoveBuffPotenciado_Err
    Dim i As Integer
    Do While i < UserList(UserIndex).EffectOverTime.EffectCount
        If UserList(UserIndex).EffectOverTime.EffectList(i).TypeId = e_EffectOverTimeType.eBuffPotenciado Then
            UserList(UserIndex).EffectOverTime.EffectList(i).RemoveMe = True
            Call RemoveEffectAtPos(UserList(UserIndex).EffectOverTime, i)
        Else
            i = i + 1
        End If
    Loop
    Exit Sub
RemoveBuffPotenciado_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.RemoveBuffPotenciado", Erl)
End Sub

Public Function ApplyEotModifier(ByRef TargetRef As t_AnyReference, ByRef EffectStats As t_EffectOverTime, Optional ByVal Modifier As Single = 0)
    If IsValidRef(TargetRef) Then
        Call UpdateIncreaseModifier(TargetRef, MagicBonus, EffectStats.MagicDamageDone + EffectStats.MagicDamageDone * Modifier)
        Call UpdateIncreaseModifier(TargetRef, PhysiccalBonus, EffectStats.PhysicalDamageDone + EffectStats.PhysicalDamageDone * Modifier, EffectStats.PhysicalBonusPveOnly)
        Call UpdateIncreaseModifier(TargetRef, MagicReduction, EffectStats.MagicDamageReduction + EffectStats.MagicDamageReduction * Modifier)
        Call UpdateIncreaseModifier(TargetRef, PhysicalReduction, EffectStats.PhysicalDamageReduction + EffectStats.PhysicalDamageReduction * Modifier)
        Call UpdateIncreaseModifier(TargetRef, MovementSpeed, EffectStats.SpeedModifier + EffectStats.SpeedModifier * Modifier)
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.HitBonus, EffectStats.HitModifier + EffectStats.HitModifier * Modifier)
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.EvasionBonus, EffectStats.EvasionModifier + EffectStats.EvasionModifier * Modifier)
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.SelfHealingBonus, EffectStats.SelfHealingBonus + EffectStats.SelfHealingBonus * Modifier)
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.MagicHealingBonus, EffectStats.MagicHealingBonus + EffectStats.MagicHealingBonus * Modifier)
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.PhysicalLinearBonus, EffectStats.PhysicalLinearBonus + EffectStats.PhysicalLinearBonus * Modifier, EffectStats.PhysicalBonusPveOnly)
        If IsSet(EffectStats.ApplyStatusMask, eCCInmunity) Then
            If TargetRef.RefType = eUser Then
                If UserList(TargetRef.ArrayIndex).flags.Inmovilizado = 1 Then
                    UserList(TargetRef.ArrayIndex).flags.Inmovilizado = 0
                    Call WriteInmovilizaOK(TargetRef.ArrayIndex)
                End If
                UserList(TargetRef.ArrayIndex).flags.Inmovilizado = 0
                UserList(TargetRef.ArrayIndex).Counters.Inmovilizado = 0
                UserList(TargetRef.ArrayIndex).Counters.Paralisis = 0
                UserList(TargetRef.ArrayIndex).flags.Paralizado = 0
            End If
            Call SetStatusMask(TargetRef, eCCInmunity)
        End If
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.DefenseBonus, EffectStats.DefenseBonus + EffectStats.DefenseBonus * Modifier)
    End If
End Function

Public Function RemoveEotModifier(ByRef TargetRef As t_AnyReference, ByRef EffectStats As t_EffectOverTime, Optional ByVal Modifier As Single = 0)
    If IsValidRef(TargetRef) Then
        Call UpdateIncreaseModifier(TargetRef, MagicBonus, -(EffectStats.MagicDamageDone + EffectStats.MagicDamageDone * Modifier))
        Call UpdateIncreaseModifier(TargetRef, PhysiccalBonus, -(EffectStats.PhysicalDamageDone + EffectStats.PhysicalDamageDone * Modifier), EffectStats.PhysicalBonusPveOnly)
        Call UpdateIncreaseModifier(TargetRef, MagicReduction, -(EffectStats.MagicDamageReduction + EffectStats.MagicDamageReduction * Modifier))
        Call UpdateIncreaseModifier(TargetRef, PhysicalReduction, -(EffectStats.PhysicalDamageReduction + EffectStats.PhysicalDamageReduction * Modifier))
        Call UpdateIncreaseModifier(TargetRef, MovementSpeed, -(EffectStats.SpeedModifier + EffectStats.SpeedModifier * Modifier))
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.HitBonus, -(EffectStats.HitModifier + EffectStats.HitModifier * Modifier))
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.EvasionBonus, -(EffectStats.EvasionModifier + EffectStats.EvasionModifier * Modifier))
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.SelfHealingBonus, -(EffectStats.SelfHealingBonus + EffectStats.SelfHealingBonus * Modifier))
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.MagicHealingBonus, -(EffectStats.MagicHealingBonus + EffectStats.MagicHealingBonus * Modifier))
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.PhysicalLinearBonus, -(EffectStats.PhysicalLinearBonus + EffectStats.PhysicalLinearBonus * Modifier), EffectStats.PhysicalBonusPveOnly)
        If IsSet(EffectStats.ApplyStatusMask, eCCInmunity) Then
            Call UnsetStatusMask(TargetRef, eCCInmunity)
        End If
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.DefenseBonus, -(EffectStats.DefenseBonus + EffectStats.DefenseBonus * Modifier))
    End If
End Function
