VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PoisonHemoEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Argentum 20 Game Server - sistema de venenos (TOGGLE26 new_poison_system)
' Hemotoxina: DoT con stacks, decay, timer global.
'   - Dano por tick = DanoBase + stacks * DanoPorStack (random en su rango si Min<Max).
'   - Cada N golpes (GolpesQueSumanStacks) suma 1 stack hasta StacksMax.
'   - En cap (StacksMax) no refresca timer ni agrega stack (configurable: RefrescaTimerAlStackear).
'   - Decay: cada IntervaloDecayStackMs sin pegar pierde 1 stack. A 0 stacks por decay -> efecto se limpia.
'   - Timer global: si llega a DuracionMs sin pegar, expira aunque tenga stacks.
Option Explicit
Implements IBaseEffectOverTime

Private DotInfo As t_BaseDotInfo
Private Source As t_AnyReference

' --- parametros (de la fuente: arma/hechizo/perfil NPC) ---
Private TickTime As Long
Private TimeSinceLastTick As Long
Private DurationLeft As Long
Private DurationTotalMs As Long
Private DanoModo As Byte
Private DanoMin As Long
Private DanoMax As Long
Private FactorPvP As Single
Private FactorPvE As Single
' Hemo-specific
Private DanoPorStackModo As Byte
Private DanoPorStackMin As Long
Private DanoPorStackMax As Long
Private StacksActuales As Integer
Private StacksMax As Integer
Private GolpesQueSumanStacks As Integer
Private GolpesAcumulados As Integer
Private IntervaloDecayStackMs As Long
Private TimeSinceLastDecay As Long
Private RefrescaTimerAlStackear As Byte
' Grupo 7 (plan 10.001, punto 7b): pico de stacks alcanzado (effect_end final=stacks maximos).
Private mPeakStacks As Integer

' Setup recibe parametros desde el caller (arma con valores propios).
Public Sub Setup(ByVal SourceIndex As Integer, _
                 ByVal SourceType As e_ReferenceType, _
                 ByVal TargetIndex As Integer, _
                 ByVal TargetType As e_ReferenceType, _
                 ByVal EffectId As Integer, _
                 ByVal Id As Long, _
                 ByVal pTickTime As Long, _
                 ByVal pDuration As Long, _
                 ByVal pDanoModo As Byte, _
                 ByVal pDanoMin As Long, _
                 ByVal pDanoMax As Long, _
                 ByVal pDanoPorStackModo As Byte, _
                 ByVal pDanoPorStackMin As Long, _
                 ByVal pDanoPorStackMax As Long, _
                 ByVal pStacksMax As Integer, _
                 ByVal pGolpesQueSumanStacks As Integer, _
                 ByVal pIntervaloDecayStackMs As Long, _
                 ByVal pRefrescaTimerAlStackear As Byte, _
                 ByVal pFactorPvP As Single, _
                 ByVal pFactorPvE As Single, _
                 ByVal pStacksInicial As Integer)
    TimeSinceLastTick = 0
    TimeSinceLastDecay = 0
    GolpesAcumulados = 0
    Call SetBaseDot(DotInfo, TargetIndex, TargetType, Id, EffectId)
    Call SetRef(Source, SourceIndex, SourceType)
    TickTime = pTickTime
    DurationLeft = pDuration
    DurationTotalMs = pDuration
    DanoModo = pDanoModo
    DanoMin = pDanoMin
    DanoMax = pDanoMax
    DanoPorStackModo = pDanoPorStackModo
    DanoPorStackMin = pDanoPorStackMin
    DanoPorStackMax = pDanoPorStackMax
    StacksMax = pStacksMax
    If StacksMax <= 0 Then StacksMax = 1
    GolpesQueSumanStacks = pGolpesQueSumanStacks
    If GolpesQueSumanStacks <= 0 Then GolpesQueSumanStacks = 1
    IntervaloDecayStackMs = pIntervaloDecayStackMs
    RefrescaTimerAlStackear = pRefrescaTimerAlStackear
    FactorPvP = pFactorPvP
    FactorPvE = pFactorPvE
    StacksActuales = pStacksInicial
    If StacksActuales <= 0 Then StacksActuales = 1
    If StacksActuales > StacksMax Then StacksActuales = StacksMax
    mPeakStacks = StacksActuales
    If TargetType = eUser Then
        UserList(TargetIndex).flags.PoisonHemoStacks = StacksActuales
        If EffectOverTime(EffectId).ClientEffectTypeId > 0 Then
            Call WriteSendSkillCdUpdate(TargetIndex, EffectOverTime(EffectId).ClientEffectTypeId, Id, pDuration, pDuration, EffectOverTime(EffectId).buffType, StacksActuales)
        End If
        Call WriteUpdatePoisonStacks(TargetIndex, Id, StacksActuales)
    End If
End Sub

' Llamado externamente cuando el atacante vuelve a pegar: intenta sumar stack.
Public Sub TryAddStack()
    GolpesAcumulados = GolpesAcumulados + 1
    If GolpesAcumulados < GolpesQueSumanStacks Then Exit Sub
    GolpesAcumulados = 0
    ' Cap: si esta en max, refresca timer solo si configurado
    If StacksActuales >= StacksMax Then
        If RefrescaTimerAlStackear <> 0 Then
            DurationLeft = DurationTotalMs
            Call NotifyTimerRefresh
        End If
        Exit Sub
    End If
    StacksActuales = StacksActuales + 1
    If StacksActuales > mPeakStacks Then mPeakStacks = StacksActuales
    TimeSinceLastDecay = 0
    ' Suma stack siempre refresca timer (esto es el comportamiento principal de Hemo)
    DurationLeft = DurationTotalMs
    Call NotifyStacks
    Call NotifyTimerRefresh
End Sub

' Reduce stacks por una pocion curativa parcial. Devuelve los stacks restantes.
' Si llegan a 0 o menos, marca el efecto para remocion.
Public Function ReduceStacks(ByVal cantidad As Integer) As Integer
    StacksActuales = StacksActuales - cantidad
    If StacksActuales <= 0 Then
        StacksActuales = 0
        If DotInfo.TargetRef.RefType = eUser Then
            UserList(DotInfo.TargetRef.ArrayIndex).flags.PoisonHemoStacks = 0
        End If
        DotInfo.RemoveEffect = True
        ReduceStacks = 0
        Exit Function
    End If
    Call NotifyStacks
    ReduceStacks = StacksActuales
End Function

Private Sub NotifyStacks()
    If DotInfo.TargetRef.RefType <> eUser Then Exit Sub
    UserList(DotInfo.TargetRef.ArrayIndex).flags.PoisonHemoStacks = StacksActuales
    Call WriteUpdatePoisonStacks(DotInfo.TargetRef.ArrayIndex, DotInfo.UniqueId, StacksActuales)
End Sub

Private Sub NotifyTimerRefresh()
    If DotInfo.TargetRef.RefType <> eUser Then Exit Sub
    If EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then
        Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, DurationTotalMs, DurationTotalMs, _
                EffectOverTime(DotInfo.EotId).buffType, StacksActuales)
    End If
End Sub

' D7: foto de los parametros reales para persistir al desloguear (ver EffectsOverTime.BuildPoisonSnapshot).
' Campos: H,eot,tick,restante,total,modo,min,max,psModo,psMin,psMax,stacks,stacksMax,golpes,decayMs,refresca,pvp,pve
Public Function SnapshotString() As String
    SnapshotString = "H," & DotInfo.EotId & "," & TickTime & "," & DurationLeft & "," & DurationTotalMs & "," & DanoModo & "," & DanoMin & "," & DanoMax & "," & _
            DanoPorStackModo & "," & DanoPorStackMin & "," & DanoPorStackMax & "," & StacksActuales & "," & StacksMax & "," & GolpesQueSumanStacks & "," & _
            IntervaloDecayStackMs & "," & RefrescaTimerAlStackear & "," & EffectsOverTime.SnapNum(FactorPvP) & "," & EffectsOverTime.SnapNum(FactorPvE)
End Function

' D7: tras recrear el efecto con los parametros guardados, devuelve el tiempo restante real.
Public Sub RestoreProgress(ByVal pDurationLeft As Long)
    If pDurationLeft <= 0 Or pDurationLeft > DurationTotalMs Then Exit Sub
    DurationLeft = pDurationLeft
    If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then
        Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, DurationLeft, DurationTotalMs, _
                EffectOverTime(DotInfo.EotId).buffType, StacksActuales)
    End If
End Sub

' --- IBaseEffectOverTime impl ---
Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.ePoisonHemo
End Property

Public Property Get IBaseEffectOverTime_SharedTypeId() As Integer
    IBaseEffectOverTime_SharedTypeId = EffectOverTime(DotInfo.EotId).SharedTypeId
End Property

Public Property Get IBaseEffectOverTime_UniqueId() As Long
    IBaseEffectOverTime_UniqueId = DotInfo.UniqueId
End Property

Public Property Get IBaseEffectOverTime_CasterIsValid() As Boolean
    IBaseEffectOverTime_CasterIsValid = IsValidRef(Source)
End Property

Public Property Get IBaseEffectOverTime_CasterArrayIndex() As Integer
    IBaseEffectOverTime_CasterArrayIndex = Source.ArrayIndex
End Property

Public Property Get IBaseEffectOverTime_CasterRefType() As e_ReferenceType
    IBaseEffectOverTime_CasterRefType = Source.RefType
End Property

Public Property Get IBaseEffectOverTime_CasterUserId() As Long
    IBaseEffectOverTime_CasterUserId = Source.UserId
End Property

Public Property Get IBaseEffectOverTime_TargetIsValid() As Boolean
    IBaseEffectOverTime_TargetIsValid = IsValidRef(DotInfo.TargetRef)
End Property

Public Property Get IBaseEffectOverTime_TargetArrayIndex() As Integer
    IBaseEffectOverTime_TargetArrayIndex = DotInfo.TargetRef.ArrayIndex
End Property

Public Property Get IBaseEffectOverTime_TargetRefType() As e_ReferenceType
    IBaseEffectOverTime_TargetRefType = DotInfo.TargetRef.RefType
End Property

Public Property Get IBaseEffectOverTime_TargetUserId() As Long
    IBaseEffectOverTime_TargetUserId = DotInfo.TargetRef.UserId
End Property

Public Property Get IBaseEffectOverTime_EffectType() As e_EffectType
    IBaseEffectOverTime_EffectType = EffectOverTime(DotInfo.EotId).buffType
End Property

Public Sub IBaseEffectOverTime_Update(ByVal DeltaTime As Long)
    If Not IsValidRef(DotInfo.TargetRef) Then
        DotInfo.RemoveEffect = True
        Exit Sub
    End If
    DurationLeft = DurationLeft - DeltaTime
    TimeSinceLastTick = TimeSinceLastTick + DeltaTime
    TimeSinceLastDecay = TimeSinceLastDecay + DeltaTime
    If TimeSinceLastTick >= TickTime Then
        TimeSinceLastTick = 0
        Call PerformTick
    End If
    If IntervaloDecayStackMs > 0 And TimeSinceLastDecay >= IntervaloDecayStackMs Then
        TimeSinceLastDecay = 0
        If StacksActuales > 1 Then
            ' Decay baja stack pero nunca llega a 0. Solo el timer global limpia el efecto.
            StacksActuales = StacksActuales - 1
            Call NotifyStacks
        End If
    End If
    If DurationLeft <= 0 Then
        ' Timer global expiro
        If DotInfo.TargetRef.RefType = eUser Then
            ' Grupo 7 (plan 10.001, punto 7b): effect_end reason=expiro (motivo 0), leido
            ' ANTES de resetear PoisonHemoStacks (necesitamos mPeakStacks, no el actual).
            Call modElementalBalanceLog.LogPoisonEffectEnd(DotInfo.TargetRef.ArrayIndex, Source.ArrayIndex, Source.RefType = eNpc, IsValidRef(Source), DotInfo.EotId, DurationTotalMs - DurationLeft, mPeakStacks, 0)
            UserList(DotInfo.TargetRef.ArrayIndex).flags.PoisonHemoStacks = 0
        End If
        DotInfo.RemoveEffect = True
    End If
End Sub

Public Property Get IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = DotInfo.RemoveEffect
End Property

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal value As Boolean)
    DotInfo.RemoveEffect = value
End Property

Public Function IBaseEffectOverTime_Reset(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal NewEffectId As Integer) As Boolean
    ' Reset Hemo = TryAddStack (este es el comportamiento principal)
    Call TryAddStack
    IBaseEffectOverTime_Reset = True
End Function

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = DotInfo.EotId
End Property

' Grupo 7 (plan 10.001, punto 7b): duracion servida y pico de stacks para effect_end.
Public Property Get TelemetryServedMs() As Long
    TelemetryServedMs = DurationTotalMs - DurationLeft
    If TelemetryServedMs < 0 Then TelemetryServedMs = 0
End Property

Public Property Get TelemetryPeakStacks() As Integer
    TelemetryPeakStacks = mPeakStacks
End Property

Public Sub IBaseEffectOverTime_OnRemove()
    If DotInfo.Removed Then Exit Sub
    DotInfo.Removed = True
    DotInfo.RemoveEffect = True
    If DotInfo.TargetRef.RefType = eUser Then
        If EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then
            Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, 0, 0, EffectOverTime( _
                    DotInfo.EotId).buffType, 0)
        End If
        ' Apaga el contador de stacks en el cliente
        Call WriteUpdatePoisonStacks(DotInfo.TargetRef.ArrayIndex, DotInfo.UniqueId, 0)
        ' OJO: NO se resetea flags.PoisonHemoStacks aca para preservar persistencia en logout.
        ' La limpieza por muerte/cura la hacen los callers respectivos.
    End If
End Sub

' --- callbacks no usados por Hemo ---
Public Sub IBaseEffectOverTime_TargetUseMagic(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal MagicId As Integer)
End Sub

Public Sub IBaseEffectOverTime_TargetWillAttack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Sub IBaseEffectOverTime_TargetDidHit(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Sub IBaseEffectOverTime_TargetFailedAttack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Sub IBaseEffectOverTime_TargetWasDamaged(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Function IBaseEffectOverTime_ApplyDamageReduction(ByVal Damage As Long, ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, _
                                                         ByVal AttackType As e_DamageSourceType) As Long
    IBaseEffectOverTime_ApplyDamageReduction = Damage
End Function

Public Sub IBaseEffectOverTime_TargetWillAttackPosition(ByVal Map As Integer, ByVal PosX As Integer, ByVal PosY As Integer)
End Sub

Public Function IBaseEffectOverTime_ChangeTarget(ByVal NewTargetIndex As Integer, ByVal NewTargetType As e_ReferenceType) As Boolean
    Call SetRef(DotInfo.TargetRef, NewTargetIndex, NewTargetType)
    IBaseEffectOverTime_ChangeTarget = True
End Function

Public Property Get IBaseEffectOverTime_CallBacksMask() As Long
    IBaseEffectOverTime_CallBacksMask = 0
End Property

Public Property Get IBaseEffectOverTime_KeepAfterDead() As Boolean
    IBaseEffectOverTime_KeepAfterDead = False
End Property

Public Sub IBaseEffectOverTime_TargetChangeTerrain()
End Sub

' --- Logica propia ---
Private Sub PerformTick()
    Dim TickDmg As Long
    Dim baseDmg As Long
    Dim stackDmg As Long
    Dim factor As Single
    ' Roll base y por-stack via receta unica compartida (modElementalCombat).
    ' RandomNumber(x,x)=x, asi que el helper cubre el caso min==max sin la guarda local.
    baseDmg = modElementalCombat.RollDotTickBase(DotInfo.TargetRef.ArrayIndex, DotInfo.TargetRef.RefType, DanoModo, DanoMin, DanoMax)
    stackDmg = modElementalCombat.RollDotTickBase(DotInfo.TargetRef.ArrayIndex, DotInfo.TargetRef.RefType, DanoPorStackModo, DanoPorStackMin, DanoPorStackMax)
    TickDmg = baseDmg + (stackDmg * StacksActuales)
    factor = 1
    If IsValidRef(Source) And Source.RefType = eUser And DotInfo.TargetRef.RefType = eUser Then
        factor = FactorPvP
    Else
        factor = FactorPvE
    End If
    If factor <= 0 Then factor = 1
    TickDmg = CLng(TickDmg * factor)
    ' Resistencia al DAÑO del tipo Veneno (motor elemental unificado - Step 4).
    ' La inmunidad/chance de no envenenarse es resist-al-EFECTO y se chequea al aplicar.
    Dim rE As t_ElementalResist
    Dim nulPoison As Boolean
    If DotInfo.TargetRef.RefType = eUser Then
        rE = GetUserElementalResist(DotInfo.TargetRef.ArrayIndex, e_ElementalDamageType.eDmgPoison)
        TickDmg = ApplyElementalResist(TickDmg, rE, e_ElementalDamageType.eDmgPoison, nulPoison)
    ElseIf DotInfo.TargetRef.RefType = eNpc Then
        rE = GetNpcElementalResist(DotInfo.TargetRef.ArrayIndex, e_ElementalDamageType.eDmgPoison)
        TickDmg = ApplyElementalResist(TickDmg, rE, e_ElementalDamageType.eDmgPoison, nulPoison)
    End If
    If TickDmg <= 0 Then Exit Sub
    ' FX al tick (todos los visibles lo ven)
    If DotInfo.TargetRef.RefType = eUser Then
        Dim tickFx As Integer
        tickFx = EffectOverTime(DotInfo.EotId).TickFX
        If tickFx > 0 Then
            ' --- Particula 339 (PoisonGas Copia / Hemo) sobre el sprite del envenenado ---
            ' Hardcoded por ahora. Backlog: TickFX desde .dat como fuente de verdad para todos los EOTs.
            If IsVisible(UserList(DotInfo.TargetRef.ArrayIndex)) Then
                Call SendData(SendTarget.ToPCAliveArea, DotInfo.TargetRef.ArrayIndex, _
                    PrepareMessageParticleFX(UserList(DotInfo.TargetRef.ArrayIndex).Char.charindex, 339, 30, False, , _
                        UserList(DotInfo.TargetRef.ArrayIndex).pos.x, UserList(DotInfo.TargetRef.ArrayIndex).pos.y))
            Else
                Call SendData(SendTarget.ToIndex, DotInfo.TargetRef.ArrayIndex, _
                    PrepareMessageParticleFX(UserList(DotInfo.TargetRef.ArrayIndex).Char.charindex, 339, 30, False, , _
                        UserList(DotInfo.TargetRef.ArrayIndex).pos.x, UserList(DotInfo.TargetRef.ArrayIndex).pos.y))
            End If
        End If
        If EffectsOverTime.DotSourceCanCredit(Source, UserList(DotInfo.TargetRef.ArrayIndex).pos.Map) Then
            Call UserMod.DoDamageOrHeal(DotInfo.TargetRef.ArrayIndex, Source.ArrayIndex, Source.RefType, -TickDmg, e_dot, DotInfo.EotId, DamageColor:=vbGreen)
        Else
            Call UserMod.DoDamageOrHeal(DotInfo.TargetRef.ArrayIndex, 0, e_ReferenceType.eNone, -TickDmg, e_dot, DotInfo.EotId, DamageColor:=vbGreen)
        End If
        Call LogPoisonEvent("tick", "", UserList(DotInfo.TargetRef.ArrayIndex).name, DotInfo.EotId, 2, TickDmg, StacksActuales, 0, factor)
    ElseIf DotInfo.TargetRef.RefType = eNpc Then
        Dim tickFxNpc As Integer
        tickFxNpc = EffectOverTime(DotInfo.EotId).TickFX
        If tickFxNpc > 0 Then
            Call SendData(SendTarget.ToNPCAliveArea, DotInfo.TargetRef.ArrayIndex, _
                PrepareMessageParticleFX(NpcList(DotInfo.TargetRef.ArrayIndex).Char.charindex, 339, 30, False, , _
                    NpcList(DotInfo.TargetRef.ArrayIndex).pos.x, NpcList(DotInfo.TargetRef.ArrayIndex).pos.y))
        End If
        If EffectsOverTime.DotSourceCanCredit(Source, NpcList(DotInfo.TargetRef.ArrayIndex).pos.Map) Then
            Call NPCs.DoDamageOrHeal(DotInfo.TargetRef.ArrayIndex, Source.ArrayIndex, Source.RefType, -TickDmg, e_dot, DotInfo.EotId, DamageColor:=vbGreen)
        Else
            Call NPCs.DoDamageOrHeal(DotInfo.TargetRef.ArrayIndex, 0, e_ReferenceType.eNone, -TickDmg, e_dot, DotInfo.EotId, DamageColor:=vbGreen)
        End If
    End If
End Sub
