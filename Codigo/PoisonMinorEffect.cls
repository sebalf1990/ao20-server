VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PoisonMinorEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Argentum 20 Game Server - sistema de venenos (TOGGLE26 new_poison_system)
' Veneno Menor: DoT que no stackea, refresca timer al reaplicar.
' Los numeros (dano, tick, duracion) vienen de la fuente (arma/hechizo/perfil NPC),
' no del preset en EffectsOverTime.dat.
Option Explicit
Implements IBaseEffectOverTime

Private DotInfo As t_BaseDotInfo
Private Source As t_AnyReference

' --- parametros propios (vienen del caller, NO del preset) ---
Private TickTime As Long
Private TimeSinceLastTick As Long
Private DurationLeft As Long
Private DurationTotalMs As Long
Private DanoModo As Byte
Private DanoMin As Long
Private DanoMax As Long
Private FactorPvP As Single
Private FactorPvE As Single

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
                 ByVal pFactorPvP As Single, _
                 ByVal pFactorPvE As Single)
    TimeSinceLastTick = 0
    Call SetBaseDot(DotInfo, TargetIndex, TargetType, Id, EffectId)
    Call SetRef(Source, SourceIndex, SourceType)
    TickTime = pTickTime
    DurationLeft = pDuration
    DurationTotalMs = pDuration
    DanoModo = pDanoModo
    DanoMin = pDanoMin
    DanoMax = pDanoMax
    FactorPvP = pFactorPvP
    FactorPvE = pFactorPvE
    If TargetType = eUser Then
        UserList(TargetIndex).flags.PoisonMinorActive = 1
        If EffectOverTime(EffectId).ClientEffectTypeId > 0 Then
            Call WriteSendSkillCdUpdate(TargetIndex, EffectOverTime(EffectId).ClientEffectTypeId, Id, pDuration, pDuration, EffectOverTime(EffectId).buffType)
        End If
    End If
End Sub

' D7: foto de los parametros reales para persistir al desloguear (ver EffectsOverTime.BuildPoisonSnapshot).
' Campos: M,eot,tick,restante,total,modo,min,max,pvp,pve
Public Function SnapshotString() As String
    SnapshotString = "M," & DotInfo.EotId & "," & TickTime & "," & DurationLeft & "," & DurationTotalMs & "," & DanoModo & "," & DanoMin & "," & DanoMax & "," & _
            EffectsOverTime.SnapNum(FactorPvP) & "," & EffectsOverTime.SnapNum(FactorPvE)
End Function

' D7: tras recrear el efecto con los parametros guardados, devuelve el tiempo restante real.
Public Sub RestoreProgress(ByVal pDurationLeft As Long)
    If pDurationLeft <= 0 Or pDurationLeft > DurationTotalMs Then Exit Sub
    DurationLeft = pDurationLeft
    If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then
        Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, DurationLeft, DurationTotalMs, _
                EffectOverTime(DotInfo.EotId).buffType)
    End If
End Sub

' --- IBaseEffectOverTime impl ---
Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.ePoisonMinor
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
    If TimeSinceLastTick >= TickTime Then
        TimeSinceLastTick = 0
        Call PerformTick
    End If
    If DurationLeft <= 0 Then
        ' Expiracion natural: limpia flag de persistencia
        If DotInfo.TargetRef.RefType = eUser Then
            ' Grupo 7 (plan 10.001, punto 7b): effect_end reason=expiro (motivo 0).
            Call modElementalBalanceLog.LogPoisonEffectEnd(DotInfo.TargetRef.ArrayIndex, Source.ArrayIndex, Source.RefType = eNpc, IsValidRef(Source), DotInfo.EotId, DurationTotalMs - DurationLeft, 0, 0)
            UserList(DotInfo.TargetRef.ArrayIndex).flags.PoisonMinorActive = 0
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
    ' Refresh: timer vuelve al maximo, no acumula dano
    TimeSinceLastTick = 0
    DurationLeft = DurationTotalMs
    If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then
        Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, DurationTotalMs, DurationTotalMs, _
                EffectOverTime(DotInfo.EotId).buffType)
    End If
    IBaseEffectOverTime_Reset = True
End Function

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = DotInfo.EotId
End Property

' Grupo 7 (plan 10.001, punto 7b): duracion servida para effect_end. No forma
' parte de IBaseEffectOverTime (no hace falta agrandar la interfaz para esto).
Public Property Get TelemetryServedMs() As Long
    TelemetryServedMs = DurationTotalMs - DurationLeft
    If TelemetryServedMs < 0 Then TelemetryServedMs = 0
End Property

Public Sub IBaseEffectOverTime_OnRemove()
    If DotInfo.Removed Then Exit Sub
    DotInfo.Removed = True
    DotInfo.RemoveEffect = True
    If DotInfo.TargetRef.RefType = eUser Then
        If EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then
            Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, 0, 0, EffectOverTime( _
                    DotInfo.EotId).buffType)
        End If
        ' OJO: NO se resetea flags.PoisonMinorActive aca. La limpieza al morir y al curar
        ' la manejan los callers (Modulo_UsUaRiOs/CurePoison) seteando el flag explicito.
        ' Esto permite que el flag permanezca en 1 al logout (persistencia).
    End If
End Sub

' --- callbacks no usados por Menor ---
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
    Dim factor As Single
    factor = 1
    ' Roll base del tick via receta unica compartida (modElementalCombat).
    ' El veneno trabaja con dano positivo; se niega mas abajo al aplicar.
    TickDmg = modElementalCombat.RollDotTickBase(DotInfo.TargetRef.ArrayIndex, DotInfo.TargetRef.RefType, DanoModo, DanoMin, DanoMax)
    ' Factor PvP/PvE
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
    If TickDmg <= 0 Then Exit Sub ' tick sin dano (resistencias / factores)
    If DotInfo.TargetRef.RefType = eUser Then
        Dim tickFx As Integer
        tickFx = EffectOverTime(DotInfo.EotId).TickFX
        If tickFx > 0 Then
            ' --- Partícula PoisonGas (32) sobre el sprite del envenenado ---
            If IsVisible(UserList(DotInfo.TargetRef.ArrayIndex)) Then
                Call SendData(SendTarget.ToPCAliveArea, DotInfo.TargetRef.ArrayIndex, _
                    PrepareMessageParticleFX(UserList(DotInfo.TargetRef.ArrayIndex).Char.charindex, e_ParticleEffects.PoisonGas, 30, False, , _
                        UserList(DotInfo.TargetRef.ArrayIndex).pos.x, UserList(DotInfo.TargetRef.ArrayIndex).pos.y))
            Else
                Call SendData(SendTarget.ToIndex, DotInfo.TargetRef.ArrayIndex, _
                    PrepareMessageParticleFX(UserList(DotInfo.TargetRef.ArrayIndex).Char.charindex, e_ParticleEffects.PoisonGas, 30, False, , _
                        UserList(DotInfo.TargetRef.ArrayIndex).pos.x, UserList(DotInfo.TargetRef.ArrayIndex).pos.y))
            End If
        End If
        If EffectsOverTime.DotSourceCanCredit(Source, UserList(DotInfo.TargetRef.ArrayIndex).pos.Map) Then
            Call UserMod.DoDamageOrHeal(DotInfo.TargetRef.ArrayIndex, Source.ArrayIndex, Source.RefType, -TickDmg, e_dot, DotInfo.EotId, DamageColor:=vbGreen)
        Else
            Call UserMod.DoDamageOrHeal(DotInfo.TargetRef.ArrayIndex, 0, e_ReferenceType.eNone, -TickDmg, e_dot, DotInfo.EotId, DamageColor:=vbGreen)
        End If
        Call LogPoisonEvent("tick", "", UserList(DotInfo.TargetRef.ArrayIndex).name, DotInfo.EotId, 1, TickDmg, 0, 0, factor)
    ElseIf DotInfo.TargetRef.RefType = eNpc Then
        Dim tickFxNpc As Integer
        tickFxNpc = EffectOverTime(DotInfo.EotId).TickFX
        If tickFxNpc > 0 Then
            Call SendData(SendTarget.ToNPCAliveArea, DotInfo.TargetRef.ArrayIndex, _
                PrepareMessageParticleFX(NpcList(DotInfo.TargetRef.ArrayIndex).Char.charindex, e_ParticleEffects.PoisonGas, 30, False, , _
                    NpcList(DotInfo.TargetRef.ArrayIndex).pos.x, NpcList(DotInfo.TargetRef.ArrayIndex).pos.y))
        End If
        If EffectsOverTime.DotSourceCanCredit(Source, NpcList(DotInfo.TargetRef.ArrayIndex).pos.Map) Then
            Call NPCs.DoDamageOrHeal(DotInfo.TargetRef.ArrayIndex, Source.ArrayIndex, Source.RefType, -TickDmg, e_dot, DotInfo.EotId, DamageColor:=vbGreen)
        Else
            Call NPCs.DoDamageOrHeal(DotInfo.TargetRef.ArrayIndex, 0, e_ReferenceType.eNone, -TickDmg, e_dot, DotInfo.EotId, DamageColor:=vbGreen)
        End If
    End If
End Sub
