VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PoisonNeuroEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Argentum 20 Game Server - sistema de venenos (TOGGLE26 new_poison_system)
' Neurotoxina: debuff tactico, no letal. Sin DoT (TickDmg = 0).
' Penalidades cacheadas en t_UserFlags.PoisonNeuro* para consulta rapida desde combate/regen.
' Setup aplica las penalidades, OnRemove no las limpia (lo hacen los callers como Menor/Hemo).
' Refresca timer al reaplicar (Reset). No stackea.
Option Explicit
Implements IBaseEffectOverTime

Private DotInfo As t_BaseDotInfo
Private Source As t_AnyReference

Private TickTime As Long
Private TimeSinceLastTick As Long
Private DurationLeft As Long
Private DurationTotalMs As Long
Private PenalidadPunteriaPct As Long
Private PenalidadEvasionPct As Long
Private PenalidadBloqueoEscudoPct As Long
Private ChancePifiaHechizoPct As Long
Private RegenManaReduccionPct As Long
Private RegenManaReduccionFija As Long
Private BloqueaRegenManaTotal As Byte

Public Sub Setup(ByVal SourceIndex As Integer, _
                 ByVal SourceType As e_ReferenceType, _
                 ByVal TargetIndex As Integer, _
                 ByVal TargetType As e_ReferenceType, _
                 ByVal EffectId As Integer, _
                 ByVal Id As Long, _
                 ByVal pTickTime As Long, _
                 ByVal pDuration As Long, _
                 ByVal pPenalidadPunteriaPct As Long, _
                 ByVal pPenalidadEvasionPct As Long, _
                 ByVal pPenalidadBloqueoEscudoPct As Long, _
                 ByVal pChancePifiaHechizoPct As Long, _
                 ByVal pRegenManaReduccionPct As Long, _
                 ByVal pRegenManaReduccionFija As Long, _
                 ByVal pBloqueaRegenManaTotal As Byte)
    TimeSinceLastTick = 0
    Call SetBaseDot(DotInfo, TargetIndex, TargetType, Id, EffectId)
    Call SetRef(Source, SourceIndex, SourceType)
    TickTime = pTickTime
    DurationLeft = pDuration
    DurationTotalMs = pDuration
    PenalidadPunteriaPct = pPenalidadPunteriaPct
    PenalidadEvasionPct = pPenalidadEvasionPct
    PenalidadBloqueoEscudoPct = pPenalidadBloqueoEscudoPct
    ChancePifiaHechizoPct = pChancePifiaHechizoPct
    RegenManaReduccionPct = pRegenManaReduccionPct
    RegenManaReduccionFija = pRegenManaReduccionFija
    BloqueaRegenManaTotal = pBloqueaRegenManaTotal
    If TargetType = eUser Then
        With UserList(TargetIndex).flags
            .PoisonNeuroActive = 1
            .PoisonNeuroPenalidadPunteriaPct = pPenalidadPunteriaPct
            .PoisonNeuroPenalidadEvasionPct = pPenalidadEvasionPct
            .PoisonNeuroPenalidadBloqueoEscudoPct = pPenalidadBloqueoEscudoPct
            .PoisonNeuroChancePifiaHechizoPct = pChancePifiaHechizoPct
            .PoisonNeuroRegenManaReduccionPct = pRegenManaReduccionPct
            .PoisonNeuroRegenManaReduccionFija = pRegenManaReduccionFija
            .PoisonNeuroBloqueaRegenManaTotal = pBloqueaRegenManaTotal
        End With
        If EffectOverTime(EffectId).ClientEffectTypeId > 0 Then
            Call WriteSendSkillCdUpdate(TargetIndex, EffectOverTime(EffectId).ClientEffectTypeId, Id, pDuration, pDuration, EffectOverTime(EffectId).buffType)
        End If
    End If
End Sub

' D7: foto de los parametros reales para persistir al desloguear (ver EffectsOverTime.BuildPoisonSnapshot).
' Campos: N,eot,tick,restante,total,punteria,evasion,bloqueoEscudo,pifia,regenPct,regenFija,bloqueaRegen
Public Function SnapshotString() As String
    SnapshotString = "N," & DotInfo.EotId & "," & TickTime & "," & DurationLeft & "," & DurationTotalMs & "," & PenalidadPunteriaPct & "," & PenalidadEvasionPct & "," & _
            PenalidadBloqueoEscudoPct & "," & ChancePifiaHechizoPct & "," & RegenManaReduccionPct & "," & RegenManaReduccionFija & "," & BloqueaRegenManaTotal
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
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.ePoisonNeuro
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
    If DurationLeft <= 0 Then
        ' Expiracion: limpia penalidades cacheadas
        If DotInfo.TargetRef.RefType = eUser Then
            ' Grupo 7 (plan 10.001, punto 7b): effect_end reason=expiro (motivo 0).
            Call modElementalBalanceLog.LogPoisonEffectEnd(DotInfo.TargetRef.ArrayIndex, Source.ArrayIndex, Source.RefType = eNpc, IsValidRef(Source), DotInfo.EotId, DurationTotalMs - DurationLeft, 0, 0)
            Call ClearNeuroCache(DotInfo.TargetRef.ArrayIndex)
        End If
        DotInfo.RemoveEffect = True
    End If
End Sub

Private Sub ClearNeuroCache(ByVal UserIndex As Integer)
    With UserList(UserIndex).flags
        .PoisonNeuroActive = 0
        .PoisonNeuroPenalidadPunteriaPct = 0
        .PoisonNeuroPenalidadEvasionPct = 0
        .PoisonNeuroPenalidadBloqueoEscudoPct = 0
        .PoisonNeuroChancePifiaHechizoPct = 0
        .PoisonNeuroRegenManaReduccionPct = 0
        .PoisonNeuroRegenManaReduccionFija = 0
        .PoisonNeuroBloqueaRegenManaTotal = 0
    End With
End Sub

Public Property Get IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = DotInfo.RemoveEffect
End Property

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal value As Boolean)
    DotInfo.RemoveEffect = value
End Property

Public Function IBaseEffectOverTime_Reset(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal NewEffectId As Integer) As Boolean
    ' Reset = refresh timer, no acumula
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

' Grupo 7 (plan 10.001, punto 7b): duracion servida para effect_end.
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
        ' OJO: NO se limpia el cache aca. Lo hacen los callers (Update en expiracion natural,
        ' RemovePoisonNeuro al curar/muerte) para preservar persistencia al logout.
    End If
End Sub

' --- callbacks no usados ---
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
