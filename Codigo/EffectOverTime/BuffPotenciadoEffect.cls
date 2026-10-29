VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuffPotenciadoEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
'
'    Efecto unificado "Potenciado" - canaliza toda alteracion temporal de
'    Fuerza/Agilidad (pociones, hechizos) hacia el panel moderno de buffs del
'    cliente como un unico icono.
'
'    Plan: ia/archivo/plans/2026/2026-05-25-unificar-buff-potenciado-panel-buffs.md
'
Option Explicit
Implements IBaseEffectOverTime

' GRH/Effect.ini index en cliente (ver dev/Recursos/init/Effects.ini[Effect49]).
Private Const CLIENT_EFFECT_TYPE_ID_POTENCIADO As Integer = 49

Private DotInfo            As t_BaseDotInfo
Private Source             As t_AnyReference
Private TotalDurationMs    As Long
Private ElapsedMs          As Long

Public Sub Setup(ByVal SourceIndex As Integer, _
                 ByVal SourceType As e_ReferenceType, _
                 ByVal TargetIndex As Integer, _
                 ByVal TargetType As e_ReferenceType, _
                 ByVal Id As Long, _
                 ByVal DurationMs As Long)
    Call SetBaseDot(DotInfo, TargetIndex, TargetType, Id, 0)
    Call SetRef(Source, SourceIndex, SourceType)
    TotalDurationMs = DurationMs
    ElapsedMs = 0
    DotInfo.RemoveEffect = False
    DotInfo.Removed = False
    If TargetType = eUser Then
        Call WriteSendSkillCdUpdate(TargetIndex, CLIENT_EFFECT_TYPE_ID_POTENCIADO, Id, DurationMs, DurationMs, eBuff)
    End If
End Sub

Public Sub RefreshTimer(ByVal SourceIndex As Integer, _
                       ByVal SourceType As e_ReferenceType, _
                       ByVal DurationMs As Long)
    Call SetRef(Source, SourceIndex, SourceType)
    TotalDurationMs = DurationMs
    ElapsedMs = 0
    DotInfo.RemoveEffect = False
    If DotInfo.TargetRef.RefType = eUser Then
        Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, CLIENT_EFFECT_TYPE_ID_POTENCIADO, DotInfo.UniqueId, DurationMs, DurationMs, eBuff)
    End If
End Sub

Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.eBuffPotenciado
End Property

Public Property Get IBaseEffectOverTime_SharedTypeId() As Integer
    IBaseEffectOverTime_SharedTypeId = 0
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
    IBaseEffectOverTime_EffectType = eBuff
End Property

Public Sub IBaseEffectOverTime_Update(ByVal DeltaTime As Long)
    ElapsedMs = ElapsedMs + DeltaTime
    If ElapsedMs >= TotalDurationMs Or Not IsValidRef(DotInfo.TargetRef) Then
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
    ' Reset estandar via IBase no aplica (Potenciado no usa tabla EffectOverTime).
    ' La renovacion la maneja EffectsOverTime.CreateBuffPotenciado via RefreshTimer.
    IBaseEffectOverTime_Reset = False
End Function

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = 0
End Property

Public Sub IBaseEffectOverTime_OnRemove()
    If DotInfo.Removed Then Exit Sub
    DotInfo.Removed = True
    DotInfo.RemoveEffect = True
    If DotInfo.TargetRef.RefType = eUser Then
        Dim UserIndex As Integer
        UserIndex = DotInfo.TargetRef.ArrayIndex
        Dim LoopX As Integer
        For LoopX = 1 To NUMATRIBUTOS
            UserList(UserIndex).Stats.UserAtributos(LoopX) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopX)
        Next LoopX
        UserList(UserIndex).flags.TomoPocion = False
        UserList(UserIndex).flags.TipoPocion = 0
        Call WriteFYA(UserIndex)
        Call WriteSendSkillCdUpdate(UserIndex, CLIENT_EFFECT_TYPE_ID_POTENCIADO, DotInfo.UniqueId, 0, 0, eBuff)
    End If
End Sub

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

Public Sub IBaseEffectOverTime_TargetWillAttackPosition(ByVal Map As Integer, ByVal PosX As Integer, ByVal PosY As Integer)
End Sub

Public Function IBaseEffectOverTime_ChangeTarget(ByVal NewTargetIndex As Integer, ByVal NewTargetType As e_ReferenceType) As Boolean
    IBaseEffectOverTime_ChangeTarget = False
End Function

Public Function IBaseEffectOverTime_ApplyDamageReduction(ByVal Damage As Long, _
                                                         ByVal SourceUserId As Integer, _
                                                         ByVal SourceType As e_ReferenceType, _
                                                         ByVal AttackType As e_DamageSourceType) As Long
    IBaseEffectOverTime_ApplyDamageReduction = Damage
End Function

Public Property Get IBaseEffectOverTime_CallBacksMask() As Long
    IBaseEffectOverTime_CallBacksMask = 0
End Property

Public Property Get IBaseEffectOverTime_KeepAfterDead() As Boolean
    IBaseEffectOverTime_KeepAfterDead = False
End Property

Public Sub IBaseEffectOverTime_TargetChangeTerrain()
End Sub
