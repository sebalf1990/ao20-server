Attribute VB_Name = "modElementalCombat"
' Argentum 20 Game Server
'
' modElementalCombat
' Sistema de danos elementales unificado (TOGGLE32 elemental_system). Plan 20.002 - Ola 0.
'
' Motor de la CAPA ADITIVA sobre el combate fisico (core intacto). Provee:
'   - Registro data-driven de tipos de dano (DamageTypes.dat) sobre e_ElementalDamageType.
'   - Roll de componentes de dano tipado.
'   - Resistencia generica por clave (2 sabores): resist-a-dano (chance/flat/pct + cap + inmune)
'     y resist-a-efecto (Ola 1).
'   - Motor de procs on-hit: atacante (alGolpear) y defensor (alSerGolpeado), kind dmgBonus/applyState.
'   - Punto de entrada gateado para el camino user->NPC (unico touchpoint de la Ola 0).
'   - Log de efectos (Logs\Elemental_YYYY-MM-DD.log) y helper de inspeccion para GM.
'
' OFF (toggle en 0) = no se ejecuta NADA: combate fisico vanilla identico. No hay fallback legacy.

Option Explicit

' Registro de tipos de dano (la "clave"). Cargado al boot por LoadDamageTypes.
Public DamageTypeReg(1 To MAX_DAMAGE_TYPE_ID) As t_DamageTypeInfo
Public DamageTypeRegLoaded As Boolean

' Crit universal (plan 20.002 Ola 2). Config data-driven desde DamageTypes.dat [INIT].
Public Const DMG_TYPE_CRIT As Long = 15   ' clave de resistencia para crit (fisico+magico)
' 06.002 Ola 3: cap duro del crit universal configurable; un valor de test (50) llego al .dat vivo.
Private Const MAX_UNIVERSAL_CRIT_CHANCE_PCT As Single = 25
' 06.002 Ola 4: ids de icono del HUD (Effects.ini, data-driven por EffectCount en el cliente).
Private Const CLIENT_EFFECT_ENCHANTED_AMMO As Integer = 57
Private Const CLIENT_EFFECT_ENCHANTED_WEAPON As Integer = 58
Private mUniversalCritChance As Single
Private mUniversalCritMult As Single

' Control / cap de slow (plan 20.002 Ola 3). Cacheado al boot (hot path de velocidad).
Public ElementalSlowCapActive As Boolean
Private mElementalSlowFloor As Single

' ============================================================================
' Toggle maestro
' ============================================================================
Public Function ElementalSystemEnabled() As Boolean
    ElementalSystemEnabled = IsFeatureEnabled("elemental_system")
End Function

' ============================================================================
' Registro de tipos de dano (data-driven)
' ============================================================================
Private Sub SetDamageTypeDefault(ByVal t As e_ElementalDamageType, ByVal nm As String, ByVal col As Long)
    If t < 1 Or t > MAX_DAMAGE_TYPE_ID Then Exit Sub
    DamageTypeReg(t).nombre = nm
    DamageTypeReg(t).NumberColor = col
    DamageTypeReg(t).ResistCapPct = 0
    DamageTypeReg(t).DefaultParticle = 0
    DamageTypeReg(t).TagElementalRelacionado = 0
    DamageTypeReg(t).Defined = 1
End Sub

Public Sub InitDamageTypeRegistryDefaults()
    Dim i As Long
    For i = 1 To MAX_DAMAGE_TYPE_ID
        DamageTypeReg(i).nombre = ""
        DamageTypeReg(i).NumberColor = vbWhite
        DamageTypeReg(i).DefaultParticle = 0
        DamageTypeReg(i).ResistCapPct = 0
        DamageTypeReg(i).TagElementalRelacionado = 0
        DamageTypeReg(i).Defined = 0
    Next i
    ' Built-in (sirven aun sin DamageTypes.dat). Luz/Oscuridad reservados al oficial.
    Call SetDamageTypeDefault(eDmgPhysical, "Fisico", vbWhite)
    Call SetDamageTypeDefault(eDmgFire, "Fuego", RGB(255, 140, 0))
    Call SetDamageTypeDefault(eDmgFrost, "Escarcha", RGB(120, 200, 255))
    Call SetDamageTypeDefault(eDmgPoison, "Veneno", RGB(120, 200, 60))
    Call SetDamageTypeDefault(eDmgAcid, "Acido", RGB(170, 230, 40))
    Call SetDamageTypeDefault(eDmgArcane, "Arcano", RGB(200, 120, 255))
    Call SetDamageTypeDefault(eDmgBleed, "Sangrado", RGB(200, 30, 30))
End Sub

' Carga DamageTypes.dat (metadata tuneable). Tolerante: si no existe, usa los defaults.
Public Sub LoadDamageTypes()
    On Error GoTo ErrHandler
    Call InitDamageTypeRegistryDefaults
    DamageTypeRegLoaded = True
    mUniversalCritChance = 5
    mUniversalCritMult = 0.5
    DamageTypeReg(DMG_TYPE_CRIT).nombre = "Critico"
    DamageTypeReg(DMG_TYPE_CRIT).NumberColor = RGB(255, 165, 0)
    ElementalSlowCapActive = ElementalSystemEnabled()
    mElementalSlowFloor = 0.45
    If Not ElementalSystemEnabled() Then Exit Sub
    Dim fname As String
    fname = DatPath & "DamageTypes.dat"
    If LenB(dir(fname)) = 0 Then Exit Sub
    Dim Leer As New clsIniManager
    Call Leer.Initialize(fname)
    ' 06.002 Ola 3: clamps de carga. Un CritChancePct de test (50) llego a produccion en el .dat;
    ' el guardrail evita que un valor fuera de rango entre en silencio.
    If val(Leer.GetValue("INIT", "CritChancePct")) > 0 Then mUniversalCritChance = val(Leer.GetValue("INIT", "CritChancePct"))
    If mUniversalCritChance > MAX_UNIVERSAL_CRIT_CHANCE_PCT Then
        Call LogThis(0, "[DamageTypes] CritChancePct=" & mUniversalCritChance & " fuera de rango, clampeado a " & MAX_UNIVERSAL_CRIT_CHANCE_PCT, vbLogEventTypeWarning)
        mUniversalCritChance = MAX_UNIVERSAL_CRIT_CHANCE_PCT
    End If
    If val(Leer.GetValue("INIT", "CritMult")) > 0 Then mUniversalCritMult = val(Leer.GetValue("INIT", "CritMult"))
    If val(Leer.GetValue("INIT", "SlowFloorPct")) > 0 Then mElementalSlowFloor = val(Leer.GetValue("INIT", "SlowFloorPct"))
    If mElementalSlowFloor > 1 Then
        ' Escrito como porcentaje (45) en vez de fraccion (0.45): interpretarlo, no congelar el juego.
        Call LogThis(0, "[DamageTypes] SlowFloorPct=" & mElementalSlowFloor & " interpretado como porcentaje -> " & (mElementalSlowFloor / 100), vbLogEventTypeWarning)
        mElementalSlowFloor = mElementalSlowFloor / 100
    End If
    Dim count As Long, i As Long, sect As String, idv As Long
    count = val(Leer.GetValue("INIT", "TypeCount"))
    For i = 1 To count
        sect = "TYPE" & i
        idv = val(Leer.GetValue(sect, "Id"))
        If idv >= 1 And idv <= MAX_DAMAGE_TYPE_ID Then
            With DamageTypeReg(idv)
                If LenB(Leer.GetValue(sect, "Name")) > 0 Then .nombre = Leer.GetValue(sect, "Name")
                .NumberColor = val(Leer.GetValue(sect, "NumberColor"))
                .DefaultParticle = val(Leer.GetValue(sect, "DefaultParticle"))
                .ResistCapPct = val(Leer.GetValue(sect, "ResistCapPct"))
                .TagElementalRelacionado = val(Leer.GetValue(sect, "TagElementalRelacionado"))
                .Defined = 1
            End With
        End If
    Next i
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.LoadDamageTypes", Erl)
End Sub

Public Function DamageTypeName(ByVal t As e_ElementalDamageType) As String
    If t >= 1 And t <= MAX_DAMAGE_TYPE_ID Then
        If LenB(DamageTypeReg(t).nombre) > 0 Then
            DamageTypeName = DamageTypeReg(t).nombre
            Exit Function
        End If
    End If
    DamageTypeName = "Tipo" & CStr(t)
End Function

Public Function DamageTypeColor(ByVal t As e_ElementalDamageType) As Long
    If t >= 1 And t <= MAX_DAMAGE_TYPE_ID Then DamageTypeColor = DamageTypeReg(t).NumberColor
    If DamageTypeColor = 0 Then DamageTypeColor = vbWhite
End Function

' Plan 20.002 Checkpoint B: particula de impacto del tipo (DamageTypes.dat DefaultParticle). 0 = ninguna.
Public Function DamageTypeDefaultParticle(ByVal t As e_ElementalDamageType) As Integer
    If t >= 1 And t <= MAX_DAMAGE_TYPE_ID Then DamageTypeDefaultParticle = DamageTypeReg(t).DefaultParticle
End Function

' Nombre del tipo elemental cuyo color de numero coincide (para mensajes de consola, plan 20.002 TP4).
' Devuelve "" si el color no corresponde a ningun tipo (ej dano fisico) -> el call-site usa el mensaje generico.
Public Function DamageTypeNameFromColor(ByVal numberColor As Long) As String
    If numberColor = 0 Or numberColor = vbWhite Then Exit Function
    Dim t As Integer
    For t = 1 To MAX_DAMAGE_TYPE_ID
        If DamageTypeReg(t).NumberColor = numberColor Then
            DamageTypeNameFromColor = DamageTypeReg(t).nombre
            Exit Function
        End If
    Next t
End Function

Private Function ResistCapForType(ByVal t As e_ElementalDamageType) As Single
    Dim cap As Single
    If t >= 1 And t <= MAX_DAMAGE_TYPE_ID Then cap = DamageTypeReg(t).ResistCapPct
    If cap <= 0 Then cap = ELEMENTAL_RESIST_CAP_DEFAULT
    ResistCapForType = cap
End Function

' ============================================================================
' Roll de dano
' ============================================================================
Public Function RollDamageComponent(ByRef c As t_DamageComponent) As Long
    Dim lo As Long, hi As Long
    lo = c.MinDamage
    hi = c.MaxDamage
    If lo < 0 Then lo = 0
    If hi < lo Then hi = lo
    If hi <= lo Then
        RollDamageComponent = lo
    Else
        RollDamageComponent = RandomNumber(lo, hi)
    End If
End Function

' ============================================================================
' Receta unica del roll de un tick de DoT enriquecido (compartida por el veneno
' y por los DoT genericos de UpdateHpOverTime). Centraliza la cuenta del dano base.
' modo: 0=fijo (Min) | 1=rango (Random Min..Max) | 2=%HP fijo | 3=%HP rango.
' Preserva el signo de los inputs: el veneno pasa valores positivos y niega al
' aplicar; los DoT genericos pasan valores negativos y aplican directo. NO aplica
' factores PvP/PvE, EffectBonus, stacks ni reducciones: eso queda en cada motor.
' ============================================================================
Public Function RollDotTickBase(ByVal TargetIndex As Integer, _
                                ByVal TargetType As e_ReferenceType, _
                                ByVal DanoModo As Byte, _
                                ByVal DanoMin As Long, _
                                ByVal DanoMax As Long) As Long
    Dim maxHpT As Long
    Select Case DanoModo
        Case 0 ' fijo
            RollDotTickBase = DanoMin
        Case 1 ' rango
            RollDotTickBase = RandomNumber(DanoMin, DanoMax)
        Case 2 ' pct_hp fijo
            maxHpT = GetDotTargetMaxHp(TargetIndex, TargetType)
            RollDotTickBase = (CLng(DanoMin) * maxHpT) \ 100
        Case 3 ' pct_hp rango
            maxHpT = GetDotTargetMaxHp(TargetIndex, TargetType)
            RollDotTickBase = (CLng(RandomNumber(DanoMin, DanoMax)) * maxHpT) \ 100
    End Select
End Function

Private Function GetDotTargetMaxHp(ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType) As Long
    If TargetType = e_ReferenceType.eUser Then
        GetDotTargetMaxHp = UserList(TargetIndex).Stats.MaxHp
    ElseIf TargetType = e_ReferenceType.eNpc Then
        GetDotTargetMaxHp = NpcList(TargetIndex).Stats.MaxHp
    End If
End Function

' ============================================================================
' Resistencia: agregacion por entidad + aplicacion con cap
' ============================================================================
Private Sub AddResistEntry(ByRef acc As t_ElementalResist, ByRef e As t_ElementalResist)
    acc.ReduceChancePct = acc.ReduceChancePct + e.ReduceChancePct
    acc.ReduceFlat = acc.ReduceFlat + e.ReduceFlat
    acc.ReducePct = acc.ReducePct + e.ReducePct
    acc.ReduceFlatMin = acc.ReduceFlatMin + e.ReduceFlatMin
    acc.ReduceFlatMax = acc.ReduceFlatMax + e.ReduceFlatMax
    acc.ReducePctMin = acc.ReducePctMin + e.ReducePctMin
    acc.ReducePctMax = acc.ReducePctMax + e.ReducePctMax
    If e.Immune <> 0 Then acc.Immune = 1
    If e.ImmuneDamage <> 0 Then acc.ImmuneDamage = 1
    If e.ImmuneEffect <> 0 Then acc.ImmuneEffect = 1
    ' 06.002 Ola 2: EffMag/EffDur retirados del esquema (se parseaban y acumulaban pero ningun
    ' consumidor los aplicaba y ningun .dat los usa; implementarlos es rediseno de la creacion de EOT).
    acc.ReduceEffectChancePct = acc.ReduceEffectChancePct + e.ReduceEffectChancePct
End Sub

Private Sub AddSetResist(ByRef acc As t_ElementalResist, ByRef rs As t_ElementalResistSet, ByVal dmgType As e_ElementalDamageType)
    Dim k As Integer
    For k = 1 To rs.Count
        If rs.Resist(k).DamageType = dmgType Then
            Call AddResistEntry(acc, rs.Resist(k))
        End If
    Next k
End Sub

Private Sub AddSlotElementalResist(ByRef acc As t_ElementalResist, ByVal ObjIndex As Integer, ByVal dmgType As e_ElementalDamageType)
    ' Robustez: un slot de equipo vacio es ObjIndex 0; evitar ObjData(0) (out of range con user-target).
    If ObjIndex > 0 Then Call AddSetResist(acc, ObjData(ObjIndex).ElementalResist, dmgType)
End Sub

Public Function GetUserElementalResist(ByVal UserIndex As Integer, ByVal dmgType As e_ElementalDamageType) As t_ElementalResist
    Dim acc As t_ElementalResist
    acc.DamageType = dmgType
    If UserIndex > 0 Then
        With UserList(UserIndex).invent
            Call AddSlotElementalResist(acc, .EquippedArmorObjIndex, dmgType)
            Call AddSlotElementalResist(acc, .EquippedHelmetObjIndex, dmgType)
            Call AddSlotElementalResist(acc, .EquippedShieldObjIndex, dmgType)
            Call AddSlotElementalResist(acc, .EquippedRingAccesoryObjIndex, dmgType)
            Call AddSlotElementalResist(acc, .EquippedAmuletAccesoryObjIndex, dmgType)
            Call AddSlotElementalResist(acc, .EquippedBackpackObjIndex, dmgType)
            Call AddSlotElementalResist(acc, .EquippedSaddleObjIndex, dmgType)
        End With
    End If
    GetUserElementalResist = acc
End Function

Public Function GetNpcElementalResist(ByVal NpcIndex As Integer, ByVal dmgType As e_ElementalDamageType) As t_ElementalResist
    Dim acc As t_ElementalResist
    acc.DamageType = dmgType
    If NpcIndex > 0 Then
        Dim t As Integer
        t = NpcList(NpcIndex).Numero
        If t > 0 Then
            Call AddSetResist(acc, NpcInfoCache(t).ElementalResist, dmgType)
        End If
    End If
    GetNpcElementalResist = acc
End Function

Private Function GetTargetResist(ByVal targetIsNpc As Boolean, ByVal targetIndex As Integer, ByVal dmgType As e_ElementalDamageType) As t_ElementalResist
    If targetIsNpc Then
        GetTargetResist = GetNpcElementalResist(targetIndex, dmgType)
    Else
        GetTargetResist = GetUserElementalResist(targetIndex, dmgType)
    End If
End Function

' Aplica la cascada de resistencia-a-dano a un numero. Devuelve el dano final (>=0).
Public Function ApplyElementalResist(ByVal RawDamage As Long, ByRef r As t_ElementalResist, ByVal dmgType As e_ElementalDamageType, ByRef outNullified As Boolean) As Long
    outNullified = False
    If RawDamage <= 0 Then Exit Function
    ' Inmunidad al daño: bloqueo absoluto (Immune legacy = ambos; ImmuneDamage = solo daño)
    If r.Immune <> 0 Or r.ImmuneDamage <> 0 Then
        outNullified = True
        Exit Function
    End If
    ' Chance de anular el componente entero
    If r.ReduceChancePct > 0 Then
        Dim ch As Long
        ch = r.ReduceChancePct
        If ch > 100 Then ch = 100
        If RandomNumber(1, 100) <= ch Then
            outNullified = True
            Exit Function
        End If
    End If
    ' Cascada: flat -> pct (capeado). Rango aleatorio si *Max>0 (Step 4: preserva el veneno).
    Dim dmg As Long
    Dim flat As Long
    If r.ReduceFlatMax > 0 Then
        If r.ReduceFlatMin >= r.ReduceFlatMax Then
            flat = r.ReduceFlatMin
        Else
            flat = RandomNumber(r.ReduceFlatMin, r.ReduceFlatMax)
        End If
    Else
        flat = r.ReduceFlat
    End If
    dmg = RawDamage - flat
    If dmg <= 0 Then Exit Function
    Dim pct As Single
    If r.ReducePctMax > 0 Then
        Dim pctInt As Long
        If r.ReducePctMin >= r.ReducePctMax Then
            pctInt = r.ReducePctMin
        Else
            pctInt = RandomNumber(r.ReducePctMin, r.ReducePctMax)
        End If
        pct = pctInt / 100#
    Else
        pct = r.ReducePct
    End If
    Dim cap As Single
    cap = ResistCapForType(dmgType)
    If pct > cap Then pct = cap
    If pct < 0 Then pct = 0
    ' Plan 10.001 decision 21: redondeo probabilistico en vez de truncar con Int().
    ' Int() sola casi siempre trunca para abajo (medido: 6% nominal daba 26.5% real
    ' en componentes chicos). La parte decimal perdida se convierte en la chance de
    ' sumar 1, asi el promedio queda en el pct nominal para cualquier tamano de dano.
    ' pct=0 -> scaled=dmg exacto, frac=0, no se tira dado y dmg queda intacto.
    Dim scaled As Double
    scaled = CDbl(dmg) * (1# - pct)
    dmg = Int(scaled)
    Dim frac As Double
    frac = scaled - dmg
    If frac > 0 Then
        Dim roll As Long
        roll = RandomNumber(1, 10000)
        If roll <= CLng(frac * 10000) Then dmg = dmg + 1
    End If
    If dmg < 0 Then dmg = 0
    ApplyElementalResist = dmg
End Function

' ============================================================================
' M6 (plan 02.001): resist/inmunidad al dano para el tick de un DoT generico
' (UpdateHpOverTime.cls) cuando EffectsOverTime.dat declara DamageType>0. Antes
' ningun tick de Quemadura/Sangrado/etc consultaba Immune/ImmuneDamage: un
' blanco inmune al dano de Fuego igual perdia HP por los ticks de Quemadura.
' Las clases de veneno (modPoisonResist) NO pasan por aca, tienen su propio
' camino y no se tocan. Gateado por el toggle maestro: OFF = rawDamage intacto.
' ============================================================================
Public Function ApplyDotTickResist(ByVal targetIsNpc As Boolean, ByVal targetIndex As Integer, ByVal dmgType As Integer, ByVal rawDamage As Long) As Long
    On Error GoTo eh
    ApplyDotTickResist = rawDamage
    If Not ElementalSystemEnabled() Then Exit Function
    If dmgType <= 0 Then Exit Function
    If rawDamage <= 0 Then Exit Function
    Dim r As t_ElementalResist
    r = GetTargetResist(targetIsNpc, targetIndex, dmgType)
    Dim nul As Boolean
    Dim dmg As Long
    dmg = ApplyElementalResist(rawDamage, r, dmgType, nul)
    ApplyDotTickResist = dmg
    Call ElementalLog("DoT tick resist target=" & IIf(targetIsNpc, "N", "U") & targetIndex & " type=" & dmgType & " raw=" & rawDamage & " final=" & dmg & " nullified=" & nul)
    If ElementalBalanceLogEnabled() Then Call LogElementalBalance("dot_tick_resist", "0", "0", ElementalBalanceActorId(targetIsNpc, targetIndex), IIf(targetIsNpc, "npc", "user"), 0, 0, dmgType, rawDamage, dmg, ElementalBalanceMap(targetIsNpc, targetIndex), 0, 0, ElementalBalanceCharId(targetIsNpc, targetIndex), 0, 0)
    Exit Function
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.ApplyDotTickResist", Erl)
    ApplyDotTickResist = rawDamage
End Function

' ============================================================================
' Resolucion de componentes y procs
' ============================================================================
' Plan 20.002 Checkpoint B: dispara el burst de impacto del tipo sobre el target (si DefaultParticle>0).
Private Sub SendImpactParticle(ByVal targetIsNpc As Boolean, ByVal targetIndex As Integer, ByVal damageType As e_ElementalDamageType)
    Const FxTime As Long = 30 ' tunable: duracion del burst de impacto
    Dim part As Integer
    part = DamageTypeDefaultParticle(damageType)
    If part <= 0 Then Exit Sub
    If targetIndex <= 0 Then Exit Sub
    If targetIsNpc Then
        Call SendData(SendTarget.ToNPCAliveArea, targetIndex, PrepareMessageParticleFX(NpcList(targetIndex).Char.charindex, part, FxTime, False, , NpcList(targetIndex).pos.x, NpcList(targetIndex).pos.y))
    Else
        Call SendData(SendTarget.ToPCAliveArea, targetIndex, PrepareMessageParticleFX(UserList(targetIndex).Char.charindex, part, FxTime, False, , UserList(targetIndex).pos.x, UserList(targetIndex).pos.y))
    End If
End Sub

Private Function ResolveComponentsVsTarget(ByRef src As t_ElementalSource, ByVal targetIsNpc As Boolean, ByVal targetIndex As Integer, ByVal attackerIndex As Integer, ByVal attackerType As e_ReferenceType, ByVal itemObjIndex As Integer, ByVal srcItemObjIndex As Integer, ByVal logCtx As String) As Long
    Dim total As Long, i As Integer
    For i = 1 To src.CompCount
        If src.Comp(i).DamageType >= 1 And src.Comp(i).DamageType <= MAX_DAMAGE_TYPE_ID Then
            Dim raw As Long
            raw = RollDamageComponent(src.Comp(i))
            Dim r As t_ElementalResist
            r = GetTargetResist(targetIsNpc, targetIndex, src.Comp(i).DamageType)
            Dim nullified As Boolean, finalDmg As Long
            finalDmg = ApplyElementalResist(raw, r, src.Comp(i).DamageType, nullified)
            total = total + finalDmg
            If finalDmg > 0 Then Call SendImpactParticle(targetIsNpc, targetIndex, src.Comp(i).DamageType)
            Call ElementalLog(logCtx & " comp " & DamageTypeName(src.Comp(i).DamageType) & " raw=" & raw & " final=" & finalDmg & IIf(nullified, " [nullified]", "") & IIf(r.Immune <> 0, " [immune]", ""))
            If ElementalBalanceLogEnabled() Then Call LogElementalBalance("component", ElementalBalanceActorId(attackerType = eNpc, attackerIndex), ElementalBalanceActorClass(attackerType = eNpc, attackerIndex), ElementalBalanceActorId(targetIsNpc, targetIndex), IIf(targetIsNpc, "npc", "user"), itemObjIndex, ElementalBalanceCatalogTier(itemObjIndex), src.Comp(i).DamageType, raw, finalDmg, ElementalBalanceMap(targetIsNpc, targetIndex), ElementalBalanceCharId(attackerType = eNpc, attackerIndex), ElementalBalanceAccountId(attackerType = eNpc, attackerIndex), ElementalBalanceCharId(targetIsNpc, targetIndex), srcItemObjIndex, ElementalBalanceFightId(ElementalBalanceCharId(attackerType = eNpc, attackerIndex), ElementalBalanceCharId(targetIsNpc, targetIndex)))
        End If
    Next i
    ResolveComponentsVsTarget = total
End Function

' Dispara los procs de una fuente con el trigger dado contra un target.
' Devuelve el dano extra (kind=dmgBonus) ya reducido por la resistencia del target.
Private Function FireProcs(ByRef src As t_ElementalSource, ByVal trig As e_ProcTrigger, ByVal targetIsNpc As Boolean, ByVal targetIndex As Integer, ByVal attackerIndex As Integer, ByVal attackerType As e_ReferenceType, ByVal itemObjIndex As Integer, ByVal srcItemObjIndex As Integer, ByVal logCtx As String, Optional ByVal onlyKind As Long = -1) As Long
    Dim total As Long, i As Integer
    For i = 1 To src.ProcCount
        If src.Proc(i).Trigger = trig And (onlyKind = -1 Or src.Proc(i).Kind = onlyKind) Then
            Dim ch As Long
            ch = src.Proc(i).ChancePct
            If ch < 0 Then ch = 0
            If ch > 100 Then ch = 100
            If ch >= 100 Or RandomNumber(1, 100) <= ch Then
                Select Case src.Proc(i).Kind
                    Case eProcDamageBonus
                        Dim c As t_DamageComponent
                        c.DamageType = src.Proc(i).DamageType
                        c.MinDamage = src.Proc(i).MinDamage
                        c.MaxDamage = src.Proc(i).MaxDamage
                        Dim raw As Long
                        raw = RollDamageComponent(c)
                        Dim r As t_ElementalResist
                        r = GetTargetResist(targetIsNpc, targetIndex, c.DamageType)
                        Dim nul As Boolean, fd As Long
                        fd = ApplyElementalResist(raw, r, c.DamageType, nul)
                        total = total + fd
                        Call ElementalLog(logCtx & " PROC dmgBonus " & DamageTypeName(c.DamageType) & " final=" & fd)
                        If ElementalBalanceLogEnabled() Then Call LogElementalBalance("proc_dmg_bonus", ElementalBalanceActorId(attackerType = eNpc, attackerIndex), ElementalBalanceActorClass(attackerType = eNpc, attackerIndex), ElementalBalanceActorId(targetIsNpc, targetIndex), IIf(targetIsNpc, "npc", "user"), itemObjIndex, ElementalBalanceCatalogTier(itemObjIndex), c.DamageType, raw, fd, ElementalBalanceMap(targetIsNpc, targetIndex), ElementalBalanceCharId(attackerType = eNpc, attackerIndex), ElementalBalanceAccountId(attackerType = eNpc, attackerIndex), ElementalBalanceCharId(targetIsNpc, targetIndex), srcItemObjIndex, ElementalBalanceFightId(ElementalBalanceCharId(attackerType = eNpc, attackerIndex), ElementalBalanceCharId(targetIsNpc, targetIndex)))
                    Case eProcApplyState
                        ' Aplica el preset (EotId) respetando inmunidad / resist-a-efecto del tipo del proc.
                        ' 06.002 Ola 3: EotId fuera del catalogo (dat mal tipeado, ej ElemProc1Eot=999) tiraba
                        ' error 9 y abortaba TODO el dano elemental del golpe. Se ignora solo ese proc.
                        If src.Proc(i).EotId > UBound(EffectOverTime) Then
                            Call ElementalLog(logCtx & " PROC applyState EotId=" & src.Proc(i).EotId & " FUERA DE CATALOGO (ignorado)")
                        ElseIf src.Proc(i).EotId > 0 Then
                            Dim rr As t_ElementalResist
                            rr = GetTargetResist(targetIsNpc, targetIndex, src.Proc(i).DamageType)
                            If rr.Immune <> 0 Or rr.ImmuneEffect <> 0 Then
                                Call ElementalLog(logCtx & " PROC applyState BLOQUEADO (inmune " & DamageTypeName(src.Proc(i).DamageType) & ")")
                            ElseIf rr.ReduceEffectChancePct > 0 And RandomNumber(1, 100) <= rr.ReduceEffectChancePct Then
                                Call ElementalLog(logCtx & " PROC applyState resistido (chance efecto)")
                            Else
                                Dim trt As e_ReferenceType
                                If targetIsNpc Then trt = eNpc Else trt = eUser
                                ' Limit-aware: si ya existe (segun Limit), Reset (refresca/suma stack); si no, crear.
                                Dim existEff As IBaseEffectOverTime
                                If targetIsNpc Then
                                    Set existEff = EffectsOverTime.FindEffectOnTarget(attackerIndex, NpcList(targetIndex).EffectOverTime, src.Proc(i).EotId)
                                Else
                                    Set existEff = EffectsOverTime.FindEffectOnTarget(attackerIndex, UserList(targetIndex).EffectOverTime, src.Proc(i).EotId)
                                End If
                                If existEff Is Nothing Then
                                    Call EffectsOverTime.CreateEffect(attackerIndex, attackerType, targetIndex, trt, src.Proc(i).EotId)
                                Else
                                    Call existEff.Reset(attackerIndex, attackerType, src.Proc(i).EotId)
                                End If
                                If attackerType = eUser And LenB(EffectOverTime(src.Proc(i).EotId).ApplyMsg) > 0 Then Call WriteConsoleMsg(attackerIndex, EffectOverTime(src.Proc(i).EotId).ApplyMsg, e_TextChannel.TEXTCHANNEL_COMBAT, e_FontTypeNames.FONTTYPE_FIGHT)
                                Call ElementalLog(logCtx & " PROC applyState EotId=" & src.Proc(i).EotId & IIf(existEff Is Nothing, " aplicado", " refrescado"))
                                If ElementalBalanceLogEnabled() Then Call LogElementalBalance("proc_apply_state", ElementalBalanceActorId(attackerType = eNpc, attackerIndex), ElementalBalanceActorClass(attackerType = eNpc, attackerIndex), ElementalBalanceActorId(targetIsNpc, targetIndex), IIf(targetIsNpc, "npc", "user"), itemObjIndex, ElementalBalanceCatalogTier(itemObjIndex), src.Proc(i).DamageType, 0, 0, ElementalBalanceMap(targetIsNpc, targetIndex), ElementalBalanceCharId(attackerType = eNpc, attackerIndex), ElementalBalanceAccountId(attackerType = eNpc, attackerIndex), ElementalBalanceCharId(targetIsNpc, targetIndex), srcItemObjIndex, ElementalBalanceFightId(ElementalBalanceCharId(attackerType = eNpc, attackerIndex), ElementalBalanceCharId(targetIsNpc, targetIndex)))
                            End If
                        Else
                            Call ElementalLog(logCtx & " PROC applyState sin EotId (ignorado)")
                        End If
                End Select
            End If
        End If
    Next i
    FireProcs = total
End Function

' Tipo elemental cuyo color de numero coincide (para reflejo-espejo de espinas). eDmgNone si no matchea/fisico.
Public Function DamageTypeFromColor(ByVal numberColor As Long) As e_ElementalDamageType
    DamageTypeFromColor = eDmgNone
    If numberColor = 0 Or numberColor = vbWhite Then Exit Function
    Dim t As Integer
    For t = 1 To MAX_DAMAGE_TYPE_ID
        If DamageTypeReg(t).NumberColor = numberColor Then
            DamageTypeFromColor = t
            Exit Function
        End If
    Next t
End Function

' === Espinas / Thorns (plan 20.002) ===
' El DEFENSOR retalia al ATACANTE tras recibir un golpe (se llama POST-dano para conocer el dano neto).
' Maneja procs onDamaged Kind=dmgBonus con el modelo configurable: fijo (Min-Max) o reflejo % de netDamage;
' fisico o elemental; directo o resistido (elemental=resist del tipo; fisico-resistido=TP3); letal o no.
' El applyState onDamaged (efectos) lo sigue disparando FireProcs en el on-hit. Anti-loop: aplica por
' DoDamageOrHeal directo (e_dot), no re-dispara espinas.
Public Sub ResolveThorns(ByRef defenderSrc As t_ElementalSource, ByVal defenderIndex As Integer, ByVal defenderType As e_ReferenceType, ByVal attackerIsNpc As Boolean, ByVal attackerIndex As Integer, ByVal netDamage As Long, ByVal attackType As e_ElementalDamageType, ByVal logCtx As String, Optional ByVal slotObjIndex As Integer = 0)
    On Error GoTo ErrHandler
    If Not ElementalSystemEnabled() Then Exit Sub
    If attackerIndex <= 0 Then Exit Sub
    Dim i As Integer
    For i = 1 To defenderSrc.ProcCount
        If defenderSrc.Proc(i).Trigger = eProcOnDamaged And defenderSrc.Proc(i).Kind = eProcDamageBonus Then
            Dim ch As Long
            ch = defenderSrc.Proc(i).ChancePct
            If ch < 0 Then ch = 0
            If ch > 100 Then ch = 100
            If ch >= 100 Or RandomNumber(1, 100) <= ch Then
                ' Tipo/naturaleza efectivos. ReflectMirror (espejo): el dano devuelto toma el TIPO del golpe del
                ' atacante; si el golpe no tuvo componente elemental (attackType=eDmgNone) la espina no actua.
                Dim effType As e_ElementalDamageType, effPhysical As Byte, doMirror As Boolean
                doMirror = True
                If defenderSrc.Proc(i).ReflectMirror = 1 Then
                    If attackType = eDmgNone Then
                        doMirror = False
                    Else
                        effType = attackType
                        effPhysical = 0
                    End If
                Else
                    effType = defenderSrc.Proc(i).DamageType
                    effPhysical = defenderSrc.Proc(i).Physical
                End If
                If doMirror Then
                    Dim dmg As Long
                    If defenderSrc.Proc(i).ReflectPct > 0 Then
                        dmg = (netDamage * defenderSrc.Proc(i).ReflectPct) \ 100
                    Else
                        Dim c As t_DamageComponent
                        c.DamageType = effType
                        c.MinDamage = defenderSrc.Proc(i).MinDamage
                        c.MaxDamage = defenderSrc.Proc(i).MaxDamage
                        dmg = RollDamageComponent(c)
                    End If
                    If defenderSrc.Proc(i).BypassResist = 0 Then
                        If effPhysical = 0 Then
                            Dim rThorn As t_ElementalResist, nulT As Boolean
                            rThorn = GetTargetResist(attackerIsNpc, attackerIndex, effType)
                            dmg = ApplyElementalResist(dmg, rThorn, effType, nulT)
                        Else
                            If attackerIsNpc Then
                                dmg = dmg * NPCs.GetPhysicDamageReduction(NpcList(attackerIndex))
                            Else
                                dmg = dmg * UserMod.GetPhysicDamageReduction(UserList(attackerIndex))
                            End If
                        End If
                    End If
                    If dmg > 0 Then Call ApplyThornsDamage(attackerIsNpc, attackerIndex, defenderIndex, defenderType, dmg, defenderSrc.Proc(i).Lethal, effPhysical, effType, slotObjIndex, logCtx)
                End If
            End If
        End If
    Next i
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.ResolveThorns", Erl)
End Sub

' Aplica el dano de espina al ATACANTE (user o NPC). Letalidad: si lethal=0, cap a HP-1. Directo (e_dot).
Private Sub ApplyThornsDamage(ByVal attackerIsNpc As Boolean, ByVal attackerIndex As Integer, ByVal defenderIndex As Integer, ByVal defenderType As e_ReferenceType, ByVal dmg As Long, ByVal lethal As Byte, ByVal physical As Byte, ByVal dmgType As e_ElementalDamageType, ByVal slotObjIndex As Integer, ByVal logCtx As String)
    On Error GoTo ErrHandler
    Dim col As Long
    If physical = 1 Then col = vbRed Else col = DamageTypeColor(dmgType)
    If attackerIsNpc Then
        If lethal = 0 Then
            If dmg >= NpcList(attackerIndex).Stats.MinHp Then dmg = NpcList(attackerIndex).Stats.MinHp - 1
        End If
        If dmg > 0 Then Call NPCs.DoDamageOrHeal(attackerIndex, defenderIndex, defenderType, -dmg, e_dot, 0, col)
    Else
        If lethal = 0 Then
            If dmg >= UserList(attackerIndex).Stats.MinHp Then dmg = UserList(attackerIndex).Stats.MinHp - 1
        End If
        If dmg > 0 Then Call UserMod.DoDamageOrHeal(attackerIndex, defenderIndex, defenderType, -dmg, e_dot, 0, , , col)
    End If
    If dmg > 0 Then Call ElementalLog(logCtx & " THORNS dmg=" & dmg & IIf(physical = 1, " fisico", " " & DamageTypeName(dmgType)) & IIf(lethal = 0, " (no letal)", ""))
    If dmg > 0 And ElementalBalanceLogEnabled() Then Call LogElementalBalance("thorns", ElementalBalanceActorId(defenderType = eNpc, defenderIndex), ElementalBalanceActorClass(defenderType = eNpc, defenderIndex), ElementalBalanceActorId(attackerIsNpc, attackerIndex), IIf(attackerIsNpc, "npc", "user"), slotObjIndex, ElementalBalanceCatalogTier(slotObjIndex), IIf(physical = 1, 0, dmgType), 0, dmg, ElementalBalanceMap(attackerIsNpc, attackerIndex), ElementalBalanceCharId(defenderType = eNpc, defenderIndex), ElementalBalanceAccountId(defenderType = eNpc, defenderIndex), ElementalBalanceCharId(attackerIsNpc, attackerIndex), 0, ElementalBalanceFightId(ElementalBalanceCharId(defenderType = eNpc, defenderIndex), ElementalBalanceCharId(attackerIsNpc, attackerIndex)))
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.ApplyThornsDamage", Erl)
End Sub

' Espinas del USER defensor (plan 20.002 TP2): recorre los 7 slots de gear y dispara los procs onDamaged
' de cada item equipado (espejo de GetUserElementalResist). Cada item contribuye sus espinas.
Public Sub FireUserThorns(ByVal defenderUserIndex As Integer, ByVal attackerIsNpc As Boolean, ByVal attackerIndex As Integer, ByVal netDamage As Long, ByVal attackType As e_ElementalDamageType, ByVal logCtx As String)
    On Error GoTo ErrHandler
    If Not ElementalSystemEnabled() Then Exit Sub
    If defenderUserIndex <= 0 Or attackerIndex <= 0 Or netDamage <= 0 Then Exit Sub
    With UserList(defenderUserIndex).invent
        Call FireSlotThorns(.EquippedArmorObjIndex, defenderUserIndex, attackerIsNpc, attackerIndex, netDamage, attackType, logCtx)
        Call FireSlotThorns(.EquippedHelmetObjIndex, defenderUserIndex, attackerIsNpc, attackerIndex, netDamage, attackType, logCtx)
        Call FireSlotThorns(.EquippedShieldObjIndex, defenderUserIndex, attackerIsNpc, attackerIndex, netDamage, attackType, logCtx)
        Call FireSlotThorns(.EquippedRingAccesoryObjIndex, defenderUserIndex, attackerIsNpc, attackerIndex, netDamage, attackType, logCtx)
        Call FireSlotThorns(.EquippedAmuletAccesoryObjIndex, defenderUserIndex, attackerIsNpc, attackerIndex, netDamage, attackType, logCtx)
        Call FireSlotThorns(.EquippedBackpackObjIndex, defenderUserIndex, attackerIsNpc, attackerIndex, netDamage, attackType, logCtx)
        Call FireSlotThorns(.EquippedSaddleObjIndex, defenderUserIndex, attackerIsNpc, attackerIndex, netDamage, attackType, logCtx)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.FireUserThorns", Erl)
End Sub

Private Sub FireSlotThorns(ByVal ObjIndex As Integer, ByVal defenderUserIndex As Integer, ByVal attackerIsNpc As Boolean, ByVal attackerIndex As Integer, ByVal netDamage As Long, ByVal attackType As e_ElementalDamageType, ByVal logCtx As String)
    If ObjIndex > 0 Then Call ResolveThorns(ObjData(ObjIndex).Elemental, defenderUserIndex, eUser, attackerIsNpc, attackerIndex, netDamage, attackType, logCtx, slotObjIndex:=ObjIndex)
End Sub

' 06.002 Ola 2: procs onDamaged de tipo applyState del defensor contra su atacante, POST-dano y
' solo si el defensor sobrevivio (los dmgBonus onDamaged son las espinas: ResolveThorns). Antes el
' applyState del NPC se disparaba pre-dano sin gate y el del gear de user no se disparaba nunca.
Public Sub FireOnDamagedStates(ByRef defenderSrc As t_ElementalSource, ByVal defenderIndex As Integer, ByVal defenderType As e_ReferenceType, ByVal attackerIsNpc As Boolean, ByVal attackerIndex As Integer, ByVal logCtx As String, Optional ByVal slotObjIndex As Integer = 0)
    On Error GoTo ErrHandler
    If Not ElementalSystemEnabled() Then Exit Sub
    If attackerIndex <= 0 Then Exit Sub
    Call FireProcs(defenderSrc, eProcOnDamaged, attackerIsNpc, attackerIndex, defenderIndex, defenderType, slotObjIndex, 0, logCtx & " onDmgState", eProcApplyState)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.FireOnDamagedStates", Erl)
End Sub

' Espejo de FireUserThorns para los applyState: recorre los 7 slots de gear del user defensor.
Public Sub FireUserOnDamagedStates(ByVal defenderUserIndex As Integer, ByVal attackerIsNpc As Boolean, ByVal attackerIndex As Integer, ByVal logCtx As String)
    On Error GoTo ErrHandler
    If Not ElementalSystemEnabled() Then Exit Sub
    If defenderUserIndex <= 0 Or attackerIndex <= 0 Then Exit Sub
    With UserList(defenderUserIndex).invent
        Call FireSlotOnDamagedStates(.EquippedArmorObjIndex, defenderUserIndex, attackerIsNpc, attackerIndex, logCtx)
        Call FireSlotOnDamagedStates(.EquippedHelmetObjIndex, defenderUserIndex, attackerIsNpc, attackerIndex, logCtx)
        Call FireSlotOnDamagedStates(.EquippedShieldObjIndex, defenderUserIndex, attackerIsNpc, attackerIndex, logCtx)
        Call FireSlotOnDamagedStates(.EquippedRingAccesoryObjIndex, defenderUserIndex, attackerIsNpc, attackerIndex, logCtx)
        Call FireSlotOnDamagedStates(.EquippedAmuletAccesoryObjIndex, defenderUserIndex, attackerIsNpc, attackerIndex, logCtx)
        Call FireSlotOnDamagedStates(.EquippedBackpackObjIndex, defenderUserIndex, attackerIsNpc, attackerIndex, logCtx)
        Call FireSlotOnDamagedStates(.EquippedSaddleObjIndex, defenderUserIndex, attackerIsNpc, attackerIndex, logCtx)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.FireUserOnDamagedStates", Erl)
End Sub

Private Sub FireSlotOnDamagedStates(ByVal ObjIndex As Integer, ByVal defenderUserIndex As Integer, ByVal attackerIsNpc As Boolean, ByVal attackerIndex As Integer, ByVal logCtx As String)
    If ObjIndex > 0 Then Call FireOnDamagedStates(ObjData(ObjIndex).Elemental, defenderUserIndex, eUser, attackerIsNpc, attackerIndex, logCtx, slotObjIndex:=ObjIndex)
End Sub

' ============================================================================
' Punto de entrada gateado: camino user -> target (NPC o USER). Generalizado en la
' Ola 5 para habilitar PvP. Devuelve el dano elemental EXTRA (ya resistido por tipo).
' ============================================================================
Public Function ElementalDamageUserVsTarget(ByVal UserIndex As Integer, ByVal targetIsNpc As Boolean, ByVal targetIndex As Integer, ByVal WeaponObjIndex As Integer, ByVal MunitionObjIndex As Integer, ByRef outColor As Long, Optional ByRef outType As Integer) As Long
    On Error GoTo ErrHandler
    outColor = vbWhite
    outType = eDmgNone
    If Not ElementalSystemEnabled() Then Exit Function
    If UserIndex <= 0 Or targetIndex <= 0 Then Exit Function
    Dim total As Long
    Dim ctx As String
    ctx = "U" & UserIndex & IIf(targetIsNpc, "->N", "->U") & targetIndex
    ' Componentes + procs onHit del arma
    If WeaponObjIndex > 0 Then
        total = total + ResolveComponentsVsTarget(ObjData(WeaponObjIndex).Elemental, targetIsNpc, targetIndex, UserIndex, eUser, WeaponObjIndex, 0, ctx & " weap")
        total = total + FireProcs(ObjData(WeaponObjIndex).Elemental, eProcOnHit, targetIsNpc, targetIndex, UserIndex, eUser, WeaponObjIndex, 0, ctx & " weap")
    End If
    ' Municion (rango): suma sus componentes/procs solo si el arma es de proyectil
    If MunitionObjIndex > 0 And WeaponObjIndex > 0 Then
        If ObjData(WeaponObjIndex).Proyectil > 0 Then
            total = total + ResolveComponentsVsTarget(ObjData(MunitionObjIndex).Elemental, targetIsNpc, targetIndex, UserIndex, eUser, MunitionObjIndex, 0, ctx & " ammo")
            total = total + FireProcs(ObjData(MunitionObjIndex).Elemental, eProcOnHit, targetIsNpc, targetIndex, UserIndex, eUser, MunitionObjIndex, 0, ctx & " ammo")
            ' Encantamiento temporal de las flechas (aceite/hechizo de flechas). Aditivo. (20.002 CP1 ammo)
            With UserList(UserIndex).flags
                If .EnchantedAmmoObjIndex > 0 And .EnchantedAmmoObjIndex = MunitionObjIndex Then
                    If .EnchantedAmmoPermanent = 1 Or Not DeadlinePassed(GetTickCountRaw(), .EnchantedAmmoDeadline) Then
                        total = total + ResolveComponentsVsTarget(.EnchantedAmmoSource, targetIsNpc, targetIndex, UserIndex, eUser, MunitionObjIndex, .EnchantedAmmoSrcItemObjIndex, ctx & " ammoench")
                        total = total + FireProcs(.EnchantedAmmoSource, eProcOnHit, targetIsNpc, targetIndex, UserIndex, eUser, MunitionObjIndex, .EnchantedAmmoSrcItemObjIndex, ctx & " ammoench")
                    End If
                End If
            End With
        End If
    End If
    ' Encantamiento temporal del arma (hechizo Encantar Arma). Aditivo al elemental base.
    With UserList(UserIndex).flags
        If .EnchantWeaponObjIndex > 0 And .EnchantWeaponObjIndex = WeaponObjIndex Then
            If .EnchantWeaponPermanent = 1 Or Not DeadlinePassed(GetTickCountRaw(), .EnchantWeaponDeadline) Then
                total = total + ResolveComponentsVsTarget(.EnchantWeaponSource, targetIsNpc, targetIndex, UserIndex, eUser, WeaponObjIndex, .EnchantWeaponSrcItemObjIndex, ctx & " ench")
                total = total + FireProcs(.EnchantWeaponSource, eProcOnHit, targetIsNpc, targetIndex, UserIndex, eUser, WeaponObjIndex, .EnchantWeaponSrcItemObjIndex, ctx & " ench")
            End If
        End If
    End With
    ' Orbe/amuleto equipado: aplica solo con arma MELEE base sin elemental propio (CP3 Step 7):
    ' R5 no afecta arcos/proyectiles; R6 no se suma si el arma ya tiene elemental propio (vive el del arma).
    Dim orbIdx As Integer
    orbIdx = UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex
    If orbIdx > 0 Then
        Dim orbAplica As Boolean
        orbAplica = True
        If WeaponObjIndex > 0 Then
            If ObjData(WeaponObjIndex).Proyectil > 0 Then orbAplica = False
            If ObjData(WeaponObjIndex).Elemental.CompCount > 0 Or ObjData(WeaponObjIndex).Elemental.ProcCount > 0 Then orbAplica = False
        End If
        If orbAplica Then
            total = total + ResolveComponentsVsTarget(ObjData(orbIdx).Elemental, targetIsNpc, targetIndex, UserIndex, eUser, orbIdx, 0, ctx & " orb")
            total = total + FireProcs(ObjData(orbIdx).Elemental, eProcOnHit, targetIsNpc, targetIndex, UserIndex, eUser, orbIdx, 0, ctx & " orb")
        End If
    End If
    ' 06.002 Ola 2: los procs onDamaged del defensor (dmgBonus=espinas, applyState=efectos) se
    ' disparan POST-dano con gate de supervivencia en los call-sites de SistemaCombate
    ' (ResolveThorns / FireOnDamagedStates). Aca ya no se dispara nada pre-dano.
    ' Color del numero elemental: tipo primario del arma base; si no tiene, del encantamiento.
    If WeaponObjIndex > 0 Then
        If ObjData(WeaponObjIndex).Elemental.CompCount > 0 Then
            outColor = DamageTypeColor(ObjData(WeaponObjIndex).Elemental.Comp(1).DamageType)
            outType = ObjData(WeaponObjIndex).Elemental.Comp(1).DamageType
        ElseIf UserList(UserIndex).flags.EnchantWeaponObjIndex = WeaponObjIndex And UserList(UserIndex).flags.EnchantWeaponSource.CompCount > 0 Then
            outColor = DamageTypeColor(UserList(UserIndex).flags.EnchantWeaponSource.Comp(1).DamageType)
            outType = UserList(UserIndex).flags.EnchantWeaponSource.Comp(1).DamageType
        End If
    End If
    If outColor = vbWhite And orbIdx > 0 Then
        If ObjData(orbIdx).Elemental.CompCount > 0 Then
            outColor = DamageTypeColor(ObjData(orbIdx).Elemental.Comp(1).DamageType)
            outType = ObjData(orbIdx).Elemental.Comp(1).DamageType
        End If
    End If
    If outColor = vbWhite And MunitionObjIndex > 0 Then
        With UserList(UserIndex).flags
            If .EnchantedAmmoObjIndex = MunitionObjIndex And .EnchantedAmmoSource.CompCount > 0 Then
                outColor = DamageTypeColor(.EnchantedAmmoSource.Comp(1).DamageType)
                outType = .EnchantedAmmoSource.Comp(1).DamageType
            ElseIf ObjData(MunitionObjIndex).Elemental.CompCount > 0 Then
                outColor = DamageTypeColor(ObjData(MunitionObjIndex).Elemental.Comp(1).DamageType)
                outType = ObjData(MunitionObjIndex).Elemental.Comp(1).DamageType
            End If
        End With
    End If
    ElementalDamageUserVsTarget = total
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.ElementalDamageUserVsTarget", Erl)
End Function

' Wrapper retrocompatible: camino user -> NPC (call-site existente sin cambios).
Public Function ElementalDamageUserVsNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal WeaponObjIndex As Integer, ByVal MunitionObjIndex As Integer, ByRef outColor As Long, Optional ByRef outType As Integer) As Long
    ElementalDamageUserVsNpc = ElementalDamageUserVsTarget(UserIndex, True, NpcIndex, WeaponObjIndex, MunitionObjIndex, outColor, outType)
End Function

Public Sub OnEnchantedWeaponSwing(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    If Not ElementalSystemEnabled() Then Exit Sub
    With UserList(UserIndex).flags
        If .EnchantWeaponObjIndex <= 0 Then Exit Sub
        ' El arma equipada debe coincidir con la encantada
        If UserList(UserIndex).invent.EquippedWeaponObjIndex <> .EnchantWeaponObjIndex Then Exit Sub
        ' Permanente: no expira ni consume cargas
        If .EnchantWeaponPermanent = 1 Then Exit Sub
        ' Expiracion por tiempo (igual que el veneno: se chequea al swingear)
        If .EnchantWeaponDeadline <> 0 Then
            If DeadlinePassed(GetTickCountRaw(), .EnchantWeaponDeadline) Then
                Call ClearEnchantedWeapon(UserIndex, "El encantamiento de tu arma se ha disipado.")
                Exit Sub
            End If
        End If
        ' Consumo de cargas (0 = sin limite de cargas, solo tiempo/permanente)
        If .EnchantWeaponCargas > 0 Then
            .EnchantWeaponCargas = .EnchantWeaponCargas - 1
            ' plan 10.001 Ola 5: charge_spent, unico evento nuevo en el hot path (un swing).
            If ElementalPlayerTelemetryEnabled() Then Call LogElementalBalance("charge_spent", ElementalBalanceActorId(False, UserIndex), ElementalBalanceActorClass(False, UserIndex), "0", "none", .EnchantWeaponObjIndex, ElementalBalanceCatalogTier(.EnchantWeaponObjIndex), 0, .EnchantWeaponCargas, 0, ElementalBalanceMap(False, UserIndex), ElementalBalanceCharId(False, UserIndex), ElementalBalanceAccountId(False, UserIndex), 0, .EnchantWeaponSrcItemObjIndex, 0)
            If .EnchantWeaponCargas <= 0 Then
                Call ClearEnchantedWeapon(UserIndex, "El encantamiento de tu arma se ha agotado.")
            Else
                ' 06.002 Ola 4: contador de cargas vivo en el HUD (espejo de OnEnchantedAmmoSwing).
                Call WriteUpdatePoisonStacks(UserIndex, .EnchantWeaponObjIndex, .EnchantWeaponCargas)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.OnEnchantedWeaponSwing", Erl)
End Sub

' Tick desde el game loop: expira el encantamiento por tiempo aunque el user no pegue (paridad con el veneno).
Public Sub CheckEnchantedWeaponTick(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    If Not ElementalSystemEnabled() Then Exit Sub
    With UserList(UserIndex).flags
        If .EnchantWeaponObjIndex <= 0 Then Exit Sub
        If .EnchantWeaponPermanent = 1 Then Exit Sub
        If .EnchantWeaponDeadline = 0 Then Exit Sub
        If DeadlinePassed(GetTickCountRaw(), .EnchantWeaponDeadline) Then
            Call ClearEnchantedWeapon(UserIndex, "El encantamiento de tu arma se ha disipado.")
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.CheckEnchantedWeaponTick", Erl)
End Sub

' Limpia el cache del encantamiento elemental del arma (cargas agotadas / tiempo expirado).
Public Sub ClearEnchantedWeapon(ByVal UserIndex As Integer, Optional ByVal msg As String = "")
    On Error GoTo ErrHandler
    If UserIndex <= 0 Then Exit Sub
    Dim oldObj As Integer
    With UserList(UserIndex).flags
        If .EnchantWeaponObjIndex = 0 And .EnchantWeaponCargas = 0 And .EnchantWeaponPermanent = 0 Then Exit Sub
        oldObj = .EnchantWeaponObjIndex
        Dim emptySrc As t_ElementalSource
        .EnchantWeaponObjIndex = 0
        .EnchantWeaponDeadline = 0
        .EnchantWeaponPermanent = 0
        .EnchantWeaponCargas = 0
        .EnchantWeaponSource = emptySrc
    End With
    ' 06.002 Ola 4: apaga el icono de estado (espejo de ClearEnchantedAmmo).
    If oldObj > 0 Then Call WriteSendSkillCdUpdate(UserIndex, CLIENT_EFFECT_ENCHANTED_WEAPON, CLng(oldObj), 0, 0, eDebuff, 0)
    If LenB(msg) > 0 Then Call WriteConsoleMsg(UserIndex, msg, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.ClearEnchantedWeapon", Erl)
End Sub

' True si el arma indicada tiene un encantamiento elemental ACTIVO (permanente, con cargas, o sin expirar).
Public Function IsWeaponEnchantedActive(ByVal UserIndex As Integer, ByVal WeaponObjIndex As Integer) As Boolean
    On Error GoTo ErrHandler
    If UserIndex <= 0 Or WeaponObjIndex <= 0 Then Exit Function
    With UserList(UserIndex).flags
        If .EnchantWeaponObjIndex <> WeaponObjIndex Then Exit Function
        If .EnchantWeaponPermanent = 1 Then
            IsWeaponEnchantedActive = True
        ElseIf .EnchantWeaponCargas > 0 Then
            IsWeaponEnchantedActive = True
        ElseIf .EnchantWeaponDeadline <> 0 Then
            IsWeaponEnchantedActive = Not DeadlinePassed(GetTickCountRaw(), .EnchantWeaponDeadline)
        End If
    End With
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.IsWeaponEnchantedActive", Erl)
End Function

' Tipo primario de una fuente elemental (componente primario, o proc primario si no hay componente).
Private Function PrimaryElementalType(ByRef src As t_ElementalSource) As Integer
    If src.CompCount > 0 Then
        PrimaryElementalType = src.Comp(1).DamageType
    ElseIf src.ProcCount > 0 Then
        PrimaryElementalType = src.Proc(1).DamageType
    Else
        PrimaryElementalType = eDmgNone
    End If
End Function

' CP2 (20.002 Step 7): valida que un arma sea encantable (Subtipo=11) y acepte el tipo de la fuente
' (TiposElementalCompatibles, CSV; vacio = todos). Devuelve False + outMsg de rechazo.
Public Function CanEnchantWeapon(ByVal UserIndex As Integer, ByVal WeaponObjIndex As Integer, ByRef src As t_ElementalSource, ByRef outMsg As String) As Boolean
    On Error GoTo ErrHandler
    outMsg = vbNullString
    If WeaponObjIndex <= 0 Then
        outMsg = "No hay un arma equipada para encantar."
        Exit Function
    End If
    If ObjData(WeaponObjIndex).Subtipo <> 11 Then
        outMsg = "Esa arma no se puede encantar."
        Exit Function
    End If
    ' Plan 03.002 R2: rama inalcanzable retirada (ningun item tiene Subtipo=11 y Proyectil=1 a la vez; el gate de Subtipo ya corta antes)
    ' CP3 (20.002 Step 7): no se puede encantar con un orbe elemental equipado (exclusividad)
    If HasElementalOrbEquipped(UserIndex) Then
        outMsg = "No podes encantar el arma con un orbe equipado."
        Exit Function
    End If
    ' Exclusividad (CP2): no encantar un arma ya untada con veneno (untado activo).
    If UserList(UserIndex).flags.PoisonedWeaponObjIndex = WeaponObjIndex And UserList(UserIndex).flags.PoisonedWeaponCargas > 0 Then
        outMsg = "Tu arma ya esta untada con veneno."
        Exit Function
    End If
    Dim csv As String
    csv = ObjData(WeaponObjIndex).TiposElementalCompatibles
    If LenB(csv) > 0 Then
        If InStr("," & csv & ",", "," & CStr(PrimaryElementalType(src)) & ",") = 0 Then
            outMsg = "El arma no acepta este tipo de encantamiento."
            Exit Function
        End If
    End If
    CanEnchantWeapon = True
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.CanEnchantWeapon", Erl)
End Function

' CP3 (20.002 Step 7): setter unificado del encantamiento de arma (aceite/hechizo). Un solo punto de verdad.
Public Sub SetEnchantedWeapon(ByVal UserIndex As Integer, ByVal WeaponObjIndex As Integer, ByRef src As t_ElementalSource, ByVal cargas As Integer, ByVal durationMs As Long, Optional ByVal srcItemObjIndex As Integer = 0)
    Dim oldObj As Integer
    oldObj = UserList(UserIndex).flags.EnchantWeaponObjIndex
    With UserList(UserIndex).flags
        .EnchantWeaponObjIndex = WeaponObjIndex
        .EnchantWeaponSource = src
        .EnchantWeaponCargas = cargas
        .EnchantWeaponSrcItemObjIndex = srcItemObjIndex ' plan 10.001 Ola 5 (Hecho 43): ObjIndex del aceite/orbe fuente
        If durationMs < 0 Then
            .EnchantWeaponPermanent = 1
            .EnchantWeaponDeadline = 0
        Else
            .EnchantWeaponPermanent = 0
            .EnchantWeaponDeadline = AddMod32(GetTickCountRaw(), durationMs)
        End If
    End With
    ' 06.002 Ola 4: icono de estado del arma encantada en el HUD (asimetria con flechas resuelta).
    ' Si habia un encantamiento previo sobre OTRA arma, apagar su icono para no dejarlo inmortal.
    If oldObj > 0 And oldObj <> WeaponObjIndex Then Call WriteSendSkillCdUpdate(UserIndex, CLIENT_EFFECT_ENCHANTED_WEAPON, CLng(oldObj), 0, 0, eDebuff, 0)
    Dim iconDur As Long
    If durationMs < 0 Then iconDur = -1 Else iconDur = durationMs
    Call WriteSendSkillCdUpdate(UserIndex, CLIENT_EFFECT_ENCHANTED_WEAPON, CLng(WeaponObjIndex), iconDur, iconDur, eDebuff, cargas)
    ' plan 10.001 Ola 5: enchant_weapon, evento raro (alimenta B5).
    If ElementalPlayerTelemetryEnabled() Then Call LogElementalBalance("enchant_weapon", ElementalBalanceActorId(False, UserIndex), ElementalBalanceActorClass(False, UserIndex), "0", "none", WeaponObjIndex, ElementalBalanceCatalogTier(WeaponObjIndex), 0, cargas, durationMs, ElementalBalanceMap(False, UserIndex), ElementalBalanceCharId(False, UserIndex), ElementalBalanceAccountId(False, UserIndex), 0, srcItemObjIndex, 0)
End Sub

' ============================================================================
' Encantar Flechas elemental (20.002 CP1 ammo) - espejo de EnchantWeapon para la municion
' ============================================================================
Public Function CanEnchantAmmo(ByVal UserIndex As Integer, ByVal AmmoObjIndex As Integer, ByRef src As t_ElementalSource, ByRef outMsg As String) As Boolean
    On Error GoTo ErrHandler
    outMsg = vbNullString
    If AmmoObjIndex <= 0 Then
        outMsg = "Equipa flechas para encantarlas."
        Exit Function
    End If
    If ObjData(AmmoObjIndex).OBJType <> e_OBJType.otArrows Then
        outMsg = "Solo se pueden encantar flechas."
        Exit Function
    End If
    ' Exclusividad (CP2): no encantar flechas con veneno propio (nativo).
    If ObjData(AmmoObjIndex).FlechaVenenoFamilia > 0 Then
        outMsg = "Esas flechas ya tienen veneno propio."
        Exit Function
    End If
    ' Exclusividad (CP2): no encantar flechas ya untadas con veneno (untado activo).
    If UserList(UserIndex).flags.PoisonedAmmoObjIndex = AmmoObjIndex And UserList(UserIndex).flags.PoisonedAmmoCargas > 0 Then
        outMsg = "Esas flechas ya estan untadas con veneno."
        Exit Function
    End If
    Dim csv As String
    csv = ObjData(AmmoObjIndex).TiposElementalCompatibles
    If LenB(csv) > 0 Then
        If InStr("," & csv & ",", "," & CStr(PrimaryElementalType(src)) & ",") = 0 Then
            outMsg = "Esas flechas no aceptan este tipo de encantamiento."
            Exit Function
        End If
    End If
    CanEnchantAmmo = True
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.CanEnchantAmmo", Erl)
End Function

' Setter del encantamiento de flechas (aceite/hechizo). durationMs<0 = permanente.
Public Sub SetEnchantedAmmo(ByVal UserIndex As Integer, ByVal AmmoObjIndex As Integer, ByRef src As t_ElementalSource, ByVal cargas As Integer, ByVal durationMs As Long, Optional ByVal srcItemObjIndex As Integer = 0)
    ' 06.002 Ola 4: si habia un encantamiento previo sobre OTRA municion, apagar su icono
    ' (quedaba un "Flechas Encantadas" inmortal al cambiar de stack).
    Dim oldObj As Integer
    oldObj = UserList(UserIndex).flags.EnchantedAmmoObjIndex
    If oldObj > 0 And oldObj <> AmmoObjIndex Then Call WriteSendSkillCdUpdate(UserIndex, CLIENT_EFFECT_ENCHANTED_AMMO, CLng(oldObj), 0, 0, eDebuff, 0)
    With UserList(UserIndex).flags
        .EnchantedAmmoObjIndex = AmmoObjIndex
        .EnchantedAmmoSource = src
        .EnchantedAmmoCargas = cargas
        .EnchantedAmmoSrcItemObjIndex = srcItemObjIndex ' plan 10.001 Ola 5 (Hecho 43): ObjIndex del aceite/orbe fuente
        If durationMs < 0 Then
            .EnchantedAmmoPermanent = 1
            .EnchantedAmmoDeadline = 0
        Else
            .EnchantedAmmoPermanent = 0
            .EnchantedAmmoDeadline = AddMod32(GetTickCountRaw(), durationMs)
        End If
    End With
    ' Icono de estado en el HUD del tirador (espejo del veneno WritePoisonedAmmoIcon). Effect57 = Flechas Encantadas.
    Dim iconDur As Long
    If durationMs < 0 Then iconDur = -1 Else iconDur = durationMs
    Call WriteSendSkillCdUpdate(UserIndex, CLIENT_EFFECT_ENCHANTED_AMMO, CLng(AmmoObjIndex), iconDur, iconDur, eDebuff, cargas)
    ' plan 10.001 Ola 5: enchant_ammo, evento raro (alimenta B5).
    If ElementalPlayerTelemetryEnabled() Then Call LogElementalBalance("enchant_ammo", ElementalBalanceActorId(False, UserIndex), ElementalBalanceActorClass(False, UserIndex), "0", "none", AmmoObjIndex, ElementalBalanceCatalogTier(AmmoObjIndex), 0, cargas, durationMs, ElementalBalanceMap(False, UserIndex), ElementalBalanceCharId(False, UserIndex), ElementalBalanceAccountId(False, UserIndex), 0, srcItemObjIndex, 0)
End Sub

' Consumo al disparar: 1 carga por flecha disparada (espejo de OnEnchantedWeaponSwing).
Public Sub OnEnchantedAmmoSwing(ByVal UserIndex As Integer, ByVal AmmoObjIndex As Integer)
    On Error GoTo ErrHandler
    If Not ElementalSystemEnabled() Then Exit Sub
    With UserList(UserIndex).flags
        If .EnchantedAmmoObjIndex <= 0 Then Exit Sub
        If .EnchantedAmmoObjIndex <> AmmoObjIndex Then Exit Sub
        If .EnchantedAmmoPermanent = 1 Then Exit Sub
        If .EnchantedAmmoDeadline <> 0 Then
            If DeadlinePassed(GetTickCountRaw(), .EnchantedAmmoDeadline) Then
                Call ClearEnchantedAmmo(UserIndex, "El encantamiento de tus flechas se ha disipado.")
                Exit Sub
            End If
        End If
        If .EnchantedAmmoCargas > 0 Then
            .EnchantedAmmoCargas = .EnchantedAmmoCargas - 1
            ' plan 10.001 Ola 5: charge_spent, unico evento nuevo en el hot path (un disparo).
            If ElementalPlayerTelemetryEnabled() Then Call LogElementalBalance("charge_spent", ElementalBalanceActorId(False, UserIndex), ElementalBalanceActorClass(False, UserIndex), "0", "none", .EnchantedAmmoObjIndex, ElementalBalanceCatalogTier(.EnchantedAmmoObjIndex), 0, .EnchantedAmmoCargas, 0, ElementalBalanceMap(False, UserIndex), ElementalBalanceCharId(False, UserIndex), ElementalBalanceAccountId(False, UserIndex), 0, .EnchantedAmmoSrcItemObjIndex, 0)
            If .EnchantedAmmoCargas <= 0 Then
                Call ClearEnchantedAmmo(UserIndex, "El encantamiento de tus flechas se ha agotado.")
            Else
                Call WriteUpdatePoisonStacks(UserIndex, .EnchantedAmmoObjIndex, .EnchantedAmmoCargas)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.OnEnchantedAmmoSwing", Erl)
End Sub

' Tick del game loop: expira por tiempo aunque el user no dispare.
Public Sub CheckEnchantedAmmoTick(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    If Not ElementalSystemEnabled() Then Exit Sub
    With UserList(UserIndex).flags
        If .EnchantedAmmoObjIndex <= 0 Then Exit Sub
        If .EnchantedAmmoPermanent = 1 Then Exit Sub
        If .EnchantedAmmoDeadline = 0 Then Exit Sub
        If DeadlinePassed(GetTickCountRaw(), .EnchantedAmmoDeadline) Then
            Call ClearEnchantedAmmo(UserIndex, "El encantamiento de tus flechas se ha disipado.")
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.CheckEnchantedAmmoTick", Erl)
End Sub

Public Sub ClearEnchantedAmmo(ByVal UserIndex As Integer, Optional ByVal msg As String = "")
    On Error GoTo ErrHandler
    If UserIndex <= 0 Then Exit Sub
    Dim oldObj As Integer
    With UserList(UserIndex).flags
        If .EnchantedAmmoObjIndex = 0 And .EnchantedAmmoCargas = 0 And .EnchantedAmmoPermanent = 0 Then Exit Sub
        oldObj = .EnchantedAmmoObjIndex
        Dim emptySrc As t_ElementalSource
        .EnchantedAmmoObjIndex = 0
        .EnchantedAmmoDeadline = 0
        .EnchantedAmmoPermanent = 0
        .EnchantedAmmoCargas = 0
        .EnchantedAmmoSource = emptySrc
    End With
    ' Apaga el icono de estado del tirador (Effect57).
    If oldObj > 0 Then Call WriteSendSkillCdUpdate(UserIndex, CLIENT_EFFECT_ENCHANTED_AMMO, CLng(oldObj), 0, 0, eDebuff, 0)
    If LenB(msg) > 0 Then Call WriteConsoleMsg(UserIndex, msg, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.ClearEnchantedAmmo", Erl)
End Sub

Public Function IsAmmoEnchantedActive(ByVal UserIndex As Integer, ByVal AmmoObjIndex As Integer) As Boolean
    On Error GoTo ErrHandler
    If UserIndex <= 0 Or AmmoObjIndex <= 0 Then Exit Function
    With UserList(UserIndex).flags
        If .EnchantedAmmoObjIndex <> AmmoObjIndex Then Exit Function
        If .EnchantedAmmoPermanent = 1 Then
            IsAmmoEnchantedActive = True
        ElseIf .EnchantedAmmoCargas > 0 Then
            IsAmmoEnchantedActive = True
        ElseIf .EnchantedAmmoDeadline <> 0 Then
            IsAmmoEnchantedActive = Not DeadlinePassed(GetTickCountRaw(), .EnchantedAmmoDeadline)
        End If
    End With
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.IsAmmoEnchantedActive", Erl)
End Function

' True si el user tiene un orbe elemental equipado (amuleto cuyo Elemental tiene componentes o procs).
Public Function HasElementalOrbEquipped(ByVal UserIndex As Integer) As Boolean
    Dim orbIdx As Integer
    orbIdx = UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex
    If orbIdx <= 0 Then Exit Function
    HasElementalOrbEquipped = (ObjData(orbIdx).Elemental.CompCount > 0 Or ObjData(orbIdx).Elemental.ProcCount > 0)
End Function

' Punto de entrada: camino NPC -> user (PvP elemental). El NPC atacante saca sus componentes/procs
' de NpcInfoCache. Devuelve el dano elemental EXTRA (ya resistido por el tipo del defensor user).
Public Function ElementalDamageNpcVsUser(ByVal NpcIndex As Integer, ByVal VictimaIndex As Integer, ByRef outColor As Long, Optional ByRef outType As Integer) As Long
    On Error GoTo ErrHandler
    outColor = vbWhite
    outType = eDmgNone
    If Not ElementalSystemEnabled() Then Exit Function
    If NpcIndex <= 0 Or VictimaIndex <= 0 Then Exit Function
    Dim t As Integer
    t = NpcList(NpcIndex).Numero
    If t <= 0 Then Exit Function
    Dim total As Long
    Dim ctx As String
    ctx = "N" & NpcIndex & "->U" & VictimaIndex
    ' Componentes + procs onHit del NPC atacante (target = user).
    total = total + ResolveComponentsVsTarget(NpcInfoCache(t).Elemental, False, VictimaIndex, NpcIndex, eNpc, 0, 0, ctx & " npc")
    total = total + FireProcs(NpcInfoCache(t).Elemental, eProcOnHit, False, VictimaIndex, NpcIndex, eNpc, 0, 0, ctx & " npc")
    ' Color y tipo del numero: componente primario del NPC (plan 04.002 M4: expone outType,
    ' mismo patron y misma precedencia que ElementalDamageUserVsTarget).
    If NpcInfoCache(t).Elemental.CompCount > 0 Then
        outColor = DamageTypeColor(NpcInfoCache(t).Elemental.Comp(1).DamageType)
        outType = NpcInfoCache(t).Elemental.Comp(1).DamageType
    End If
    ElementalDamageNpcVsUser = total
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.ElementalDamageNpcVsUser", Erl)
End Function

' ============================================================================
' Tooling: log de efectos + inspeccion de resistencias
' ============================================================================
Public Sub ElementalLog(ByVal msg As String)
    On Error GoTo ErrHandler
    If Not ElementalSystemEnabled() Then Exit Sub
    Dim fname As String
    fname = App.Path & "\Logs\Elemental_" & Format$(Date, "yyyy-mm-dd") & ".log"
    Dim fnum As Integer
    fnum = FreeFile
    Open fname For Append As #fnum
    Print #fnum, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & msg
    Close #fnum
    Exit Sub
ErrHandler:
    On Error Resume Next
    Close #fnum
End Sub

' Devuelve un texto con la resistencia efectiva del target para cada tipo built-in.
Public Function DescribeEntityResist(ByVal targetIsNpc As Boolean, ByVal targetIndex As Integer) As String
    Dim s As String, t As Long
    Dim r As t_ElementalResist
    For t = eDmgPhysical To eDmgBleed
        r = GetTargetResist(targetIsNpc, targetIndex, t)
        If r.Immune <> 0 Or r.ReduceChancePct <> 0 Or r.ReduceFlat <> 0 Or r.ReducePct <> 0 Then
            s = s & DamageTypeName(t) & ": "
            If r.Immune <> 0 Then
                s = s & "INMUNE"
            Else
                s = s & "chance=" & r.ReduceChancePct & "% flat=" & r.ReduceFlat & " pct=" & Format$(r.ReducePct, "0.00") & " (cap " & Format$(ResistCapForType(t), "0.00") & ")"
            End If
            s = s & "  "
        End If
    Next t
    If LenB(s) = 0 Then s = "sin resistencias elementales"
    DescribeEntityResist = s
End Function

' ============================================================================
' Parsers de datos (esquema centralizado; lo llaman FileIO/MODULO_NPCs)
' ============================================================================
Public Sub ParseElementalSourceFromIni(ByRef Leer As clsIniManager, ByVal sect As String, ByRef src As t_ElementalSource)
    Dim n As Integer, i As Integer, p As String
    src.CompCount = 0
    src.ProcCount = 0
    n = val(Leer.GetValue(sect, "ElemCompCount"))
    If n > MAX_DAMAGE_COMPONENTS Then n = MAX_DAMAGE_COMPONENTS
    If n < 0 Then n = 0
    For i = 1 To n
        p = "ElemComp" & i
        src.Comp(i).DamageType = val(Leer.GetValue(sect, p & "Type"))
        src.Comp(i).MinDamage = val(Leer.GetValue(sect, p & "Min"))
        src.Comp(i).MaxDamage = val(Leer.GetValue(sect, p & "Max"))
    Next i
    src.CompCount = n
    n = val(Leer.GetValue(sect, "ElemProcCount"))
    If n > MAX_ELEMENTAL_PROCS Then n = MAX_ELEMENTAL_PROCS
    If n < 0 Then n = 0
    For i = 1 To n
        p = "ElemProc" & i
        src.Proc(i).ChancePct = val(Leer.GetValue(sect, p & "Chance"))
        src.Proc(i).Kind = val(Leer.GetValue(sect, p & "Kind"))
        src.Proc(i).Trigger = val(Leer.GetValue(sect, p & "Trigger"))
        src.Proc(i).DamageType = val(Leer.GetValue(sect, p & "Type"))
        src.Proc(i).MinDamage = val(Leer.GetValue(sect, p & "Min"))
        src.Proc(i).MaxDamage = val(Leer.GetValue(sect, p & "Max"))
        src.Proc(i).EotId = val(Leer.GetValue(sect, p & "Eot"))
        src.Proc(i).ReflectPct = val(Leer.GetValue(sect, p & "ReflectPct"))
        src.Proc(i).Physical = val(Leer.GetValue(sect, p & "Physical"))
        src.Proc(i).BypassResist = val(Leer.GetValue(sect, p & "BypassResist"))
        src.Proc(i).Lethal = val(Leer.GetValue(sect, p & "Lethal"))
        src.Proc(i).ReflectMirror = val(Leer.GetValue(sect, p & "ReflectMirror"))
    Next i
    src.ProcCount = n
End Sub

Public Sub ParseElementalResistFromIni(ByRef Leer As clsIniManager, ByVal sect As String, ByRef rs As t_ElementalResistSet)
    Dim n As Integer, i As Integer, p As String
    rs.Count = 0
    n = val(Leer.GetValue(sect, "ElemResistCount"))
    If n > MAX_ELEMENTAL_RESISTS Then n = MAX_ELEMENTAL_RESISTS
    If n < 0 Then n = 0
    For i = 1 To n
        p = "ElemResist" & i
        rs.Resist(i).DamageType = val(Leer.GetValue(sect, p & "Type"))
        rs.Resist(i).ReduceChancePct = val(Leer.GetValue(sect, p & "Chance"))
        rs.Resist(i).ReduceFlat = val(Leer.GetValue(sect, p & "Flat"))
        rs.Resist(i).ReducePct = val(Leer.GetValue(sect, p & "Pct"))
        rs.Resist(i).ReduceFlatMin = val(Leer.GetValue(sect, p & "FlatMin"))
        rs.Resist(i).ReduceFlatMax = val(Leer.GetValue(sect, p & "FlatMax"))
        rs.Resist(i).ReducePctMin = val(Leer.GetValue(sect, p & "PctMin"))
        rs.Resist(i).ReducePctMax = val(Leer.GetValue(sect, p & "PctMax"))
        rs.Resist(i).Immune = val(Leer.GetValue(sect, p & "Immune"))
        rs.Resist(i).ImmuneDamage = val(Leer.GetValue(sect, p & "ImmuneDmg"))
        rs.Resist(i).ImmuneEffect = val(Leer.GetValue(sect, p & "ImmuneEff"))
        ' 06.002 Ola 2: claves EffMag/EffDur retiradas del esquema (nunca implementadas).
        rs.Resist(i).ReduceEffectChancePct = val(Leer.GetValue(sect, p & "EffChance"))
    Next i
    rs.Count = n
End Sub

' ============================================================================
' Crit universal (Ola 2): fallback para clases sin firma propia (bandido/asesino).
' Resistencia a crit via clave DMG_TYPE_CRIT (reduce chance y dano; inmune = sin crit).
' ============================================================================
Public Function UniversalCritActive() As Boolean
    UniversalCritActive = ElementalSystemEnabled()
End Function

' Devuelve el dano EXTRA de crit (0 si no critea). baseDamage = dano base del golpe.
Public Function TryUniversalCrit(ByVal UserIndex As Integer, ByVal targetIsNpc As Boolean, ByVal targetIndex As Integer, ByVal baseDamage As Long, ByRef outColor As Long) As Long
    On Error GoTo eh
    outColor = DamageTypeColor(DMG_TYPE_CRIT)
    If Not ElementalSystemEnabled() Then Exit Function
    If baseDamage <= 0 Then Exit Function
    Dim cr As t_ElementalResist
    cr = GetTargetResist(targetIsNpc, targetIndex, DMG_TYPE_CRIT)
    ' M7 (plan 02.001): Immune Y ImmuneDamage anulan el critico universal (antes
    ' solo Immune) -- consistente con ApplyElementalResist, que trata ambos igual.
    If cr.Immune <> 0 Or cr.ImmuneDamage <> 0 Then Exit Function
    Dim chance As Single
    ' M7 (plan 02.001): la resist al critico es MULTIPLICATIVA, no sustractiva.
    ' Antes "chance = base - ReduceChancePct" hacia que CUALQUIER ReduceChancePct
    ' >= la base (5) anulara el critico universal por completo (medido: Chance=50
    ' en el .dat de test daba critico IMPOSIBLE, no "la mitad" como se esperaba).
    Dim pct As Single
    pct = cr.ReduceChancePct
    If pct > 100 Then pct = 100
    If pct < 0 Then pct = 0
    chance = mUniversalCritChance * (1 - pct / 100)
    If chance <= 0 Then Exit Function
    If chance > 100 Then chance = 100
    If RandomNumber(1, 100) <= chance Then
        Dim bonus As Long
        bonus = Int(CDbl(baseDamage) * mUniversalCritMult)
        Dim nul As Boolean
        ' 06.002 Ola 2: ReduceChancePct ya redujo la chance de critear arriba; dejarlo vivo
        ' hacia que ApplyElementalResist lo rolee de nuevo para anular el bono (resistencia doble).
        cr.ReduceChancePct = 0
        bonus = ApplyElementalResist(bonus, cr, DMG_TYPE_CRIT, nul)
        If bonus < 0 Then bonus = 0
        TryUniversalCrit = bonus
        Call ElementalLog("U" & UserIndex & " universal crit base=" & baseDamage & " bonus=" & bonus)
        If ElementalBalanceLogEnabled() Then Call LogElementalBalance("universal_crit", ElementalBalanceActorId(False, UserIndex), ElementalBalanceActorClass(False, UserIndex), ElementalBalanceActorId(targetIsNpc, targetIndex), IIf(targetIsNpc, "npc", "user"), 0, 0, DMG_TYPE_CRIT, 0, bonus, ElementalBalanceMap(targetIsNpc, targetIndex), ElementalBalanceCharId(False, UserIndex), ElementalBalanceAccountId(False, UserIndex), ElementalBalanceCharId(targetIsNpc, targetIndex), 0, ElementalBalanceFightId(ElementalBalanceCharId(False, UserIndex), ElementalBalanceCharId(targetIsNpc, targetIndex)))
    End If
    Exit Function
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.TryUniversalCrit", Erl)
End Function

' ============================================================================
' Cap de control (Ola 3): nunca congelar por slow. Devuelve el multiplicador de
' velocidad (1+mod) con el slow capeado. Toggle OFF = vanilla (floor 0 = max(0,...)).
' El freeze duro sigue siendo por flag Paralizado/Inmovilizado, aparte de esto.
' ============================================================================
Public Function CappedSpeedMult(ByVal moveSpeedMod As Single) As Single
    Dim mult As Single
    mult = 1 + moveSpeedMod
    Dim floorVal As Single
    If ElementalSlowCapActive Then
        floorVal = mElementalSlowFloor
    Else
        floorVal = 0
    End If
    If mult < floorVal Then mult = floorVal
    CappedSpeedMult = mult
End Function
