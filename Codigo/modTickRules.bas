Attribute VB_Name = "modTickRules"
' Argentum 20 Game Server
'
'    Copyright (C) 2025-2026 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
' Plan 21.003: reglas de tiempo como funciones PURAS con el tick como parametro. El
' codigo de juego las llama pasandole GetTickCountRaw(); la suite Unit_TickWrap las llama
' con ticks inyectados. Asi el test ejercita el mismo codigo que corre en juego y no una copia.
'
' ESTADO ACTUAL (fase 4): todas son wrap-safe. Usan aritmetica sin signo mod 2^32
' (modElapsedTime.TicksElapsed) en vez de la resta de Longs con signo, que desbordaba (error 6)
' al cruzar &H7FFFFFFF -> &H80000000 o el wrap &HFFFFFFFF -> 0, y de los centinelas "> 0",
' que fallan con ticks negativos (uptime de Windows entre 24,8 y 49,7 dias).
'
' REGLAS:
'  - TicksElapsed(start, now): el tick VIEJO va PRIMERO. Con los argumentos invertidos compila
'    y da ~4,29e9 sin avisar. En este modulo siempre es TicksElapsed(<guardado>, nowTick).
'  - El tick 0 esta RESERVADO como "nunca guardado" (TickStampIsSet). Quien guarde un tick que
'    luego se pregunte con TickStampIsSet/CooldownActive lo guarda con NonZeroTick() /
'    NonZeroTickFrom(): si el reloj da 0 exacto (1 ms cada 2^32) guarda 1 (error de 1 ms).
'    Asignar 0 para limpiar el campo es correcto y es lo que significa "nunca".
Option Explicit

' Periodo minimo entre actualizaciones del radar de minimapa (antes Private en Radar.bas).
Public Const NPC_RADAR_TICK_MS As Long = 300

' Tick actual con el 0 reservado: devuelve t, o 1 si t = 0. Para guardar un tick que despues se
' consulta con TickStampIsSet/CooldownActive (0 = "nunca"). Esta aca y no en modElapsedTime
' porque ese modulo es del oficial (plan 21.003, desvio de D6).
Public Function NonZeroTickFrom(ByVal t As Long) As Long
    If t = 0 Then
        NonZeroTickFrom = 1
    Else
        NonZeroTickFrom = t
    End If
End Function

' Lo mismo leyendo el reloj ahora (si el llamador ya tiene un nowTick local, usar NonZeroTickFrom).
Public Function NonZeroTick() As Long
    NonZeroTick = NonZeroTickFrom(GetTickCountRaw())
End Function

' Cooldown de pociones curativas (InvUsuario.UseInvItem, CuresPoison).
' True si hay cooldown activo: la pocion se uso hace menos de cooldownMs.
Public Function CooldownActive(ByVal nowTick As Long, ByVal lastTick As Long, ByVal cooldownMs As Long) As Boolean
    CooldownActive = False
    If cooldownMs > 0 And TickStampIsSet(lastTick) Then
        If TicksElapsed(lastTick, nowTick) < cooldownMs Then CooldownActive = True
    End If
End Function

' Centinela "este tick fue guardado alguna vez": 0 es "nunca" (ver NonZeroTick). Un tick
' negativo (uptime de Windows entre 24,8 y 49,7 dias) es un tick guardado valido.
Public Function TickStampIsSet(ByVal t As Long) As Boolean
    TickStampIsSet = (t <> 0)
End Function

' Vencimiento del untado de veneno (arma/flecha): True si paso mas de maxMs desde appliedTick.
' El llamador mantiene la guardia "maxMs > 0".
' EffectsOverTime.CheckPoisonedWeaponTick guardaba la resta en un Long (elapsed) y despues
' comparaba: es la misma comparacion, asi que usa esta misma funcion (no hace falta PoisonElapsedMs).
Public Function PoisonApplicationExpired(ByVal nowTick As Long, ByVal appliedTick As Long, ByVal maxMs As Long) As Boolean
    PoisonApplicationExpired = (TicksElapsed(appliedTick, nowTick) > maxMs)
End Function

' Radar de minimapa: True si ya paso el periodo desde la ultima actualizacion.
' lastTick = 0 (recien logueado) funciona sin centinela: el tiempo desde el tick 0 es el tick
' sin signo, que es >= 300 salvo en el primer tramo del arranque de Windows.
Public Function RadarDue(ByVal nowTick As Long, ByVal lastTick As Long) As Boolean
    RadarDue = (TicksElapsed(lastTick, nowTick) >= NPC_RADAR_TICK_MS)
End Function

' Duracion de la sesion (ms) para la telemetria session_end (TCP.ClearAndSaveUser).
' Acotada a 2147483647 (una sesion de mas de 24,8 dias seguidos no cabe en un Long).
Public Function SessionElapsedMs(ByVal nowTick As Long, ByVal startTick As Long) As Long
    Dim e As Double
    e = TicksElapsed(startTick, nowTick)
    If e > 2147483647# Then
        SessionElapsedMs = 2147483647
    Else
        SessionElapsedMs = CLng(e)
    End If
End Function

' Duracion de una pelea elemental (ms) para fight_end (modElementalBalanceLog.LogElementalFightEnd).
' Acotada a 2147483647 como SessionElapsedMs.
Public Function FightDurationMs(ByVal nowTick As Long, ByVal startTick As Long) As Long
    Dim e As Double
    e = TicksElapsed(startTick, nowTick)
    If e > 2147483647# Then
        FightDurationMs = 2147483647
    Else
        FightDurationMs = CLng(e)
    End If
End Function

' True si el tick a es mas viejo que el tick b (modElementalBalanceLog.FreeElementalFightSlot).
' Compara la EDAD contra nowTick (mas edad = mas viejo), no los valores absolutos, que se
' invierten al cruzar &H7FFFFFFF. Con la misma edad no es mas viejo.
Public Function TickIsOlder(ByVal nowTick As Long, ByVal a As Long, ByVal b As Long) As Boolean
    TickIsOlder = (TicksElapsed(a, nowTick) > TicksElapsed(b, nowTick))
End Function

' Texto de un tick en las columnas ts_ms de los CSV de telemetria: el valor SIN signo
' 0..4294967295 (plan 21.003, D3). Format$(...,"0") y no CStr: con "0" nunca sale notacion
' exponencial ni separador decimal/de miles, sea cual sea la configuracion regional del host;
' CStr de un Double entero tambien daria lo mismo hasta 15 digitos, pero Format$ no depende de eso.
Public Function TickToCsv(ByVal t As Long) As String
    TickToCsv = Format$(TicksElapsed(0, t), "0")
End Function
