@echo off
setlocal enabledelayedexpansion

REM ===========================================================================
REM  AO20 - build del servidor VB6
REM
REM  Deja EVIDENCIA auditable en build-stamp.txt: flags usados, resultado real y
REM  hash del binario antes/despues. Nacio del plan 30.001, despues de perder
REM  tiempo adivinando con que constantes habia salido un binario.
REM
REM  Cambios sobre la version anterior:
REM   1. ORDEN DE ARGUMENTOS. Antes: /make /d "FLAGS" /out log Server.VBP
REM      La sintaxis de VB6 es /make <proyecto> ... /d <constantes>: el proyecto
REM      va inmediatamente despues de /make. Con el orden viejo VB6 puede tomar
REM      "/d" como argumento de /make y descartar las constantes, cayendo al
REM      CondComp del .vbp (que tiene DEBUGGING=1 y UsarQueSocket sin definir).
REM   2. Los flags viven en UNA variable y quedan escritos en el stamp.
REM   3. Corta si server.exe esta corriendo (antes: "Permission denied" enterrado
REM      en un log acumulativo que igual decia "succeeded" mas abajo).
REM   4. Borra el log antes: el /out de VB6 APPENDEA. Un vb6build.log con 60
REM      "succeeded" no dice si TU build salio bien.
REM   5. Compara hash antes/despues: avisa si el binario NO se reescribio.
REM   6. Exit code 0/1 real.
REM
REM  OJO al editar: usar !VAR! y no %VAR% dentro de los bloques if ( ), porque
REM  la ruta del compilador tiene "(x86)" y el parentesis rompe el parseo.
REM ===========================================================================

REM --- unica fuente de verdad de los flags -----------------------------------
set "AO20_FLAGS=UsarQueSocket=1:ConUpTime=1:AntiExternos=0:Lac=1:DEBUGGING=0:PYMMO=0:UNLOCK_CPU=0:DIRECT_PLAY=0"

set "VB6=C:\Program Files (x86)\Microsoft Visual Studio\VB98\vb6.exe"
set "HERE=%~dp0"
set "PROJ=!HERE!Server.VBP"
set "EXE=!HERE!Server.exe"
set "LOG=!HERE!vb6build.log"
set "STAMP=!HERE!build-stamp.txt"

echo ============================================================
echo  AO20 - build del servidor
echo ============================================================
REM --- flags reales: el CondComp del .vbp (el /d de AO20_FLAGS lo ignora VB6) ---
set "CONDCOMP=(sin CondComp en el .vbp)"
for /f "tokens=* delims=" %%C in ('findstr /B /C:"CondComp" "!PROJ!"') do set "CONDCOMP=%%C"
echo  flags reales : !CONDCOMP!
echo  /d ignorado  : !AO20_FLAGS!
echo.

if not exist "!VB6!" goto :sin_compilador
if not exist "!PROJ!" goto :sin_proyecto

REM --- server.exe NO puede estar corriendo ----------------------------------
tasklist /FI "IMAGENAME eq server.exe" 2>nul | find /I "server.exe" >nul
if not errorlevel 1 goto :server_corriendo

REM --- hash previo ----------------------------------------------------------
set "HASH_ANTES=(no habia binario)"
if exist "!EXE!" call :hash "!EXE!" HASH_ANTES

REM --- log limpio -----------------------------------------------------------
if exist "!LOG!" del /q "!LOG!" >nul 2>&1

REM --- compilar (proyecto inmediatamente despues de /make) ------------------
echo [1/2] Compilando...
"!VB6!" /make "!PROJ!" /out "!LOG!" /d "!AO20_FLAGS!"
set "VB6_RC=!ERRORLEVEL!"

REM VB6 no da un exit code confiable: manda el log.
set "OK=0"
if not exist "!LOG!" goto :sin_log
findstr /C:"succeeded" "!LOG!" >nul 2>&1 && set "OK=1"
findstr /C:"failed" "!LOG!" >nul 2>&1 && set "OK=0"
:sin_log

REM --- hash posterior -------------------------------------------------------
set "HASH_DESPUES=(no se genero binario)"
if exist "!EXE!" call :hash "!EXE!" HASH_DESPUES

REM --- stamp auditable ------------------------------------------------------
echo [2/2] Escribiendo build-stamp.txt...
> "!STAMP!" echo # AO20 build stamp - lo genera build_tmp.bat, no editar a mano
>> "!STAMP!" echo fecha          = %DATE% %TIME%
>> "!STAMP!" echo flags          = !CONDCOMP!
>> "!STAMP!" echo flags_d_ignorados = !AO20_FLAGS!
>> "!STAMP!" echo proyecto       = !PROJ!
>> "!STAMP!" echo vb6_exit_code  = !VB6_RC!
>> "!STAMP!" echo build_ok       = !OK!
>> "!STAMP!" echo hash_antes     = !HASH_ANTES!
>> "!STAMP!" echo hash_despues   = !HASH_DESPUES!
if exist "!LOG!" >> "!STAMP!" echo # ---- vb6build.log ----
if exist "!LOG!" type "!LOG!" >> "!STAMP!"

echo.
echo ------------------------------------------------------------
if "!OK!"=="0" goto :fallo
if "!HASH_ANTES!"=="!HASH_DESPUES!" goto :sin_cambio
echo  BUILD OK - binario nuevo
echo  hash: !HASH_DESPUES!
echo  evidencia: !STAMP!
echo ------------------------------------------------------------
exit /b 0

:sin_cambio
echo  BUILD OK - pero el binario NO CAMBIO
echo  El hash es igual al de antes: o no habia nada que recompilar,
echo  o VB6 no pudo escribir el archivo.
echo  evidencia: !STAMP!
echo ------------------------------------------------------------
exit /b 0

:fallo
echo  BUILD FALLIDO
echo  detalle en: !LOG!
if exist "!LOG!" type "!LOG!"
echo ------------------------------------------------------------
exit /b 1

:sin_compilador
echo [ERROR] No se encontro el compilador VB6 en:
echo         !VB6!
exit /b 1

:sin_proyecto
echo [ERROR] No se encontro el proyecto: !PROJ!
exit /b 1

:server_corriendo
echo [ERROR] server.exe esta CORRIENDO. VB6 no va a poder sobreescribir el
echo         binario y el build falla con "Permission denied".
echo         Cerra el server y volve a intentar.
exit /b 1

REM --- subrutina: hash sha256 de %1 en la variable %2 -----------------------
REM  No usar certutil: en Windows en espanol imprime "CertUtil: -hashfile comando
REM  completado correctamente" y ese texto se cuela en el parseo, con lo cual los
REM  dos hashes salen iguales y el script reporta un falso "el binario NO CAMBIO".
:hash
set "_h="
for /f "usebackq delims=" %%H in (`powershell -NoProfile -Command "(Get-FileHash -LiteralPath '%~1' -Algorithm SHA256).Hash"`) do if not defined _h set "_h=%%H"
set "%~2=!_h!"
goto :eof
