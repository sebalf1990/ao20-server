Attribute VB_Name = "modElementalBalanceLog"
' Argentum 20 Game Server
'
' modElementalBalanceLog
' Log estructurado de balance del motor elemental (plan 08.003, Ola 3 punto 6).
' Calcado de modPoisonLogging.bas: header escrito una vez, filas separadas por
' ";", costo cero cuando AMBOS toggles estan apagados. NO reemplaza ni toca
' modElementalCombat.ElementalLog (texto libre, Logs\Elemental_YYYY-MM-DD.log):
' este modulo agrega, EN PARALELO, una segunda escritura estructurada.
'
' ============================================================================
' Esquema v2 (plan 10.001, Ola 5, tajada A) -- decisiones y por que
' ============================================================================
' El esquema v1 (informe 2026-09-08, ver historial de este archivo en git) no
' podia responder telemetria de jugadores (Hecho 22 del plan 10.001): "item" y
' "tier" iban siempre en 0, y "attacker"/"victim" son indices de slot que se
' reusan entre jugadores. v2 agrega:
'
'   - char_id / account_id (columnas nuevas, al final): id estable del
'     ATACANTE (UserList(i).Id / .AccountID, Declares.bas:3265,3267). 0 si el
'     atacante es NPC o el slot no es un user logueado.
'   - victim_char_id (columna nueva, al final): id estable de la VICTIMA
'     cuando es un user logueado; 0 si es NPC.
'   - src_item (columna nueva): ObjIndex del aceite/orbe que originó un
'     encantamiento activo (Hecho 43). Solo se conoce en el camino
'     ench/ammoench de ElementalDamageUserVsTarget y en los eventos
'     enchant_weapon/enchant_ammo/charge_spent; 0 en cualquier otro evento.
'   - fight_id (columna nueva): identificador de pelea PvP (decision 12).
'     0 si el evento no es user-vs-user (una pelea son SIEMPRE dos
'     personajes; NPC vs user no arma "pelea" para B4).
'   - schema (columna nueva, al final): version numerica de este esquema
'     (constante ELEMENTAL_BALANCE_LOG_SCHEMA_VERSION), para que un lector
'     que junte varios archivos sepa sin ambiguedad que fila es de que forma.
'   - item / tier: se documentaban SIEMPRE 0; ahora llevan el ObjIndex real y
'     el tier del catalogo (04.003/07.001) donde el call site lo tiene a
'     mano. Ver "Decision item/tier por call site" mas abajo.
'
' Decision item/tier por call site (documentada en la bitacora de la Ola 5,
' plan 10.001, con la razon completa; resumen aca para quien lea el codigo):
'   - ResolveComponentsVsTarget y FireProcs (las dos funciones genericas que
'     el Hecho 22 senala) GANARON parametros nuevos (attackerIndex,
'     attackerType, itemObjIndex, srcItemObjIndex): las 6 llamadas que hace
'     ElementalDamageUserVsTarget/ElementalDamageNpcVsUser YA conocen esos
'     datos en el momento de llamar (arma/municion/orbe/encantamiento
'     equipados, o 0 en el camino NPC). Se opto por PROPAGAR, no por dejar
'     item=0: es la unica forma de que exista B3 (Objetivo B punto 9).
'   - ApplyDotTickResist NO gano attacker/item: solo recibe datos del target
'     (Hecho 22 lo documenta como limitacion de la funcion, no de este
'     modulo); agrandar su firma es rediseno del motor DoT, fuera de esta
'     Ola. Si gana victim_char_id, que SI esta a mano.
'   - ApplyThornsDamage/ResolveThorns/FireSlotThorns SI ganaron un parametro
'     de ObjIndex (el slot de gear que disparo la espina): FireSlotThorns ya
'     lo tenia y no lo pasaba (Hecho documentado en el header v1); es una
'     propagacion de una sola funcion, no un fan-in como el de arriba.
'   - TryUniversalCrit: el atacante (UserIndex) ya era parametro; se agregan
'     char_id/account_id/victim_char_id/fight_id calculados con eso. item
'     sigue en 0 (el critico universal no esta atado a un item, ya
'     documentado en v1).
'
' Tier del catalogo (OBJ9084-9161, docs/catalogo-elemental-jugable.md): solo
' los rangos que ese documento etiqueta explicitamente "T1"/"T2" devuelven
' 1/2. El resto del catalogo (flechas, Aceites Superiores, orbes, amuletos,
' gear de espinas, veneno-aceites, antidotos, pergaminos) NO tiene una T
' explicita en el doc: se devuelve 0 (no disponible) en vez de adivinar.
' Ver ElementalBalanceCatalogTier(). ElementalBalanceInCatalog() es un chequeo
' de RANGO aparte (9084-9161 completo) para filtrar craft/equip por catalogo
' sin depender de si ese item tiene tier asignado.
'
' Gotcha de compatibilidad del parser (Hecho 26, decision tomada en la Ola 5):
' `parsear_log_balance()` (tools/testclient/calibrar_dos_fuentes.py) es
' header-driven pero ESTRICTO en ancho: una fila con mas/menos campos que el
' ultimo header leido se descarta EN SILENCIO. Si v2 escribiera en el MISMO
' archivo que ya tiene el header v1 (escrito antes de este cambio), todas las
' filas v2 se perderian sin aviso porque LogElementalBalance solo escribe
' header cuando el archivo no existe (`Dir(fname)` vacio). Se eligio
' ARCHIVO POR VERSION DE ESQUEMA en vez de forzar un nuevo header en el
' archivo v1 a mitad de dia: separa fisicamente datos de forma distinta,
' hace trivial que Ola 6 sepa que esquema esta leyendo por el nombre del
' archivo, y no depende de que el escritor detecte "hoy ya tiene header
' viejo" (una condicion mas para romperse en silencio si alguien la toca sin
' leer este comentario). Archivo: Logs\ElementalBalance_YYYY-MM-DD_v2.log
' (el v1 sin sufijo, si algun dia se reactivara, seguiria en
' Logs\ElementalBalance_YYYY-MM-DD.log intacto). Probado en
' tools/testclient/tests/test_calibrar_dos_fuentes.py: un texto con header v1
' y otro con header v2 parsean cada uno con sus propias columnas.
'
' Retencion (decision 10, plan 10.001): esta Ola NO implementa rotacion
' automatica ("borrado cuando se pida, sin rotacion automatica"). Comando
' exacto para el operador cuando el dueno pida purgar (PowerShell, parado en
' dev/server; CONSERVAR una copia si todavia no se corrio
' analizar_telemetria.py sobre esos dias):
'   Remove-Item ".\Logs\ElementalBalance_*.log"
' La Ola 7 evalua si las tareas de backup ya existentes en la VM alcanzan
' para cubrir tambien esta carpeta antes de automatizar nada (decision 10,
' inciso a). Hasta entonces el archivo crece sin limite: riesgo aceptado por
' el dueno, no un olvido de esta Ola.
'
' Toggle de jugadores (plan 10.001, Objetivo B punto 4 y Ola 5 punto 4):
' TOGGLE33 "elemental_player_telemetry", independiente de "elemental_balance_log".
' Los eventos "raros" (enchant_weapon, enchant_ammo, craft, buy, equip,
' kill, death) y el unico evento de hot path nuevo (charge_spent) se gatean
' con ElementalPlayerTelemetryEnabled(), NO con ElementalBalanceLogEnabled():
' asi se puede prender la telemetria de jugadores en beta con el log de
' calibracion por golpe (el que pesa por swing, Hecho 24) apagado. Los 6
' eventos v1 (component, dot_tick_resist, proc_dmg_bonus, proc_apply_state,
' thorns, universal_crit) siguen gateados SOLO por elemental_balance_log,
' sin cambios de comportamiento.
'
' IMPORTANTE -- LogElementalBalance() DEJO de autogatear en
' "elemental_balance_log": con dos toggles independientes escribiendo al
' mismo Sub, ese gate unico hubiera apagado los eventos de jugador cada vez
' que el log de calibracion este OFF (justo el escenario que la decision 7
' pide habilitar). El Sub ahora sale temprano solo si LOS DOS toggles estan
' apagados (red de seguridad barata); cada call site sigue siendo responsable
' de chequear SU propio toggle antes de llamar, exactamente como ya hacian
' los 6 call sites v1 (`If ElementalBalanceLogEnabled() Then Call
' LogElementalBalance(...)`).
'
' Esquema de fila v3 (columna run_id agregada por el fix 2026-09-12, ver
' bloque de documentacion propio mas abajo; v1/v2 en el historial de git):
'   ts_ms;event;attacker;attacker_class;victim;victim_type;item;tier;
'   dmg_type;raw;final;resist_pct;pvp;map;char_id;account_id;
'   victim_char_id;src_item;fight_id;schema;run_id
'
' Los primeros 14 campos conservan el significado v1 (ver historial git para
' el docstring completo pre-v2); lo que cambia es que item/tier ahora pueden
' venir poblados. Campos nuevos (v2, salvo run_id que es v3):
'   char_id        = UserList(idx).Id del ATACANTE si es user logueado; 0 si
'                    es NPC o el slot no tiene un user logueado (Hecho 23).
'   account_id     = UserList(idx).AccountID del ATACANTE, mismo criterio.
'   victim_char_id = UserList(idx).Id de la VICTIMA si es user logueado; 0 si
'                    es NPC.
'   src_item       = ObjIndex del aceite/orbe fuente del encantamiento activo
'                    (Hecho 43); 0 si el evento no viene de un encantamiento.
'   fight_id       = ver ElementalBalanceFightId(). 0 si no aplica (no es
'                    user-vs-user, o el evento no participa de una pelea).
'   schema         = ELEMENTAL_BALANCE_LOG_SCHEMA_VERSION (3).
'   run_id         = ElementalBalanceRunId(). Unico por arranque del server;
'                    ver el bloque "Fix 2026-09-12" mas abajo del todo.
'
' Eventos nuevos de esta Ola (todos detras de elemental_player_telemetry):
'   enchant_weapon / enchant_ammo -> SetEnchantedWeapon/SetEnchantedAmmo
'     (modElementalCombat.bas). raw=cargas otorgadas, final=duracion en ms
'     (-1 = permanente). item=arma/municion encantada, src_item=aceite/orbe.
'   charge_spent -> OnEnchantedWeaponSwing/OnEnchantedAmmoSwing. UNICO evento
'     nuevo en el hot path (un swing/disparo con encantamiento activo por
'     cargas). raw=cargas restantes tras el descuento.
'   craft -> Trabajo.bas (4 oficios), solo si el item construido cae en el
'     catalogo (ElementalBalanceInCatalog). raw=cantidad construida.
'   buy -> Comercio.bas, compra ya consumada (Hecho 38). raw=cantidad,
'     final=precio total pagado. Sin filtro de catalogo (Hecho 38 no lo
'     pide; el filtro se aplica en el lector, Ola 6, y el evento es raro).
'   equip -> InvUsuario.EquiparInvItem, solo catalogo. Sin cantidad/precio.
'   kill / death -> Modulo_UsUaRiOs.DoDamageOrHeal, bloque de muerte
'     (:3112-3117 en la version citada por el plan), ANTES de que
'     CustomScenarios.UserDie/ActStats->UserDie borren el equipo (Hecho 42).
'     Una sola fila por muerte. item = arma equipada del ATACANTE (si es
'     user); src_item = arma equipada de la VICTIMA. Decision documentada en
'     la bitacora: NO se saca una foto completa de los 7 slots de gear de
'     ambos lados en esta tajada (hubiera exigido varias filas o columnas
'     nuevas por slot); item/src_item dan la pieza mas relevante (el arma)
'     de cada lado. raw=final=dano de la aplicacion letal (Abs(amount), ya
'     post-mitigacion). Cierra el fight_id si attacker y victim son ambos
'     user (decision 12: la pelea termina con la muerte de uno).
'
' Campos en 0 documentados por punto de insercion v1 (sin cambios de esta
' Ola salvo lo ya descripto arriba):
'   :dot_tick_resist ApplyDotTickResist -> attacker/attacker_class/item/
'         char_id/account_id/src_item/fight_id en 0: la funcion solo recibe
'         datos del target (targetIsNpc/targetIndex/dmgType/rawDamage), no
'         conoce el origen del DoT. victim_char_id SI se llena (el target es
'         conocido).
'   :component ResolveComponentsVsTarget -> desde esta Ola SI recibe
'         attacker/item/src_item (ver arriba); sigue en 0 solo cuando el
'         call site no los tiene (camino NPC->user).
'   :proc_dmg_bonus / :proc_apply_state FireProcs -> attacker/attacker_class
'         ya estaban disponibles (parametros de FireProcs); item/src_item
'         ahora tambien, mismo criterio que component. raw/final de
'         proc_apply_state siguen en 0 (no tiene un numero de dano propio).
'   :thorns ApplyThornsDamage -> item ahora SI se llena (ObjIndex del slot
'         que disparo la espina, propagado desde FireSlotThorns); raw sigue
'         en 0 (el dano pre-reduccion se pisa en ResolveThorns, no llega
'         aca). OJO roles invertidos: 'attackerIndex' es quien RECIBE el
'         reflejo (victim de este evento) y 'defenderIndex' quien lo origina
'         (attacker de este evento) -- son los nombres del golpe ORIGINAL
'         que genero las espinas.
'   :universal_crit TryUniversalCrit -> item sigue en 0 (el critico
'         universal no esta atado a un item puntual); raw en 0 (el bonus
'         pre-resistencia se pisa en la misma variable antes de este punto).
'
' ============================================================================
' Ola 5, tajada B (plan 10.001, punto 3b) -- eventos adicionales
' ============================================================================
' B0 -- ElementalBalanceUserReady(): la garantia de identidad de "equip" (y de
' todo evento nuevo de esta tajada que atribuye una accion a un jugador) deja
' de depender de que CADA call site pase bien el Optional UserIsLoggingIn.
' Antes, InvUsuario.bas:1675 solo chequeaba "Not UserIsLoggingIn"; TCP.bas:333
' (RellenarInventario, alta de personaje) y GameLogic.bas:1449 (resetPj, HOY
' sin caller vivo -- unico Call esta comentado en Protocol.bas:7939) no pasan
' ese flag y corren con UserLogged=False. El dia que un item de catalogo entre
' a un kit inicial o a un grant de GM, esos call sites escribirian una fila
' con char_id=0/account_id=0 sin que nadie lo note. ElementalBalanceUserReady
' ata la garantia al MISMO criterio que ya usa ElementalBalanceCharId
' (UserLogged=True), no al Optional. Los eventos "raros" de esta tajada
' (session_start/end, cast, gear_snapshot, npc_kill) lo usan tambien.
'
' Nuevos eventos (todos detras de elemental_player_telemetry salvo donde se
' indique, con salida temprana barata):
'   header -> primera fila de cada archivo del dia (mismo momento que el
'     header de texto). attacker=BuildStamp() (flags de compilacion, string
'     ya embebido en el .exe -- NO es un hash real: VB6 no trae SHA256 nativo
'     y calcularlo a mano es rediseno fuera de esta Ola). victim="n/a".
'     raw=elemental_balance_log (0/1). final=elemental_player_telemetry (0/1).
'     Para saber CON QUE obj.dat/binario se escribio cada fila, la regla
'     sigue siendo la que dejo la tajada A: cruzar ts_ms contra
'     build-stamp.txt (el server no puede autoreportar su propio hash).
'   death_cause -> mismo punto que "death" (Modulo_UsUaRiOs.bas, bloque de
'     DoDamageOrHeal). dmg_type = 0 pvp / 1 npc / 2 dot / 3 other. Verificado
'     en el codigo (no asumido): DamageSourceType=e_dot NO alcanza para
'     detectar "fue un tick de DoT" -- se reusa tambien para la rafaga
'     elemental embebida en un swing en vivo (SistemaCombate.bas:487,618,1403
'     pasan e_dot con SourceType=eUser/eNpc). La senal confiable es
'     SourceType=eNone (SourceIndex=0): ese patron SOLO lo usan
'     UpdateHpOverTime.cls/PoisonMinorEffect.cls/PoisonHemoEffect.cls cuando
'     el atacante original ya no es resoluble (tick sin fuente viva). Una
'     muerte por DoT con atacante todavia resoluble cae en pvp/npc, no en
'     dot: simplificacion documentada, no error.
'   drop_on_death -> InvUsuario.TirarTodosLosItems, en el DropObj real
'     (:3296), solo catalogo. item=ObjIndex tirado, raw=cantidad,
'     victim_char_id=el que murio.
'   cast -> modHechizos.LanzarHechizo, tras resolver SpellCastSuccess. item =
'     indice de Hechizos() (NO ObjIndex: este evento es la unica excepcion a
'     esa convencion, documentada aca porque no existe columna de spell id).
'     raw=mana gastado (diff antes/despues). final=delta de HP del target
'     (positivo=dano, negativo=cura), medido por diferencia de HP antes/
'     despues del Handle* -- no hay acceso directo al numero de dano en este
'     nivel de LanzarHechizo. Solo hechizos con target valido (user o npc);
'     los de terreno/mascota no generan esta fila (sin "victima" que loguear).
'   gear_snapshot -> al completar el login (Modulo_UsUaRiOs.ConnectUser_Complete,
'     antes de "ConnectUser_Complete = True"), una fila por cada slot de
'     catalogo equipado (arma/escudo/casco/armadura/anillo-accesorio/
'     municion/amuleto). DISTINTO de "equip": equip sigue excluyendo el login
'     (Not UserIsLoggingIn, ver S3 de la tajada A) para no contar la
'     reposicion de gear como una accion nueva; gear_snapshot es justamente
'     la foto que "equip" no puede dar sin volver a contarla como accion.
'   npc_kill -> MODULO_NPCs.MuereNpc, rama "lo mato un usuario" (UserIndex>0).
'     item = arma equipada del atacante. NO incluye "tiempo desde el primer
'     golpe": exigiria una tabla nueva de primer-golpe por NPC (reset en el
'     respawn, set en el primer damage), enganche en un lugar distinto
'     (NPCs.DoDamageOrHeal) y mas superficie de revision de la que esta
'     tajada puede cerrar con cuidado -- NO implementado, no es un olvido.
'   session_start / session_end -> Modulo_UsUaRiOs.ConnectUser_Complete /
'     TCP.CloseUser. raw=0 en session_start; en session_end, raw=milisegundos
'     jugados (modTickRules.SessionElapsedMs, wrap-safe; campo nuevo en
'     t_UserCounters, Declares.bas). Ambos exigen ElementalBalanceUserReady
'     (B0): un intento de cerrar sesion sobre un slot que nunca completo el
'     login no genera fila.
'
' Motor de agregados por pelea (fight_start / fight_end, plan 10.001, Ola 5
' punto 3b, prioridad 1): ARCHIVO SEPARADO, Logs\ElementalFights_<fecha>.log
' (ver ElementalFightLogFileName). Decision de diseno: el esquema v2 (20
' columnas fijas, ver arriba) no tiene lugar para nivel/clase/raza/faccion/HP
' de AMBOS lados ni para los 14 contadores de swings/hits/miss/blocks/
' pociones/casts que pide el punto 3b -- forzarlos en columnas que ya
' significan otra cosa (ej. "item"=raza) hubiera sido cambiar el significado
' de una columna sin cambiar su nombre, exactamente lo que este modulo evita
' en todos lados. Un archivo con su propio esquema (mismo patron que
' modElementalCombat.ElementalLog en paralelo a este modulo) no toca el
' ancho de las 20+ columnas ni los 26 call sites existentes de
' LogElementalBalance.
'
' Ciclo de vida de una "pelea" (par de char_id, decision 12 ya fijada en la
' tajada A): se CREA en el primer swing/cast hostil identificado entre dos
' chars (ElementalBalanceFightId, ahora con snapshot opcional de
' nivel/clase/raza/faccion/HP/mapa si el caller pasa los UserIndex --
' Optional, los 7 call sites viejos de componentes elementales no los pasan
' y esa fight_start queda con esos campos en blanco/0, degradacion aceptada
' porque en la practica casi toda pelea nace de un swing, no de un componente
' elemental suelto). Los contadores (swings/hits/miss/blocks via
' ElementalBalanceFightSwing desde SistemaCombate.bas; pociones via
' ElementalBalanceFightPotion desde InvUsuario.bas, SOLO si ya hay pelea
' activa, nunca crean una; casts via ElementalBalanceFightCast desde
' modHechizos.bas, solo hechizos hostiles) se acumulan EN MEMORIA (nada nuevo
' por swing en disco, tal como pide el punto 6) y se escriben una unica vez
' en fight_end. Se cierra (fight_end) en dos casos: "death" (el mismo
' closeFight:=True que la tajada A ya invocaba desde el bloque de muerte) y
' "timeout" (deteccion PEREZOSA: solo se nota si el MISMO par vuelve a
' cruzarse despues de los 30s de silencio -- una pelea donde uno de los dos
' se desconecta para siempre, o que nunca se repite, NO genera fight_end.
' Limitacion heredada del diseno de fight_id de la tajada A, no nueva de esta
' tajada; requeriria un barrido periodico activo, fuera de alcance aqui).

' ============================================================================
' Fix 2026-09-12 -- fight_id NO es unico entre reinicios (defecto real,
' cross-check del lector de la Ola 6 contra los logs en vivo del plan 10.001)
' ============================================================================
' Defecto: FightId sale de mNextFightId (Private, module-level), que arranca
' en 0 en CADA arranque del proceso. El archivo del dia es append-only y NO
' se abre uno nuevo por reinicio -- solo por fecha calendario (Format$(Date,
' "yyyy-mm-dd")) -- asi que dos arranques el MISMO dia escriben el MISMO
' fight_id para pares de personajes totalmente distintos. Medido en
' Logs\ElementalFights_2026-09-12.log: fight_id=1 aparece en 7 filas de DOS
' pares (56 vs 59, y 36 vs 66); fight_id=2 igual (56 vs 59 huerfano + 36 vs
' 66 cerrado). El lector (tools/analizar_telemetria.py, Ola 6) lo parcheaba
' agrupando y ordenando por ts_ms dentro de (dia, fight_id) -- funciona
' MIENTRAS el proceso vive (ts_ms = timeGetTime(), uptime del SISTEMA
' OPERATIVO, modElapsedTime.bas: sigue subiendo entre reinicios del PROCESO),
' pero es una inferencia, no una clave real, y se cae por completo si algun
' dia el SISTEMA OPERATIVO reinicia (ts_ms vuelve a un valor bajo): la Ola 7
' cosecha esto mismo con jugadores reales en la VM, donde un reinicio de
' maquina no es hipotetico.
'
' Decision: agregar un identificador de RUN, fijo una sola vez por arranque,
' como COLUMNA NUEVA al FINAL de cada fila (esquema aditivo, mismo criterio
' que toda columna nueva de este archivo) en vez de tocar el significado de
' fight_id. Candidatos considerados y por que se descartaron:
'   - GetTickCountRaw() en el momento del arranque (un "boot tick"): NO
'     sirve solo -- es el MISMO uptime del SO que ya es ambiguo entre
'     reinicios del PROCESO (dos arranques del proceso pueden capturarlo con
'     valores parecidos si son rapidos) y puede literalmente REPETIRSE tras
'     un reinicio de LA MAQUINA (vuelve a un valor bajo, igual que fight_id).
'   - GUID/random: exige una fuente de aleatoriedad que VB6 no trae nativa
'     sin declarar una API adicional; mas superficie que la que este fix
'     necesita.
'   - Format$(Now, "yyyymmddhhnnss") -- ELEGIDO: reloj de PARED (no de
'     uptime), unico entre reinicios de PROCESO y sobrevive igual a un
'     reinicio de LA MAQUINA (el reloj de pared no se resetea). Precision de
'     SEGUNDO alcanza de sobra: reiniciar_server.py tarda ~30s con
'     --after-build, muy por encima de la resolucion de este identificador
'     (dos arranques en el MISMO segundo son, en la practica, imposibles con
'     el flujo de reinicio de este proyecto). Se calcula LAZY (la primera
'     vez que un evento de telemetria lo necesita, no en Sub Main de
'     General.bas): evita tocar un archivo ajeno para este fix y sigue
'     siendo "una sola vez por arranque" porque el valor queda fijo en la
'     variable module-level (mRunId) el resto del proceso. Ver
'     ElementalBalanceRunId() mas abajo.
'
' Alcance de los dos archivos que llevan fight_id:
'   - ElementalFights_<fecha>.log (37 columnas, schema=1): pasa a
'     ElementalFights_<fecha>_v2.log (38 columnas, schema=2, +run_id al
'     final) -- mismo mecanismo de "archivo por version de esquema" que la
'     Ola 5 tajada A establecio para el log de balance (evita que un header
'     de 37 columnas ya escrito hoy trague en silencio filas de 38 columnas,
'     el mismo "Gotcha de compatibilidad del parser" documentado arriba).
'     El archivo viejo sin sufijo queda intacto (esquema v1 historico).
'   - ElementalBalance_<fecha>_v2.log (20 columnas, schema=2): TAMBIEN carga
'     fight_id (columna existente, la escriben los 7 call sites de
'     componentes elementales via ElementalBalanceFightId) -- la MISMA
'     ambiguedad entre reinicios aplica si algun consumidor futuro cruza
'     esta fila contra ElementalFights_*.log por fight_id (hoy el lector
'     Ola 6 no lo hace -- cruza por char_id+ts_ms -- pero dejar el campo
'     ambiguo en un archivo que se sigue escribiendo es una deuda latente,
'     no una limitacion documentada). Pasa a ElementalBalance_<fecha>_v3.log
'     (21 columnas, schema=3, +run_id al final), mismo mecanismo de archivo
'     por version que ya tenia.
' run_id es el MISMO valor (mRunId) en ambos archivos: un consumidor que
' cruce balance vs fights por (run_id, fight_id) en vez de solo fight_id ya
' tiene la clave real disponible desde este fix.

' ============================================================================
' Grupo 7 (plan 10.001, punto 7b) -- ultimo grupo de eventos, 2026-09-12
' ============================================================================
' Cierra los eventos que la Ola 5 tajada B dejo afuera por falta de tiempo
' ("si queda tiempo", bitacora de esa tajada). Los 8 eventos nuevos reusan las
' MISMAS 21 columnas del esquema v3 (nada de columna nueva, nada de version
' nueva de archivo): donde el evento no tiene un ObjIndex real que poner en
' "item", se documenta la excepcion aca mismo (mismo criterio que "cast" ya
' establecio para el indice de hechizo).
'
'   pickup / sell -> InvUsuario.PickObj / Comercio.Comercio (rama Venta).
'     Solo catalogo (ElementalBalanceInCatalog), igual criterio que craft/buy.
'     item=ObjIndex real, raw=cantidad, final=precio (0 en pickup).
'
'   craft_fail -> Trabajo.bas, un Else nuevo al lado de cada "craft" existente
'     (Herrero/Carpintero/Alquimista/Sastre). item=ObjIndex real (solo catalogo).
'     dmg_type REUSADO como motivo (mismo patron que death_cause): 0=materiales,
'     1=skill, 2=receta (tipo invalido para el oficio O no aprendida via
'     KnowsCraftingRecipe), 3=herramienta. Prioridad fija materiales>skill>
'     receta>herramienta cuando fallan varias a la vez. Herreria no tiene
'     concepto de herramienta equipable: nunca reporta motivo 3.
'
'   effect_end / antidote_use -> EffectsOverTime.bas + los 3 .cls de veneno
'     nuevo (PoisonMinorEffect/PoisonHemoEffect/PoisonNeuroEffect) +
'     InvUsuario.bas (Case CuresPoison). VERIFICADO EN EL CODIGO, no asumido
'     del plan: la reaplicacion de un veneno del mismo tipo NUNCA reemplaza la
'     instancia activa (CreatePoisonMinor/Hemo/Neuro llaman .Reset() sobre la
'     MISMA instancia via FindEffectOnTarget) -- el motivo "pisado por Override"
'     que pedia el punto 7b NO EXISTE para esta familia de efectos y se omite
'     (3 motivos, no 4): 0=expiro (natural, en el propio Update() de cada .cls),
'     1=curado (RemovePoisonMinor/Hemo/Neuro, los 12 call sites existentes son
'     TODOS curas de pocion/hechizo, verificado con grep), 2=muerte (barrido
'     nuevo LogPoisonEffectsEndOnDeath, llamado ANTES de ClearEffectList en
'     Modulo_UsUaRiOs.UserDie -- esa llamada generica borra los EOT de TODOS los
'     tipos sin distincion, por eso el barrido de muerte vive en un Sub aparte
'     que sabe leer los 3 tipos de veneno antes de que se pierdan).
'     item=EotId (indice de EffectsOverTime.dat, NO ObjIndex -- excepcion
'     documentada, mismo criterio que "cast"). raw=duracion servida en ms
'     (DurationTotalMs-DurationLeft, leido ANTES de remover). final=stacks
'     maximos alcanzados (solo Hemo stackea; Menor/Neuro siempre 0). Alcance
'     deliberadamente acotado a targets USER (telemetria de JUGADOR): un NPC
'     envenenado no genera fila.
'     antidote_use: Case e_PotionType.CuresPoison (TipoPocion=25, confirmado en
'     Declares.bas), solo cuando el item se consume de verdad (algoCurado=True).
'     item=ObjIndex real del antidoto (SI sigue la convencion normal, no es
'     excepcion). final=bitmask de que se curo (1=Menor, 2=Hemo, 4=Neuro).
'
'   recipe_learned / quest_complete -> modHechizos.AgregarHechizo (lectura de
'     un pergamino suelto) y ModQuest.FinishQuest (la quest entrega el hechizo
'     de receta directo, sin pasar por AgregarHechizo). Verificado en
'     Hechizos.dat: HECHIZO414-419 son "Receta: ..." (Forja Elemental Maestra,
'     Joyeria Elemental, Aceites Elementales Superiores, Toxinas y Antidotos
'     Superiores, Espinas, Vestimenta Elemental) -- confirma el rango que el
'     plan daba por sentado. item=indice de Hechizos() (excepcion, igual que
'     "cast") en recipe_learned; item=QuestIndex en quest_complete, acotado a
'     las quests cuya recompensa incluye un hechizo de receta (no todas las
'     quests, siguiendo el punto 7b al pie de la letra).
'
'   level_up / skill_assign -> Modulo_UsUaRiOs.CheckUserLevel y
'     Protocol.HandleModifySkills. El esquema de 21 columnas no tiene lugar
'     para un vector de ~20 skills sin resignificar columnas ya usadas: se
'     adapto el pedido del plan ("level_up con foto de los skills asignados")
'     separando la foto en su propio evento (skill_assign, una fila por skill
'     que de verdad cambio en la sesion de reparto) en vez de forzarla dentro
'     de level_up -- se cruzan por char_id+ts_ms si hace falta reconstruir la
'     distribucion completa. level_up: raw=puntos de skill otorgados en esta
'     pasada (Pts, puede cubrir mas de un nivel de una sola vez), final=nivel
'     alcanzado (.Stats.ELV). skill_assign: item=id de skill (e_Skill, otra
'     excepcion documentada a la convencion ObjIndex), raw=puntos asignados a
'     ESE skill en este submit, final=valor final del skill (post-cap 100).
' ============================================================================

' ============================================================================
' Fix 2026-09-12 -- D1/D2: deudas de la revision fresca de VB6 sobre la Ola 5
' tajada B (bitacora, bloque "Dos advertencias que la revision dejo COMO DEUDA
' para la Ola 6"), cerradas antes del deploy a la VM de beta.
' ============================================================================
' D1 -- overflow de la tabla de 256 peleas activas perdia el fight_end de la
' pelea reciclada. FreeElementalFightSlot() elegia el slot mas viejo por
' LastTickMs y lo devolvia tal cual; el llamador (ElementalBalanceFightId) lo
' reseteaba y arrancaba la pelea nueva ahi SIN pasar antes por
' LogElementalFightEnd -- sus contadores en memoria se perdian y el archivo
' quedaba con un fight_start huerfano indistinguible de una desconexion real.
' Fix: FreeElementalFightSlot() cierra el slot reciclado ANTES de devolverlo,
' con el mismo LogElementalFightEnd que ya usa el camino de timeout, pero con
' un outcome propio (ELEMENTAL_FIGHT_OUTCOME_EVICTED = "evicted") para que el
' lector (tools/analizar_telemetria.py) distinga "se quedo sin lugar" de
' "gano/perdio" o "se enfrio". mNextFightId sigue incrementando de forma
' GLOBAL (no es un numero de slot): el fight_id de la pelea evictada NO se
' repite con el de la pelea nueva que toma el slot -- solo el indice FISICO
' del array se recicla (eso ya era asi antes del fix; se documenta aca porque
' es lo que garantiza que cerrar el slot reciclado no invalida ningun fight_id
' ya entregado). Los contadores de la pelea nueva arrancan en cero igual que
' siempre: ResetElementalFightSlot() se sigue llamando DESPUES de este cierre,
' desde ElementalBalanceFightId.
' D2 -- LogElementalFightEnd() barria UserList linealmente DOS veces (una por
' lado) en CADA cierre de pelea, no solo en el timeout: toda muerte PvP cierra
' con closeFight:=True (Modulo_UsUaRiOs.bas, bloque de muerte), asi que el
' barrido se pagaba en el hot path de cada kill, no solo en el caso raro del
' timeout de 30s. Con MaxUsers=1000 (Server.ini de este deployment) son hasta
' 2000 comparaciones por cierre de pelea; razonado (no medido con profiler en
' vivo sobre PvP real, ver bitacora del 2026-09-12) como un costo chico en una
' muerte aislada pero acumulable bajo PvP masivo en la VM -- exactamente el
' escenario que motiva cerrar esta deuda ahora. Fix: el slot ya conoce
' CharA/CharB desde que nace la pelea; ahora tambien guarda el UserIndex de
' cada lado en ese momento (SnapshotElementalFightStart, mismo Optional que ya
' trae el idx). Al cerrar, ResolveFightUserIndex() usa ese indice cacheado
' SOLO si todavia apunta al MISMO char_id (UserLogged=True Y .Id=charId) -- un
' slot que se reciclo a otro jugador entre el ultimo contacto y el cierre
' (reconexion con otro UserIndex, por ejemplo) no puede colarse: cae al
' barrido de FindUserIndexByCharId de siempre, exactamente el mismo camino que
' ya existia. Ninguna estructura global nueva: el cache vive en el mismo
' t_ElementalFight que ya existia. Los 7 call sites viejos de componentes
' elementales que no pasan idx (Optional, degradacion ya documentada arriba)
' dejan UserIndexA/B en 0 -- ResolveFightUserIndex cae al barrido igual que
' antes, sin regresion.
' ============================================================================

Option Explicit

Private Const ELEMENTAL_BALANCE_LOG_HEADER As String = "ts_ms;event;attacker;attacker_class;victim;victim_type;item;tier;dmg_type;raw;final;resist_pct;pvp;map;char_id;account_id;victim_char_id;src_item;fight_id;schema;run_id"
Private Const ELEMENTAL_BALANCE_LOG_SCHEMA_VERSION As Long = 3

' Catalogo elemental jugable (plan 04.003/07.001): OBJ9084-9161, 78 items.
Private Const ELEMENTAL_CATALOG_MIN As Long = 9084
Private Const ELEMENTAL_CATALOG_MAX As Long = 9161

' Decision 12 (plan 10.001): ventana de silencio que cierra una pelea sin
' muerte de por medio. 30s: corto para no fundir dos peleas distintas,
' largo para tolerar una persecucion o una pausa por pocion.
Private Const ELEMENTAL_FIGHT_SILENCE_MS As Long = 30000
Private Const MAX_ACTIVE_ELEMENTAL_FIGHTS As Long = 256

' D1 (fix 2026-09-12): outcome que cierra una pelea reciclada por overflow de
' la tabla (los 256 slots activos), distinto de "death"/"timeout" para que el
' lector (tools/analizar_telemetria.py) pueda distinguir un cierre por falta
' de espacio de un cierre real de combate.
Private Const ELEMENTAL_FIGHT_OUTCOME_EVICTED As String = "evicted"

Private Type t_ElementalFight
    CharA As Long
    CharB As Long
    FightId As Long
    LastTickMs As Long
    Active As Boolean
    ' --- Agregados por pelea (Ola 5 tajada B, punto 3b prioridad 1) ---
    StartTickMs As Long
    AccountA As Long
    AccountB As Long
    ClassA As String
    ClassB As String
    RaceA As Long
    RaceB As Long
    LevelA As Long
    LevelB As Long
    FactionA As Long
    FactionB As Long
    StartMap As Integer
    StartHpA As Long
    StartHpB As Long
    SwingsA As Long
    SwingsB As Long
    HitsA As Long
    HitsB As Long
    MissA As Long
    MissB As Long
    BlocksA As Long
    BlocksB As Long
    PotionsRedA As Long
    PotionsRedB As Long
    PotionsBlueA As Long
    PotionsBlueB As Long
    CastsA As Long
    CastsB As Long
    ' D2 (fix 2026-09-12): UserIndex cacheado de cada lado, para evitar el
    ' barrido lineal de UserList al cerrar la pelea (ver ResolveFightUserIndex
    ' mas abajo). Nunca se usa a ciegas: se valida contra CharA/CharB antes.
    UserIndexA As Integer
    UserIndexB As Integer
End Type

' Resultado de un swing fisico/a distancia PvP (ElementalBalanceFightSwing).
Public Enum e_ElementalSwingOutcome
    eSwingMiss = 0
    eSwingHit = 1
    eSwingBlock = 2
End Enum

' Esquema del archivo separado de agregados por pelea (ver bloque de
' documentacion "Motor de agregados por pelea" mas arriba).
Private Const ELEMENTAL_FIGHT_LOG_HEADER As String = "ts_ms;event;fight_id;char_a;account_a;char_b;account_b;class_a;class_b;race_a;race_b;level_a;level_b;faction_a;faction_b;map;start_hp_a;start_hp_b;end_hp_a;end_hp_b;duration_ms;outcome;swings_a;swings_b;hits_a;hits_b;miss_a;miss_b;blocks_a;blocks_b;potions_red_a;potions_red_b;potions_blue_a;potions_blue_b;casts_a;casts_b;schema;run_id"
Private Const ELEMENTAL_FIGHT_LOG_SCHEMA_VERSION As Long = 2

Private mFights(1 To MAX_ACTIVE_ELEMENTAL_FIGHTS) As t_ElementalFight
Private mNextFightId As Long

' Identificador de RUN (fix 2026-09-12, ver bloque de documentacion arriba):
' fijo la PRIMERA vez que se pide, se mantiene igual el resto del proceso.
Private mRunId As String
Private mRunIdReady As Boolean

Private Function ElementalBalanceRunId() As String
    If Not mRunIdReady Then
        mRunId = Format$(Now, "yyyymmddhhnnss")
        mRunIdReady = True
    End If
    ElementalBalanceRunId = mRunId
End Function

Public Function ElementalBalanceLogEnabled() As Boolean
    ElementalBalanceLogEnabled = IsFeatureEnabled("elemental_balance_log")
End Function

' TOGGLE33 (plan 10.001, decision 7): telemetria de eventos de jugador,
' independiente del log de calibracion por golpe. Ver el bloque "Toggle de
' jugadores" arriba del todo de este archivo.
Public Function ElementalPlayerTelemetryEnabled() As Boolean
    ElementalPlayerTelemetryEnabled = IsFeatureEnabled("elemental_player_telemetry")
End Function

' Arma el identificador "U<idx>"/"N<idx>" para attacker/victim. "0" si idx<=0.
Public Function ElementalBalanceActorId(ByVal isNpc As Boolean, ByVal idx As Integer) As String
    If idx <= 0 Then
        ElementalBalanceActorId = "0"
    ElseIf isNpc Then
        ElementalBalanceActorId = "N" & idx
    Else
        ElementalBalanceActorId = "U" & idx
    End If
End Function

' Nombre de clase (ListaClases) si el actor es user; "0" si es NPC o el
' indice es invalido (con guard de error: un UserIndex desconectado entre el
' momento del golpe y el logueo no debe tirar la fila entera).
Public Function ElementalBalanceActorClass(ByVal isNpc As Boolean, ByVal idx As Integer) As String
    On Error GoTo eh
    If isNpc Or idx <= 0 Then
        ElementalBalanceActorClass = "0"
        Exit Function
    End If
    ElementalBalanceActorClass = ListaClases(UserList(idx).clase)
    Exit Function
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.ElementalBalanceActorClass", Erl)
    ElementalBalanceActorClass = "0"
End Function

' Mapa del actor (target o atacante). 0 si el indice es invalido.
Public Function ElementalBalanceMap(ByVal isNpc As Boolean, ByVal idx As Integer) As Integer
    On Error GoTo eh
    If idx <= 0 Then Exit Function
    If isNpc Then
        ElementalBalanceMap = NpcList(idx).pos.Map
    Else
        ElementalBalanceMap = UserList(idx).pos.Map
    End If
    Exit Function
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.ElementalBalanceMap", Erl)
    ElementalBalanceMap = 0
End Function

' Id estable de personaje (Hecho 23: UserList(idx).Id, NO el UserIndex/slot
' que se reusa entre jugadores). 0 si es NPC, el indice es invalido, o el
' slot no tiene hoy un user logueado (evita atribuir a un slot reciclado).
Public Function ElementalBalanceCharId(ByVal isNpc As Boolean, ByVal idx As Integer) As Long
    On Error GoTo eh
    If isNpc Or idx <= 0 Then Exit Function
    If idx > UBound(UserList) Then Exit Function
    If Not UserList(idx).flags.UserLogged Then Exit Function
    ElementalBalanceCharId = UserList(idx).Id
    Exit Function
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.ElementalBalanceCharId", Erl)
    ElementalBalanceCharId = 0
End Function

' Id estable de cuenta (Hecho 23: UserList(idx).AccountID). Mismo criterio
' que ElementalBalanceCharId.
Public Function ElementalBalanceAccountId(ByVal isNpc As Boolean, ByVal idx As Integer) As Long
    On Error GoTo eh
    If isNpc Or idx <= 0 Then Exit Function
    If idx > UBound(UserList) Then Exit Function
    If Not UserList(idx).flags.UserLogged Then Exit Function
    ElementalBalanceAccountId = UserList(idx).AccountID
    Exit Function
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.ElementalBalanceAccountId", Erl)
    ElementalBalanceAccountId = 0
End Function

' B0 (plan 10.001, Ola 5 tajada B): guardia unica de identidad para eventos
' que SIEMPRE deben atribuirse a un jugador real y nunca a un slot que
' todavia no complete el login (o que ya lo abandono). Mismo criterio que
' ElementalBalanceCharId/AccountId (UserLogged=True), pero expuesto como
' condicion booleana para que un call site pueda usarlo como guard directo
' ("And modElementalBalanceLog.ElementalBalanceUserReady(UserIndex)") en vez
' de confiar en que un Optional como UserIsLoggingIn se pase siempre bien.
' Ver el bloque de documentacion "Ola 5, tajada B -- B0" al inicio del
' archivo para el caso real que motivo esto (InvUsuario.bas:1675).
Public Function ElementalBalanceUserReady(ByVal UserIndex As Integer) As Boolean
    On Error GoTo eh
    If UserIndex <= 0 Then Exit Function
    If UserIndex > UBound(UserList) Then Exit Function
    ElementalBalanceUserReady = UserList(UserIndex).flags.UserLogged
    Exit Function
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.ElementalBalanceUserReady", Erl)
    ElementalBalanceUserReady = False
End Function

' True si el ObjIndex cae en el catalogo elemental jugable (OBJ9084-9161,
' plan 04.003). Usado para filtrar craft/equip "solo catalogo" (plan 10.001,
' Ola 5 punto 3) SIN depender de si ese item tiene un tier explicito: hay
' items del catalogo (orbes, flechas, pergaminos) sin T1/T2 declarado en
' docs/catalogo-elemental-jugable.md pero que igual son del catalogo.
Public Function ElementalBalanceInCatalog(ByVal objIndex As Long) As Boolean
    ElementalBalanceInCatalog = (objIndex >= ELEMENTAL_CATALOG_MIN And objIndex <= ELEMENTAL_CATALOG_MAX)
End Function

' Tier (1/2) SOLO para los rangos que docs/catalogo-elemental-jugable.md
' etiqueta explicitamente "Arma/Anillo/Tunica/Armadura T1" o "T2". El resto
' del catalogo (flechas 9096-9100, Aceites Superiores 9106-9110, orbes
' 9111-9115, gear de espinas 9130-9132, veneno-aceites 9133-9135,
' antidotos 9136-9137, pergaminos 9138-9142/9161) no tiene una T explicita
' en el doc: se devuelve 0 (no disponible) en vez de adivinar. Decision
' documentada en la bitacora de la Ola 5 (plan 10.001).
Public Function ElementalBalanceCatalogTier(ByVal objIndex As Long) As Long
    Select Case objIndex
        Case 9084 To 9089            ' Armas T1
            ElementalBalanceCatalogTier = 1
        Case 9090 To 9095            ' Armas T2
            ElementalBalanceCatalogTier = 2
        Case 9101 To 9105            ' Aceites T1
            ElementalBalanceCatalogTier = 1
        Case 9116 To 9121            ' Anillos T1
            ElementalBalanceCatalogTier = 1
        Case 9122 To 9127            ' Anillos T2
            ElementalBalanceCatalogTier = 2
        Case 9128 To 9129            ' Amuletos T2
            ElementalBalanceCatalogTier = 2
        Case 9143 To 9148            ' Tunicas T1
            ElementalBalanceCatalogTier = 1
        Case 9149 To 9154            ' Armaduras T1
            ElementalBalanceCatalogTier = 1
        Case 9155 To 9157            ' Tunicas T2
            ElementalBalanceCatalogTier = 2
        Case 9158 To 9160            ' Armaduras T2
            ElementalBalanceCatalogTier = 2
        Case Else
            ElementalBalanceCatalogTier = 0
    End Select
End Function

' Grupo 7 (plan 10.001, punto 7b): HECHIZO414-419 son las 6 recetas T2/T3 del
' catalogo elemental (verificado en Hechizos.dat, ver bloque de documentacion
' arriba). Usado por recipe_learned (modHechizos.AgregarHechizo) y
' quest_complete (ModQuest.FinishQuest, para acotar a las quests de receta).
Public Function ElementalBalanceIsRecipeSpell(ByVal hIndex As Integer) As Boolean
    ElementalBalanceIsRecipeSpell = (hIndex >= 414 And hIndex <= 419)
End Function

' Grupo 7: craft_fail compartido por los 4 oficios (Herrero/Carpintero/
' Alquimista/Sastre). Reason: 0=materiales, 1=skill, 2=receta, 3=herramienta.
' Cheap early exit propio (toggle + catalogo) para no repetirlo en cada uno
' de los 4 call sites.
Public Sub LogCraftFail(ByVal UserIndex As Integer, ByVal ItemIndex As Long, ByVal Reason As Long)
    If Not ElementalPlayerTelemetryEnabled() Then Exit Sub
    If Not ElementalBalanceInCatalog(ItemIndex) Then Exit Sub
    Call LogElementalBalance("craft_fail", ElementalBalanceActorId(False, UserIndex), ElementalBalanceActorClass(False, UserIndex), "0", "none", ItemIndex, ElementalBalanceCatalogTier(ItemIndex), Reason, 0, 0, ElementalBalanceMap(False, UserIndex), ElementalBalanceCharId(False, UserIndex), ElementalBalanceAccountId(False, UserIndex), 0, 0, 0)
End Sub

' Grupo 7: effect_end de los 3 venenos nuevos (Minor/Hemo/Neuro), target SIEMPRE
' user (telemetria de jugador). SourceValid=False -> ataca/victima sin fuente
' resoluble (0/"0", igual criterio que el resto del modulo). item=EotId (NO
' ObjIndex, excepcion documentada arriba). Reason: 0=expiro,1=curado,2=muerte.
Public Sub LogPoisonEffectEnd(ByVal TargetUserIndex As Integer, ByVal SourceIndex As Integer, ByVal SourceIsNpc As Boolean, ByVal SourceValid As Boolean, ByVal EotId As Long, ByVal ServedMs As Long, ByVal StacksAtEnd As Long, ByVal Reason As Long)
    On Error GoTo eh
    If Not ElementalPlayerTelemetryEnabled() Then Exit Sub
    Dim att As String, attCls As String, srcCharId As Long, srcAccountId As Long
    If SourceValid Then
        att = ElementalBalanceActorId(SourceIsNpc, SourceIndex)
        attCls = ElementalBalanceActorClass(SourceIsNpc, SourceIndex)
        srcCharId = ElementalBalanceCharId(SourceIsNpc, SourceIndex)
        srcAccountId = ElementalBalanceAccountId(SourceIsNpc, SourceIndex)
    Else
        att = "0"
        attCls = "0"
    End If
    Dim served As Long
    served = ServedMs
    If served < 0 Then served = 0
    Call LogElementalBalance("effect_end", att, attCls, ElementalBalanceActorId(False, TargetUserIndex), "user", EotId, 0, Reason, served, StacksAtEnd, ElementalBalanceMap(False, TargetUserIndex), srcCharId, srcAccountId, ElementalBalanceCharId(False, TargetUserIndex), 0, 0)
    Exit Sub
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.LogPoisonEffectEnd", Erl)
End Sub

' Grupo 7: barrido de muerte. Se llama ANTES de ClearEffectList(ClearForDeath:=
' True) en Modulo_UsUaRiOs.UserDie -- esa llamada es generica (borra TODOS los
' tipos de efecto sin distincion) asi que este Sub es el unico lugar que sabe
' identificar los 3 tipos de veneno nuevo y leer su duracion servida/stacks
' ANTES de que se pierdan. Solo lee: el borrado real lo sigue haciendo
' ClearEffectList sin cambios.
Public Sub LogPoisonEffectsEndOnDeath(ByVal UserIndex As Integer)
    On Error GoTo eh
    If Not ElementalPlayerTelemetryEnabled() Then Exit Sub
    Dim i As Long
    Dim mn As PoisonMinorEffect
    Dim hf As PoisonHemoEffect
    Dim nf As PoisonNeuroEffect
    Dim served As Long, stacksAtEnd As Long, isPoisonEffect As Boolean
    With UserList(UserIndex).EffectOverTime
        For i = 0 To .EffectCount - 1
            isPoisonEffect = True
            stacksAtEnd = 0
            Select Case .EffectList(i).TypeId
                Case e_EffectOverTimeType.ePoisonMinor
                    Set mn = .EffectList(i)
                    served = mn.TelemetryServedMs
                Case e_EffectOverTimeType.ePoisonHemo
                    Set hf = .EffectList(i)
                    served = hf.TelemetryServedMs
                    stacksAtEnd = hf.TelemetryPeakStacks
                Case e_EffectOverTimeType.ePoisonNeuro
                    Set nf = .EffectList(i)
                    served = nf.TelemetryServedMs
                Case Else
                    isPoisonEffect = False
            End Select
            If isPoisonEffect Then
                Call LogPoisonEffectEnd(UserIndex, .EffectList(i).CasterArrayIndex, .EffectList(i).CasterRefType = eNpc, .EffectList(i).CasterIsValid, .EffectList(i).EotId, served, stacksAtEnd, 2)
            End If
        Next i
    End With
    Exit Sub
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.LogPoisonEffectsEndOnDeath", Erl)
End Sub

Private Function FindElementalFightSlot(ByVal a As Long, ByVal b As Long) As Long
    Dim i As Long
    For i = 1 To MAX_ACTIVE_ELEMENTAL_FIGHTS
        If mFights(i).Active Then
            If (mFights(i).CharA = a And mFights(i).CharB = b) Or (mFights(i).CharA = b And mFights(i).CharB = a) Then
                FindElementalFightSlot = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function FreeElementalFightSlot() As Long
    Dim i As Long
    For i = 1 To MAX_ACTIVE_ELEMENTAL_FIGHTS
        If Not mFights(i).Active Then
            FreeElementalFightSlot = i
            Exit Function
        End If
    Next i
    ' Tabla llena (raro: 256 pares de PvP simultaneos): reusar el slot mas
    ' viejo por LastTickMs en vez de perder el evento nuevo.
    Dim oldest As Long, oldestTick As Long, nowTick As Long
    oldest = 1
    nowTick = GetTickCountRaw()
    oldestTick = mFights(1).LastTickMs
    For i = 2 To MAX_ACTIVE_ELEMENTAL_FIGHTS
        If TickIsOlder(nowTick, mFights(i).LastTickMs, oldestTick) Then
            oldestTick = mFights(i).LastTickMs
            oldest = i
        End If
    Next i
    ' D1 (fix 2026-09-12): cerrar el slot reciclado ANTES de devolverlo, igual
    ' que el camino de timeout de ElementalBalanceFightId -- si no, sus
    ' contadores se pierden en memoria y el archivo queda con un fight_start
    ' huerfano indistinguible de una desconexion real. El llamador sigue
    ' llamando ResetElementalFightSlot() sobre este mismo indice antes de
    ' escribir la pelea nueva, asi que sus contadores arrancan en cero.
    Call LogElementalFightEnd(oldest, ELEMENTAL_FIGHT_OUTCOME_EVICTED)
    FreeElementalFightSlot = oldest
End Function

' Identificador de pelea (decision 12, plan 10.001): par {A,B} de char_id
' ESTABLES (no UserIndex). Arranca en el primer golpe entre los dos, se
' mantiene mientras haya eventos separados por menos de
' ELEMENTAL_FIGHT_SILENCE_MS, y se asigna en el SERVER (no se reconstruye en
' el lector). Devuelve 0 si cualquiera de los dos lados no es un personaje
' identificado (charA<=0 Or charB<=0): B4 es tiempo-hasta-matar PvP, un NPC
' no arma "pelea". closeFight=True cierra el par (llamado desde kill/death,
' decision 12: "cierra con la muerte de uno").
Public Function ElementalBalanceFightId(ByVal charA As Long, ByVal charB As Long, Optional ByVal closeFight As Boolean = False, Optional ByVal idxAForSnapshot As Integer = 0, Optional ByVal idxBForSnapshot As Integer = 0) As Long
    On Error GoTo eh
    If charA <= 0 Or charB <= 0 Then Exit Function
    Dim nowTick As Long
    nowTick = GetTickCountRaw()
    Dim slot As Long
    slot = FindElementalFightSlot(charA, charB)
    If slot > 0 Then
        If DeadlinePassed(nowTick, AddMod32(mFights(slot).LastTickMs, ELEMENTAL_FIGHT_SILENCE_MS)) Then
            ' Silencio > ventana: la pelea anterior ya se cerro sola: esta cuenta como una nueva.
            ' Ola 5 tajada B: antes de resetear, cerramos la vieja con fight_end (outcome
            ' "timeout") si llego a acumular algo (deteccion perezosa: solo se nota si el
            ' MISMO par vuelve a cruzarse; ver "Motor de agregados por pelea" en la cabecera).
            Call LogElementalFightEnd(slot, "timeout")
            mFights(slot).Active = False
            slot = 0
        End If
    End If
    If slot = 0 Then
        slot = FreeElementalFightSlot()
        Call ResetElementalFightSlot(slot)
        mNextFightId = mNextFightId + 1
        mFights(slot).CharA = charA
        mFights(slot).CharB = charB
        mFights(slot).FightId = mNextFightId
        mFights(slot).Active = True
        mFights(slot).StartTickMs = nowTick
        Call SnapshotElementalFightStart(slot, idxAForSnapshot, idxBForSnapshot)
        Call LogElementalFightStart(slot)
    End If
    mFights(slot).LastTickMs = nowTick
    ElementalBalanceFightId = mFights(slot).FightId
    If closeFight Then
        Call LogElementalFightEnd(slot, "death")
        mFights(slot).Active = False
    End If
    Exit Function
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.ElementalBalanceFightId", Erl)
End Function

Private Sub ResetElementalFightSlot(ByVal slot As Long)
    Dim blank As t_ElementalFight
    mFights(slot) = blank
End Sub

' Completa nivel/clase/raza/faccion/HP/mapa de ambos lados AL CREAR la pelea,
' si el caller paso los UserIndex (swings y casts los pasan siempre; los 7
' call sites viejos de componentes elementales no los pasan -- Optional,
' quedan en blanco/0, degradacion documentada en la cabecera del archivo).
Private Sub SnapshotElementalFightStart(ByVal slot As Long, ByVal idxA As Integer, ByVal idxB As Integer)
    On Error GoTo eh
    With mFights(slot)
        If idxA > 0 Then
            If idxA <= UBound(UserList) Then
                If UserList(idxA).flags.UserLogged Then
                    .UserIndexA = idxA
                    .AccountA = UserList(idxA).AccountID
                    .ClassA = ListaClases(UserList(idxA).clase)
                    .RaceA = UserList(idxA).raza
                    .LevelA = UserList(idxA).Stats.ELV
                    .FactionA = UserList(idxA).Faccion.Status
                    .StartHpA = UserList(idxA).Stats.MinHp
                    .StartMap = UserList(idxA).pos.Map
                End If
            End If
        End If
        If idxB > 0 Then
            If idxB <= UBound(UserList) Then
                If UserList(idxB).flags.UserLogged Then
                    .UserIndexB = idxB
                    .AccountB = UserList(idxB).AccountID
                    .ClassB = ListaClases(UserList(idxB).clase)
                    .RaceB = UserList(idxB).raza
                    .LevelB = UserList(idxB).Stats.ELV
                    .FactionB = UserList(idxB).Faccion.Status
                    .StartHpB = UserList(idxB).Stats.MinHp
                    If .StartMap = 0 Then .StartMap = UserList(idxB).pos.Map
                End If
            End If
        End If
    End With
    Exit Sub
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.SnapshotElementalFightStart", Erl)
End Sub

' Busca una pelea ACTIVA por uno solo de sus dos lados (para pociones: nunca
' crean pelea, solo suman si el jugador ya esta en una).
Private Function FindElementalFightSlotByChar(ByVal charId As Long) As Long
    Dim i As Long
    For i = 1 To MAX_ACTIVE_ELEMENTAL_FIGHTS
        If mFights(i).Active Then
            If mFights(i).CharA = charId Or mFights(i).CharB = charId Then
                FindElementalFightSlotByChar = i
                Exit Function
            End If
        End If
    Next i
End Function

' Reconstruye el UserIndex actual a partir de un char_id estable (para leer
' el HP "final" de una pelea al cerrarla). Barrido lineal: solo se llama al
' CERRAR una pelea (raro), nunca en el hot path. 0 si el char ya no esta
' logueado (por ejemplo, se desconecto antes de que la pelea cerrara por
' timeout): el HP final queda en 0/no disponible, documentado en el archivo.
Private Function FindUserIndexByCharId(ByVal charId As Long) As Integer
    On Error GoTo eh
    If charId <= 0 Then Exit Function
    Dim i As Integer
    For i = 1 To UBound(UserList)
        If UserList(i).flags.UserLogged Then
            If UserList(i).Id = charId Then
                FindUserIndexByCharId = i
                Exit Function
            End If
        End If
    Next i
    Exit Function
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.FindUserIndexByCharId", Erl)
End Function

' D2 (fix 2026-09-12): usa el UserIndex cacheado en el slot (SnapshotElemental
' FightStart) si TODAVIA apunta al mismo char_id -- evita el barrido lineal de
' UserList en el camino comun. Cae a FindUserIndexByCharId (el barrido de
' siempre) si el cache esta vacio (call site viejo que no paso idx) o quedo
' obsoleto (el slot se reciclo a otro jugador, ej. reconexion con otro
' indice). Sin estructura global nueva: solo lee el cache que ya vive en
' t_ElementalFight.
Private Function ResolveFightUserIndex(ByVal cachedIdx As Integer, ByVal charId As Long) As Integer
    On Error GoTo eh
    If cachedIdx > 0 Then
        If cachedIdx <= UBound(UserList) Then
            If UserList(cachedIdx).flags.UserLogged Then
                If UserList(cachedIdx).Id = charId Then
                    ResolveFightUserIndex = cachedIdx
                    Exit Function
                End If
            End If
        End If
    End If
    ResolveFightUserIndex = FindUserIndexByCharId(charId)
    Exit Function
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.ResolveFightUserIndex", Erl)
    ResolveFightUserIndex = FindUserIndexByCharId(charId)
End Function

Private Function ElementalFightLogFileName() As String
    ElementalFightLogFileName = App.Path & "\Logs\ElementalFights_" & Format$(Date, "yyyy-mm-dd") & "_v2.log"
End Function

Private Sub WriteElementalFightRow(ByVal evento As String, ByVal slot As Long, ByVal outcome As String, ByVal endHpA As Long, ByVal endHpB As Long, ByVal durationMs As Long)
    On Error GoTo ErrHandler
    If Not ElementalPlayerTelemetryEnabled() Then Exit Sub
    Dim fname As String
    fname = ElementalFightLogFileName()
    Dim fnum As Integer
    fnum = FreeFile
    Dim writeHeader As Boolean
    writeHeader = (LenB(Dir(fname)) = 0)
    Open fname For Append As #fnum
    If writeHeader Then Print #fnum, ELEMENTAL_FIGHT_LOG_HEADER
    With mFights(slot)
        Print #fnum, TickToCsv(GetTickCountRaw()) & ";" & evento & ";" & .FightId & ";" & _
            .CharA & ";" & .AccountA & ";" & .CharB & ";" & .AccountB & ";" & _
            .ClassA & ";" & .ClassB & ";" & .RaceA & ";" & .RaceB & ";" & _
            .LevelA & ";" & .LevelB & ";" & .FactionA & ";" & .FactionB & ";" & _
            .StartMap & ";" & .StartHpA & ";" & .StartHpB & ";" & endHpA & ";" & endHpB & ";" & _
            durationMs & ";" & outcome & ";" & _
            .SwingsA & ";" & .SwingsB & ";" & .HitsA & ";" & .HitsB & ";" & .MissA & ";" & .MissB & ";" & _
            .BlocksA & ";" & .BlocksB & ";" & .PotionsRedA & ";" & .PotionsRedB & ";" & _
            .PotionsBlueA & ";" & .PotionsBlueB & ";" & .CastsA & ";" & .CastsB & ";" & ELEMENTAL_FIGHT_LOG_SCHEMA_VERSION & ";" & ElementalBalanceRunId()
    End With
    Close #fnum
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.WriteElementalFightRow", Erl)
    On Error Resume Next
    Close #fnum
End Sub

Private Sub LogElementalFightStart(ByVal slot As Long)
    Call WriteElementalFightRow("fight_start", slot, "0", 0, 0, 0)
End Sub

' outcome: "death" (cierre por muerte, closeFight:=True) o "timeout" (el
' mismo par volvio a cruzarse despues de los 30s de silencio).
Private Sub LogElementalFightEnd(ByVal slot As Long, ByVal outcome As String)
    On Error GoTo eh
    Dim idxA As Integer, idxB As Integer, endHpA As Long, endHpB As Long, durationMs As Long
    idxA = ResolveFightUserIndex(mFights(slot).UserIndexA, mFights(slot).CharA)
    idxB = ResolveFightUserIndex(mFights(slot).UserIndexB, mFights(slot).CharB)
    If idxA > 0 Then endHpA = UserList(idxA).Stats.MinHp
    If idxB > 0 Then endHpB = UserList(idxB).Stats.MinHp
    ' Wrap-safe (plan 21.003): FightDurationMs usa TicksElapsed (aritmetica sin signo mod 2^32),
    ' asi que una pelea que cruce &H7FFFFFFF o el wrap de timeGetTime (~49.7 dias) da la
    ' duracion real. Antes era una resta con signo que desbordaba (error 6) en ese cruce y el
    ' handler eh perdia la fila fight_end. La duracion se acota a 2147483647 ms.
    durationMs = FightDurationMs(GetTickCountRaw(), mFights(slot).StartTickMs)
    Call WriteElementalFightRow("fight_end", slot, outcome, endHpA, endHpB, durationMs)
    Exit Sub
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.LogElementalFightEnd", Erl)
End Sub

' --- Hooks de agregados: llamados desde SistemaCombate.bas/InvUsuario.bas/
' modHechizos.bas. Los tres salen temprano y barato si el toggle esta apagado. ---

' Un swing fisico/a distancia entre dos users identificados (SistemaCombate.
' UsuarioAtacaUsuario). Crea la pelea si hace falta (con snapshot completo:
' ES el camino principal de creacion de una pelea, a diferencia de los call
' sites viejos de componentes elementales).
Public Sub ElementalBalanceFightSwing(ByVal attackerIdx As Integer, ByVal victimIdx As Integer, ByVal outcome As e_ElementalSwingOutcome)
    On Error GoTo eh
    If Not ElementalPlayerTelemetryEnabled() Then Exit Sub
    Dim charA As Long, charB As Long
    charA = ElementalBalanceCharId(False, attackerIdx)
    charB = ElementalBalanceCharId(False, victimIdx)
    If charA <= 0 Or charB <= 0 Then Exit Sub
    Call ElementalBalanceFightId(charA, charB, False, attackerIdx, victimIdx)
    Dim slot As Long
    slot = FindElementalFightSlot(charA, charB)
    If slot = 0 Then Exit Sub
    With mFights(slot)
        If .CharA = charA Then
            .SwingsA = .SwingsA + 1
            Select Case outcome
                Case eSwingHit: .HitsA = .HitsA + 1
                Case eSwingBlock: .BlocksA = .BlocksA + 1
                Case Else: .MissA = .MissA + 1
            End Select
        Else
            .SwingsB = .SwingsB + 1
            Select Case outcome
                Case eSwingHit: .HitsB = .HitsB + 1
                Case eSwingBlock: .BlocksB = .BlocksB + 1
                Case Else: .MissB = .MissB + 1
            End Select
        End If
    End With
    Exit Sub
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.ElementalBalanceFightSwing", Erl)
End Sub

' Pocion roja/azul consumida DURANTE una pelea activa (InvUsuario.UseInvItem).
' Nunca crea pelea: una pocion fuera de combate no es un dato de "pelea".
Public Sub ElementalBalanceFightPotion(ByVal UserIndex As Integer, ByVal potionKind As String)
    On Error GoTo eh
    If Not ElementalPlayerTelemetryEnabled() Then Exit Sub
    Dim charId As Long
    charId = ElementalBalanceCharId(False, UserIndex)
    If charId <= 0 Then Exit Sub
    Dim slot As Long
    slot = FindElementalFightSlotByChar(charId)
    If slot = 0 Then Exit Sub
    With mFights(slot)
        If .CharA = charId Then
            If potionKind = "red" Then .PotionsRedA = .PotionsRedA + 1 Else .PotionsBlueA = .PotionsBlueA + 1
        Else
            If potionKind = "red" Then .PotionsRedB = .PotionsRedB + 1 Else .PotionsBlueB = .PotionsBlueB + 1
        End If
    End With
    Exit Sub
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.ElementalBalanceFightPotion", Erl)
End Sub

' Hechizo HOSTIL contra un user identificado (modHechizos.LanzarHechizo). Si
' es el primer contacto entre los dos, arranca la pelea (un ataque a
' distancia con magia inicia un enfrentamiento igual que un swing).
Public Sub ElementalBalanceFightCast(ByVal casterIdx As Integer, ByVal targetIdx As Integer)
    On Error GoTo eh
    If Not ElementalPlayerTelemetryEnabled() Then Exit Sub
    Dim charA As Long, charB As Long
    charA = ElementalBalanceCharId(False, casterIdx)
    charB = ElementalBalanceCharId(False, targetIdx)
    If charA <= 0 Or charB <= 0 Then Exit Sub
    Call ElementalBalanceFightId(charA, charB, False, casterIdx, targetIdx)
    Dim slot As Long
    slot = FindElementalFightSlot(charA, charB)
    If slot = 0 Then Exit Sub
    With mFights(slot)
        If .CharA = charA Then .CastsA = .CastsA + 1 Else .CastsB = .CastsB + 1
    End With
    Exit Sub
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.ElementalBalanceFightCast", Erl)
End Sub

' Foto del equipo elemental al completar el login (Modulo_UsUaRiOs.
' ConnectUser_Complete). Una fila "gear_snapshot" por cada slot de catalogo
' equipado. Requiere ElementalBalanceUserReady (B0): sin login completo, no
' hay char_id que atribuir.
Public Sub ElementalBalanceLogGearSnapshot(ByVal UserIndex As Integer)
    On Error GoTo eh
    If Not ElementalPlayerTelemetryEnabled() Then Exit Sub
    If Not ElementalBalanceUserReady(UserIndex) Then Exit Sub
    Dim slots(1 To 6) As Long
    slots(1) = UserList(UserIndex).invent.EquippedWeaponObjIndex
    slots(2) = UserList(UserIndex).invent.EquippedShieldObjIndex
    slots(3) = UserList(UserIndex).invent.EquippedHelmetObjIndex
    slots(4) = UserList(UserIndex).invent.EquippedArmorObjIndex
    slots(5) = UserList(UserIndex).invent.EquippedRingAccesoryObjIndex
    slots(6) = UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex
    Dim i As Long
    For i = 1 To 6
        If ElementalBalanceInCatalog(slots(i)) Then
            Call LogElementalBalance("gear_snapshot", ElementalBalanceActorId(False, UserIndex), ElementalBalanceActorClass(False, UserIndex), "0", "none", slots(i), ElementalBalanceCatalogTier(slots(i)), 0, 0, 0, ElementalBalanceMap(False, UserIndex), ElementalBalanceCharId(False, UserIndex), ElementalBalanceAccountId(False, UserIndex), 0, 0, 0)
        End If
    Next i
    Exit Sub
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.ElementalBalanceLogGearSnapshot", Erl)
End Sub

' Nombre del archivo del dia para el esquema v2 (Hecho 26: archivo por
' version de esquema, ver el bloque de comentarios "Gotcha de compatibilidad
' del parser" arriba).
Private Function ElementalBalanceLogFileName() As String
    ElementalBalanceLogFileName = App.Path & "\Logs\ElementalBalance_" & Format$(Date, "yyyy-mm-dd") & "_v3.log"
End Function

Public Sub LogElementalBalance(ByVal evento As String, _
                                ByVal attacker As String, _
                                ByVal attackerClass As String, _
                                ByVal victim As String, _
                                ByVal victimType As String, _
                                ByVal itemObjIndex As Long, _
                                ByVal tier As Long, _
                                ByVal dmgType As Long, _
                                ByVal raw As Long, _
                                ByVal finalDmg As Long, _
                                ByVal mapNumber As Integer, _
                                ByVal charId As Long, _
                                ByVal accountId As Long, _
                                ByVal victimCharId As Long, _
                                ByVal srcItem As Long, _
                                ByVal fightId As Long)
    On Error GoTo ErrHandler
    ' Red de seguridad barata (ver "IMPORTANTE" arriba): el gate PRECISO es
    ' responsabilidad de cada call site (su propio toggle). Esto solo evita
    ' escribir CUALQUIER cosa si los dos toggles estan apagados.
    If Not (ElementalBalanceLogEnabled() Or ElementalPlayerTelemetryEnabled()) Then Exit Sub
    Dim resistPct As Long
    If raw > 0 Then resistPct = (raw - finalDmg) * 100 \ raw
    Dim pvp As Byte
    If victimType = "user" Then pvp = 1
    Dim fname As String
    fname = ElementalBalanceLogFileName()
    Dim fnum As Integer
    fnum = FreeFile
    Dim writeHeader As Boolean
    writeHeader = (LenB(dir(fname)) = 0)
    Open fname For Append As #fnum
    If writeHeader Then
        Print #fnum, ELEMENTAL_BALANCE_LOG_HEADER
        ' Ola 5 tajada B, evento "header" (punto 3b, prioridad 6): primera
        ' fila de datos del archivo del dia. attacker = BuildStamp() (flags
        ' de compilacion; el server NO puede autoreportar el hash de su
        ' propio binario ni de obj.dat -- VB6 no trae SHA256 nativo. La
        ' correlacion fila<->binario sigue siendo cruzar ts_ms contra
        ' build-stamp.txt, como ya establecio la tajada A). raw/final =
        ' estado de los dos toggles al momento de abrir el archivo.
        Dim ebBuildFlags As String
        ebBuildFlags = Replace(BuildStamp(), ";", ",")
        Print #fnum, TickToCsv(GetTickCountRaw()) & ";header;" & ebBuildFlags & ";0;n/a;none;0;0;0;" & _
                     IIf(ElementalBalanceLogEnabled(), 1, 0) & ";" & IIf(ElementalPlayerTelemetryEnabled(), 1, 0) & ";0;0;0;0;0;0;0;0;" & ELEMENTAL_BALANCE_LOG_SCHEMA_VERSION & ";" & ElementalBalanceRunId()
    End If
    Print #fnum, TickToCsv(GetTickCountRaw()) & ";" & _
                 evento & ";" & attacker & ";" & attackerClass & ";" & _
                 victim & ";" & victimType & ";" & itemObjIndex & ";" & tier & ";" & _
                 dmgType & ";" & raw & ";" & finalDmg & ";" & resistPct & ";" & _
                 pvp & ";" & mapNumber & ";" & charId & ";" & accountId & ";" & _
                 victimCharId & ";" & srcItem & ";" & fightId & ";" & ELEMENTAL_BALANCE_LOG_SCHEMA_VERSION & ";" & ElementalBalanceRunId()
    Close #fnum
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.LogElementalBalance", Erl)
    On Error Resume Next
    Close #fnum
End Sub
