# Brief: soporte de calendarios en outlook-desktop-mcp (macOS)

Contexto traspasado desde una sesión previa que no tenía shell en esta máquina.
Fecha: 2026-09-22. Outlook para Mac 16.113.1.

## Objetivo

José tiene 23 calendarios repartidos en 2 cuentas Exchange y quiere reorganizar
sus eventos: los personales a su cuenta de Gmail, y los laborales repartidos
entre **Docencia**, **Investigación** y **Tutorías** (ojo: el calendario se llama
"Investigación", no "Research").

El servidor macOS ya **ve** los calendarios (lectura completa, ver abajo). Lo
que falta es **mover** eventos entre ellos.

## Estado del código (actualizado 2026-09-22, ya con shell)

Hecho e integrado en `server_mac.py`, con 20 tests nuevos en
`tests/unit_mac_test.py` (76/76 en verde) y verificado contra Outlook real:

- **`list_calendars()`** — id, nombre, cuenta y nº de eventos de los 21
  calendarios. Omite las 2 raíces sin nombre. Tarda 1,6 s.
- **`list_events(..., calendar_id=None)`** — acota a un calendario, y cada
  evento devuelve `calendar_id`, `calendar` e `is_recurring`.
- **`search_events(..., calendar_id=None)`** — igual.
- **`create_event(..., calendar_id=None)`** — `make new ... at calendar id N`,
  y confirma en qué calendario aterrizó.
- **`get_event`** — añade `calendar_id`, `calendar`, `is_recurring`.

Tres arreglos que salieron por el camino:

- `list_events` **ignoraba `start_date` y `end_date`**: los recibía, no
  filtraba nada y devolvía los primeros `count*3` de los 9381 globales. Ahora
  filtra de verdad. Lo mismo en `search_events`.
- El filtrado se hace con `whose` **dentro de Outlook**, no iterando en Python.
  Medido sobre Diverso (2393 eventos): **0,25 s con `whose` frente a 44 s**
  iterando, con recuentos idénticos. El comentario del código que decía que el
  `whose` con fechas "no es fiable en Outlook para Mac" es falso.
- Las fechas salían como texto en locale del sistema ("miércoles, 23 de
  septiembre de 2026"), que ni se ordena ni se parsea. Ahora **ISO 8601**,
  construido por componentes numéricos, como en Windows.

Sigue pendiente: transporte **stdio only**, y `update_event` no escribe la
pertenencia a un calendario.

## Inventario real

Cuentas Exchange: `jagalindo@us.es`, `jagalindo@outlook.com`

**us.es** — 131 Calendario (795), 132 Diverso (2393), 133 Docencia (212),
134 Investigación (269), 135 Tutorías (380), 136 Cumpleaños (0),
137 Reservas Seminarios LSI (746), 138 DiversoLab (0), 139 Eventos DiversoLab (9)

**outlook.com** — 146 Calendario (125), 147 Cumpleaños (13), 148 Gestión (63),
149 Google Principal (653), 150 Investigación compartido (0),
151 Real Betis Balompié (38), 152 Sevilla (47), 153 Tu familia (0)

**Sin cuenta atribuible** — 13 Calendario (0), 102 (raíz us.es), 185 (raíz),
202 Familia (1), 205 malawito@gmail.com (2330), 211 Investigación Compartido (1307)

La suma de todos da exactamente 9381, que es lo que devuelve `calendar events`.
Eso confirma que esa colección es global, no por calendario.

## Hechos establecidos por sondeo

- La colección `calendars` existe. `id of c` y `name of c` funcionan.
  **Hay nombres duplicados** ("Calendario" x3, "Cumpleaños" x2) → trabajar por id.
- `account of <calendar>` resuelve la cuenta en los calendarios Exchange, no en
  202/205/211 (ver trampas).
- **`calendar of <event>` es legible.** Se puede saber dónde está cada evento.
- **`make new calendar event at <calendario real> with properties {...}` funciona**
  y el evento aterriza donde se pide (verificado: apareció en Docencia).
- **102 y 185 son carpetas raíz, no calendarios**: `name` = `missing value`,
  0 eventos, y su `folder` enumera los calendarios hijos.

## Preguntas ya respondidas (sonda v4, ejecutada 2026-09-22)

1. **¿`set calendar of <event> to <calendar>`?** → **NO.** Falla igual entre dos
   calendarios reales (133 Docencia → 135 Tutorías), no solo contra una raíz:
   *"No puede ajustarse calendar of calendar event a calendar id 135"*.
   No hay reasignación directa. **La restricción de diseño de abajo aplica.**
2. **¿`duplicate <event> to <calendar>`?** → **SÍ funciona.** Ojo: la sonda
   imprimió que no, y se equivocaba. `duplicate` **no devuelve resultado**, así
   que `set cp to duplicate ev to dstCal` deja `cp` sin definir y el error salta
   al leer `calendar of cp`, no al duplicar. La copia sí aterriza. Para
   quedarse con ella hay que **buscarla después** en el destino, nunca usar el
   resultado del comando.
3. **Campos sobre eventos reales:**
   - `subject`, `class`, `all day flag`, `location` → OK.
   - `attendees` → `count` OK; `class of item 1` = `resource attendee` (los dos
     muestreados eran reuniones de Teams, o sea la sala). **`address of
     (email address of a)` no coacciona a texto.** Falta camino de acceso.
   - `organizer` → es **`text`**, no un objeto. No tiene `email address`.
   - `category` → es una **`list`**, confirmado. Hay que iterarla.
   - `recurrence` → la propiedad existe (`class rREc`), `class of` no coacciona
     a texto, pero **`(recurrence of e) is not missing value` sí sirve** para
     saber si el evento recurre. Es lo que usa `is_recurring`.

## `duplicate` es FIEL — medido 2026-09-22

Sonda: `tests/probe_duplicate_fidelity_mac.applescript`. Dos casos, porque la
recurrencia no se puede crear por AppleScript:

- **A, evento sintético** con asistente, categoría, sitio y cuerpo. El asistente
  fue `malawito@gmail.com`, la otra cuenta del propio José, así que no salió
  nada hacia terceros.
- **B, evento REAL recurrente** (133/12716 "IISSI2 - F1.32") con **0
  asistentes**, que por eso no envía nada. Se duplicó y se borró **solo la
  copia**; el original sigue ahí.

Resultado: **la copia conserva todo lo medible** — asunto, inicio/fin, sitio,
`all day`, cuerpo (con sus saltos de línea), organizador, nº de asistentes,
categorías **y la recurrencia** (B: `recurrence: SI` → `SI`). Y aterriza en el
calendario destino.

Así que `move_event` = `duplicate` + borrar el original es **fiel**, no mutila.

Un matiz que sí importa: al asistente que creé como `required attendee` Outlook
lo devuelve con `class` = **`resource attendee`**, igual que una sala de Teams.
O sea que **no se puede distinguir una persona de una sala** leyendo la clase.
Por eso la guarda tiene que ser sobre el recuento, no sobre el tipo.

## Trampas ya pisadas — no repetirlas

- **No muestrear `item 1` / `item 2` de `calendars`**: son 102 y 185, las raíces
  degeneradas. Usar ids explícitos (133 Docencia, 135 Tutorías).
- **No sondear campos sobre el evento de prueba recién creado**: no tiene
  organizador, asistentes ni categorías, así que todo sale "NO EXPUESTO" y no
  significa nada. Muestrear eventos reales, saltando los `ZZ-MCP-PROBE`.
- **Variables declaradas dentro de un bloque `tell application` no son visibles
  en bloques `tell` posteriores.** Da "La variable X no está definida", el script
  aborta y **se pierde toda la salida acumulada**. Ni `global` lo arregló.
  Solución: envolver el cuerpo en `try` y hacer la limpieza por asunto, sin
  variables. Esto tumbó dos iteraciones seguidas.
- **`category of e` probablemente devuelve una lista**, por eso
  `name of (category of e)` falla. Comprobar `class of` antes.
- Los eventos de prueba se llaman `ZZ-MCP-PROBE-BORRAR`. Barrer por asunto al
  terminar, en todos los calendarios.
- **`delete e` con `e` como variable de bucle sobre `calendar events of c`
  FALLA**: *"No puede obtenerse item 1 of every calendar event of calendar id
  133"*. Leer `subject of e` sí funciona, borrar no. Como la limpieza de la
  sonda lo envolvía en un `try` mudo, **decía "0 borrados" mientras dejaba todo
  puesto**: v2, v3 y v4 dejaron 3 eventos en Docencia y 3 en Tutorías, ya
  barridos. Se borra bien con referencia explícita:
  `delete (calendar event id N of calendar id M)`.
- **`account of <calendar>` no siempre resuelve.** En 131-153 da la cuenta
  Exchange con `name`/`email address`. En 202, 205 y 211 devuelve un objeto de
  `class class` del que no se saca nombre. Por eso `list_calendars` deja la
  cuenta vacía ahí en vez de fallar.
- **La colección de eventos no viene ordenada.** Los últimos items de Docencia
  son de 2022 y el item 4 es de 2026. Hay que ordenar en Python.

## Restricción de diseño — DECIDIDA por José (2026-09-22)

La pregunta 1 **es NO**, así que **mover un evento es duplicar + borrar**. En
eventos con asistentes eso envía una cancelación y luego una invitación nueva a
todo el mundo.

**Decisión de José: los eventos con asistentes NO se mueven.** `move_event` se
niega en seco si `count of attendees > 0`. Pendiente de que José diga si quiere
además un flag de escape para forzarlo; mientras no lo diga, **no hay flag**.

La guarda **es exigible**, verificado: `count of (attendees of e)` se leyó
correctamente en **los 6030 eventos** de los calendarios 132, 211 y 205, sin un
solo fallo. (`address of` sigue sin coaccionar a texto, pero para la guarda solo
hace falta el recuento.)

Censo de asistentes — cuánto bloquearía la guarda:

| Calendario | Eventos | Con asistentes | % |
|---|---|---|---|
| 132 Diverso | 2393 | 707 | 29,5 % |
| 211 Investigación Compartido | 1307 | 90 | 6,9 % |
| 205 malawito@gmail.com | 2330 | 250 | 10,7 % |

Ojo con el matiz: los asistentes muestreados eran `resource attendee`, es decir
**la sala de Teams**. Un evento que solo "tiene asistentes" porque reservó sala
también quedará bloqueado. Es el lado seguro, pero explica por qué el 29,5 % de
Diverso es tan alto.

## Diverso (132) vs Investigación Compartido (211): NO son el mismo

Comprobado volcando los 6030 eventos de 132, 211 y 205 y comparándolos.

- **Los alimenta un calendario de Google distinto cada uno**, y no comparten
  ninguna dirección:
  - 132 Diverso → `qmikl7cf80cgmduftne4so9csc@group.calendar.google.com` (699)
  - 211 Inv. Compartido → `j3lgpk3tg8d26f2t103109mmqs@group.calendar.google.com` (1268)
- **Solape de eventos: 2 de 1307/2393 (0,2 %)**, y sigue siendo 2 aflojando la
  comparación a solo fecha + asunto. No es un desfase de horas: son distintos.
- **Composición distinta**: Diverso es mayoritariamente Exchange de `us.es`
  (1693 de 2393, 71 %) con un trozo de Google (699, 29 %). Investigación
  Compartido es 97 % Google.
- **Rango temporal distinto**: Diverso 2022–2026; Inv. Compartido 2007–2026.

Trampa relacionada: hay **dos** calendarios de nombre casi idéntico. El 150
"Investigación compartido" (outlook.com) está **vacío**, 0 eventos. El que tiene
los 1307 es el 211 "Investigación Compartido", sin cuenta atribuible.

Y el 205 `malawito@gmail.com` es un tercero aparte: 2206 de sus 2330 eventos los
organiza una dirección `@gmail.com`. Es el Gmail personal, no un compartido.

## Herramientas — hecho

- `move_event(entry_id, target_calendar_id)` — **sin flag de escape**, como
  pidió José. Se niega si `count of attendees > 0`.
- `create_meeting(..., calendar_id=None)` — faltaba por coherencia con
  `create_event`.

Orden de seguridad dentro de `move_event`, que es lo que hay que no romper:

1. cuenta asistentes y **se planta ahí** si hay alguno (antes de duplicar nada);
2. duplica;
3. **localiza la copia por id** (no por el resultado de `duplicate`, que no
   devuelve nada): el id más alto del destino que no estaba antes;
4. **compara el asunto** de la copia con el del original;
5. solo entonces borra el original, por id explícito.

Si el paso 3 o el 4 fallan, no borra: devuelve `failed` y el original se queda.

Verificado en vivo contra Outlook: evento con asistente → `refused` y el evento
intacto en Docencia, sin copia suelta; evento sin asistentes → `moved`, una sola
copia y en el destino, con sitio y cuerpo conservados; mover a donde ya está →
`unchanged`.

## Decisiones pendientes de José (no son de código)

- **Qué está desordenado exactamente.** Docencia, Investigación y Tutorías ya
  existen con 212, 269 y 380 eventos. ¿Hay que repartir los cajones de sastre
  "Calendario" (795) y "Diverso" (2393), o corregir lo mal puesto dentro de los
  tres? Son 600 eventos frente a 3200.
- **Si el calendario 205 `malawito@gmail.com` (2330 eventos) sincroniza de
  verdad con Google.** Si sí, "los personales al Gmail" es un movimiento entre
  calendarios más y no una migración entre plataformas.

## Lo aprendido operando de verdad (2026-09-22, sesión larga)

Cosas que no estaban en el brief y que cuestan caro descubrir dos veces:

- **Gmail (205) y el compartido (211) ignoran `duplicate` en silencio.** No dan
  error y no crean nada. `make new` sí funciona en ambos. Por eso `move_event`
  reconstruye el evento cuando la copia no aparece, y lo declara en la respuesta
  (`method: "recreated"`). Se niega a reconstruir eventos recurrentes, porque
  perdería la serie.
- **132 Diverso es de SOLO LECTURA**: `make new` da *"Insufficient permissions to
  create an object in target folder"* y `duplicate` da *"One or more objects
  could not be duplicated"*. No se puede mover nada hacia Diverso. Nunca.
- **Recorrer el calendario destino no escala.** La primera versión de
  `move_event` buscaba la copia iterando todos los eventos del destino: contra
  Gmail (2300 eventos) reventaba el timeout de 120 s sin completar un solo
  movimiento. Con un `whose` acotado (mismo asunto, inicio ±60 s) baja a ~1 s.
- **La igualdad de fechas dentro de un `whose` no casa nada.**
  `whose start time is <fecha>` devuelve 0 coincidencias incluso sobre el propio
  evento del que se sacó la fecha. Hay que usar un rango.
- **`st` no vale como nombre de variable** en AppleScript: da *"Se esperaba
  expresión pero se ha encontrado st"*. Usar `theStart`.
- **No se puede distinguir una persona de una sala.** Un `required attendee`
  creado a mano se lee después como `resource attendee`, igual que una sala de
  Teams. Por eso la guarda de `move_event` va sobre el RECUENTO de asistentes y
  no sobre su tipo.
- **Las categorías están vacías en la práctica.** 0 de 2330 en Gmail, 0 de 801 en
  131. El perfil tiene 49 categorías definidas y ningún evento las usa. No sirven
  como criterio de clasificación.
- **Outlook expone el mismo calendario de Google dos veces.** 149 "Google
  Principal" era copia de 205 en un 78,7% (444 de 625). Limpiar uno no limpiaba
  el otro. 149 ya fue eliminado por José.
- **El comando `sync` existe**: `sync <account>` y `sync <folder>`, según
  `Outlook.sdef`. Útil para empujar los cambios al servidor tras una tanda.
- **Barrer solo por palabras clave y solo en futuros deja mucho fuera.** Lo
  personal de 131 estaba casi todo en pasado y bajo términos que no se habían
  buscado: 75 eventos aparecieron en la segunda pasada, no en la primera.
