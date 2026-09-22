-- probe_move_event_mac_v3.applescript
--
-- Corrige dos errores de la v2:
--   1. `dupEvt` se declaraba dentro del bloque `tell` y AppleScript no lo
--      veía en la limpieza -> el script abortaba y se perdía TODA la
--      salida. Ahora las variables se declaran arriba del todo y la
--      limpieza va en su propio bloque, pase lo que pase.
--   2. El sondeo de campos de un evento se hacía sobre el propio evento
--      de prueba (sintético, sin organizador ni asistentes), así que sus
--      "NO EXPUESTO" no significaban nada. Ahora se mide contra eventos
--      REALES, saltando explícitamente los de prueba.
--
-- ATENCION: crea y borra un evento de prueba en Docencia. Nada más se
-- modifica.
--
-- Uso:  osascript probe_move_event_mac_v3.applescript

-- Declaradas ARRIBA: este era el bug.
global testEvt, dupEvt
set testEvt to missing value
set dupEvt to missing value

set LF to linefeed
set out to "=== sonda de movimiento v3 ===" & LF
set testSubject to "ZZ-MCP-PROBE-BORRAR"

set srcId to 133 -- Docencia
set dstId to 135 -- Tutorías

tell application "Microsoft Outlook"

	try
		set srcCal to calendar id srcId
		set dstCal to calendar id dstId
		set srcName to (name of srcCal) as text
		set dstName to (name of dstCal) as text
	on error errMsg
		return out & "ABORTADO: calendarios no resueltos: " & errMsg & LF
	end try

	set out to out & "origen:  " & srcName & "  destino: " & dstName & LF & LF

	set startD to (current date) + (1 * days)
	set time of startD to 3 * hours
	set endD to startD + (30 * minutes)

	------------------------------------------------------------------
	-- 1. Crear dentro de un calendario concreto (ya confirmado, se
	--    repite para tener el evento con el que probar el movimiento).
	------------------------------------------------------------------
	set out to out & "--- 1. CREAR EN CALENDARIO CONCRETO ---" & LF
	try
		set testEvt to make new calendar event at srcCal with properties {subject:testSubject, start time:startD, end time:endD}
		set out to out & "  creado en: " & ((name of (calendar of testEvt)) as text) & LF
	on error errMsg
		return out & "  ABORTADO: " & errMsg & LF
	end try

	------------------------------------------------------------------
	-- 2. LA PREGUNTA QUE DECIDE EL DISEÑO.
	------------------------------------------------------------------
	set out to out & LF & "--- 2. REASIGNAR CALENDARIO (calendarios reales) ---" & LF
	try
		set calendar of testEvt to dstCal
		set out to out & "  asignación sin error" & LF
		try
			set nowIn to (name of (calendar of testEvt)) as text
			set out to out & "  ahora está en: " & nowIn & LF
			if nowIn is dstName then
				set out to out & "  >>> REASIGNACION REAL. Mover es limpio." & LF
			else
				set out to out & "  >>> IGNORADA EN SILENCIO." & LF
			end if
		on error errMsg
			set out to out & "  no verificable: " & errMsg & LF
		end try
	on error errMsg
		set out to out & "  >>> NO SE PUEDE: " & errMsg & LF
	end try

	------------------------------------------------------------------
	-- 3. Plan B: duplicar a otro calendario.
	------------------------------------------------------------------
	set out to out & LF & "--- 3. DUPLICATE A OTRO CALENDARIO ---" & LF
	try
		set dupEvt to duplicate testEvt to dstCal
		set out to out & "  >>> FUNCIONA, copia en: " & ((name of (calendar of dupEvt)) as text) & LF
	on error errMsg
		set out to out & "  >>> NO: " & errMsg & LF
	end try

	------------------------------------------------------------------
	-- 4. Campos de EVENTOS REALES. Esto es lo que determina con qué
	--    criterios puedo clasificar automáticamente.
	--
	--    Se saltan los eventos de prueba. Se prueba `category` y
	--    `categories` por separado (probablemente sea una lista, y por
	--    eso `name of (category of e)` falló antes), y se informa de la
	--    CLASE de cada valor, que es lo realmente diagnóstico.
	------------------------------------------------------------------
	set out to out & LF & "--- 4. CAMPOS DE EVENTOS REALES ---" & LF

	-- Un calendario con reuniones de verdad (Docencia) y el cajón (131).
	repeat with probeCalId in {133, 131}
		try
			set c to calendar id probeCalId
			set out to out & LF & "CALENDARIO: " & ((name of c) as text) & LF
			set evs to calendar events of c
			set picked to missing value
			-- Cogemos el primer evento que NO sea de prueba.
			repeat with e in evs
				try
					if (subject of e) does not start with "ZZ-MCP-PROBE" then
						set picked to e
						exit repeat
					end if
				end try
			end repeat
			if picked is missing value then
				set out to out & "  sin eventos reales" & LF
			else
				set e to picked
				try
					set out to out & "  subject: " & (subject of e) & LF
				end try
				try
					set out to out & "  all day flag: " & ((all day flag of e) as text) & LF
				end try
				try
					set out to out & "  location: " & ((location of e) as text) & LF
				on error
					set out to out & "  location: NO EXPUESTO" & LF
				end try
				-- asistentes
				try
					set na to (count of (attendees of e))
					set out to out & "  n attendees: " & (na as text) & LF
					if na > 0 then
						set at1 to item 1 of (attendees of e)
						try
							set out to out & "    attendee 1 class: " & ((class of at1) as text) & LF
						end try
						try
							set out to out & "    attendee 1 address: " & ((address of (email address of at1)) as text) & LF
						on error errMsg
							set out to out & "    attendee address: NO (" & errMsg & ")" & LF
						end try
					end if
				on error errMsg
					set out to out & "  attendees: NO EXPUESTO (" & errMsg & ")" & LF
				end try
				-- organizador
				try
					set org to organizer of e
					set out to out & "  organizer class: " & ((class of org) as text) & LF
					try
						set out to out & "  organizer address: " & ((address of (email address of org)) as text) & LF
					end try
				on error errMsg
					set out to out & "  organizer: NO EXPUESTO (" & errMsg & ")" & LF
				end try
				-- categorias: probablemente una LISTA
				try
					set cats to category of e
					set out to out & "  category class: " & ((class of cats) as text) & LF
					set out to out & "  n categories: " & ((count of cats) as text) & LF
					repeat with ct in cats
						try
							set out to out & "    cat: " & ((name of ct) as text) & LF
						end try
					end repeat
				on error errMsg
					set out to out & "  category: NO (" & errMsg & ")" & LF
				end try
				try
					set cats2 to categories of e
					set out to out & "  categories (plural): " & ((count of cats2) as text) & LF
				on error
					set out to out & "  categories (plural): NO EXPUESTO" & LF
				end try
				-- recurrencia
				try
					set rc to recurrence of e
					set out to out & "  recurrence class: " & ((class of rc) as text) & LF
				on error errMsg
					set out to out & "  recurrence: NO (" & errMsg & ")" & LF
				end try
				-- ¿es una reunión o una cita?
				try
					set out to out & "  event class: " & ((class of e) as text) & LF
				end try
			end if
		on error errMsg
			set out to out & "  ERROR en calendario: " & errMsg & LF
		end try
	end repeat

end tell

------------------------------------------------------------------
-- LIMPIEZA: fuera del flujo principal, se ejecuta siempre.
------------------------------------------------------------------
set out to out & LF & "--- LIMPIEZA ---" & LF
tell application "Microsoft Outlook"
	if dupEvt is not missing value then
		try
			delete dupEvt
			set out to out & "  copia borrada" & LF
		on error errMsg
			set out to out & "  !!! copia NO borrada: " & errMsg & LF
		end try
	end if
	if testEvt is not missing value then
		try
			delete testEvt
			set out to out & "  evento de prueba borrado" & LF
		on error errMsg
			set out to out & "  !!! NO borrado: " & errMsg & LF
		end try
	end if
	-- Barrido de seguridad: cualquier resto de ejecuciones anteriores.
	try
		repeat with cid in {133, 135}
			set c to calendar id cid
			repeat with e in (calendar events of c)
				try
					if (subject of e) starts with "ZZ-MCP-PROBE" then
						delete e
						set out to out & "  resto borrado de " & ((name of c) as text) & LF
					end if
				end try
			end repeat
		end repeat
	on error errMsg
		set out to out & "  barrido: " & errMsg & LF
	end try
end tell

return out
