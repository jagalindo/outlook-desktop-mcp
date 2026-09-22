-- probe_move_event_mac.applescript
--
-- ATENCIÓN: esta sonda SÍ ESCRIBE. Crea un evento de prueba llamado
-- "ZZ-MCP-PROBE-BORRAR" mañana a las 03:00, intenta reasignarlo de
-- calendario, y lo borra al terminar. No toca ningún evento tuyo.
--
-- Ejecuta ANTES probe_calendars_mac.applescript. Si aquello no listó al
-- menos dos calendarios, esta sonda no tiene sentido.
--
-- Uso:  osascript probe_move_event_mac.applescript
-- Outlook tiene que estar abierto.

set LF to linefeed
set out to "=== sonda de movimiento entre calendarios ===" & LF
set testSubject to "ZZ-MCP-PROBE-BORRAR"

tell application "Microsoft Outlook"

	-- Necesitamos dos calendarios: origen y destino.
	try
		set cals to calendars
	on error errMsg
		return out & "ABORTADO: no hay colección `calendars`: " & errMsg & LF
	end try

	if (count of cals) < 2 then
		return out & "ABORTADO: hacen falta 2+ calendarios, hay " & (count of cals) & LF
	end if

	set srcCal to item 1 of cals
	set dstCal to item 2 of cals
	set out to out & "origen:  " & (name of srcCal) & LF
	set out to out & "destino: " & (name of dstCal) & LF & LF

	set startD to (current date) + (1 * days)
	set time of startD to 3 * hours
	set endD to startD + (30 * minutes)

	------------------------------------------------------------------
	-- PREGUNTA CLAVE nº3: ¿se puede crear un evento DIRECTAMENTE dentro
	-- de un calendario concreto? Si sí, ya vale para clasificar lo nuevo.
	------------------------------------------------------------------
	set out to out & "--- CREAR DENTRO DE UN CALENDARIO CONCRETO ---" & LF
	set testEvt to missing value
	try
		set testEvt to make new calendar event at srcCal with properties {subject:testSubject, start time:startD, end time:endD}
		set out to out & "  >>> FUNCIONA (make new ... at <calendar>)" & LF
	on error errMsg
		set out to out & "  >>> FALLA: " & errMsg & LF
		set out to out & "  reintentando sin especificar calendario..." & LF
		try
			set testEvt to make new calendar event with properties {subject:testSubject, start time:startD, end time:endD}
			set out to out & "  creado en el calendario por defecto" & LF
		on error errMsg2
			return out & "  ABORTADO: no se pudo crear el evento de prueba: " & errMsg2 & LF
		end try
	end try

	------------------------------------------------------------------
	-- PREGUNTA CLAVE nº4, la que decide todo: ¿se puede REASIGNAR el
	-- calendario de un evento que ya existe?
	--
	-- Si SÍ  -> mover es limpio, sin tocar invitaciones.
	-- Si NO  -> mover = crear en destino + borrar en origen, lo que en
	--           reuniones con asistentes manda cancelaciones y nuevas
	--           invitaciones a todo el mundo.
	------------------------------------------------------------------
	set out to out & LF & "--- REASIGNAR CALENDARIO DE UN EVENTO EXISTENTE ---" & LF
	try
		set calendar of testEvt to dstCal
		set out to out & "  set calendar of e to <cal>: SIN ERROR" & LF
		-- Verificamos que el cambio se haya aplicado de verdad: algunas
		-- propiedades aceptan la asignación y la ignoran en silencio.
		try
			set nowIn to name of (calendar of testEvt)
			set out to out & "  calendario tras la asignación: " & nowIn & LF
			if nowIn is (name of dstCal) then
				set out to out & "  >>> REASIGNACIÓN REAL. Mover es limpio." & LF
			else
				set out to out & "  >>> IGNORADA EN SILENCIO. No sirve." & LF
			end if
		on error errMsg
			set out to out & "  no se pudo verificar: " & errMsg & LF
		end try
	on error errMsg
		set out to out & "  >>> NO SE PUEDE: " & errMsg & LF
		set out to out & "  (mover tendría que ser crear+borrar)" & LF
	end try

	------------------------------------------------------------------
	-- Limpieza: borrar el evento de prueba.
	------------------------------------------------------------------
	set out to out & LF & "--- LIMPIEZA ---" & LF
	try
		delete testEvt
		set out to out & "  evento de prueba borrado" & LF
	on error errMsg
		set out to out & "  !!! NO SE PUDO BORRAR: " & errMsg & LF
		set out to out & "  !!! busca \"" & testSubject & "\" y bórralo a mano" & LF
	end try

end tell

return out
