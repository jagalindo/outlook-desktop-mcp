-- probe_move_event_mac_v2.applescript
--
-- Corrige la sonda anterior, que por error usó los calendarios 102 y 185
-- (ambos `missing value`, 0 eventos, probablemente nodos raíz). Contra
-- objetivos así los resultados no valen. Esta versión usa calendarios
-- reales, por id explícito.
--
-- ATENCIÓN: SÍ ESCRIBE. Crea un evento "ZZ-MCP-PROBE-BORRAR" mañana a las
-- 03:00 y lo borra al terminar. No toca ningún evento tuyo.
--
-- Uso:  osascript probe_move_event_mac_v2.applescript

set LF to linefeed
set out to "=== sonda de movimiento v2 (calendarios reales) ===" & LF
set testSubject to "ZZ-MCP-PROBE-BORRAR"

-- Docencia -> Tutorías: dos calendarios reales, misma cuenta de trabajo.
set srcId to 133
set dstId to 135

tell application "Microsoft Outlook"

	try
		set srcCal to calendar id srcId
		set dstCal to calendar id dstId
	on error errMsg
		return out & "ABORTADO: no se pudieron resolver los calendarios: " & errMsg & LF
	end try

	set out to out & "origen:  id=" & (srcId as text) & " " & ((name of srcCal) as text) & LF
	set out to out & "destino: id=" & (dstId as text) & " " & ((name of dstCal) as text) & LF & LF

	set startD to (current date) + (1 * days)
	set time of startD to 3 * hours
	set endD to startD + (30 * minutes)

	------------------------------------------------------------------
	-- 1. Crear directamente dentro de un calendario REAL, y verificar
	--    que el evento acaba de verdad ahí (no en el de por defecto).
	------------------------------------------------------------------
	set out to out & "--- CREAR DENTRO DE UN CALENDARIO REAL ---" & LF
	set testEvt to missing value
	try
		set testEvt to make new calendar event at srcCal with properties {subject:testSubject, start time:startD, end time:endD}
		set out to out & "  make new ... at <cal>: sin error" & LF
		try
			set landedIn to (name of (calendar of testEvt)) as text
			set out to out & "  aterrizó en: " & landedIn & LF
			if landedIn is ((name of srcCal) as text) then
				set out to out & "  >>> CREAR EN CALENDARIO CONCRETO: FUNCIONA" & LF
			else
				set out to out & "  >>> IGNORADO: fue al calendario por defecto" & LF
			end if
		on error errMsg
			set out to out & "  no verificable: " & errMsg & LF
		end try
	on error errMsg
		return out & "  ABORTADO: no se pudo crear: " & errMsg & LF
	end try

	------------------------------------------------------------------
	-- 2. LA PREGUNTA QUE DECIDE EL DISEÑO: reasignar el calendario de un
	--    evento existente, entre dos calendarios reales.
	------------------------------------------------------------------
	set out to out & LF & "--- REASIGNAR ENTRE CALENDARIOS REALES ---" & LF
	try
		set calendar of testEvt to dstCal
		set out to out & "  set calendar of e to <cal>: sin error" & LF
		try
			set nowIn to (name of (calendar of testEvt)) as text
			set out to out & "  calendario tras asignar: " & nowIn & LF
			if nowIn is ((name of dstCal) as text) then
				set out to out & "  >>> REASIGNACIÓN REAL. Mover es limpio." & LF
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
	-- 3. Plan B: ¿se puede duplicar un evento a otro calendario? Si
	--    `duplicate` funciona, copiar+borrar conserva mejor los campos
	--    que reconstruir el evento a mano propiedad a propiedad.
	------------------------------------------------------------------
	set out to out & LF & "--- PLAN B: duplicate ... to <calendar> ---" & LF
	set dupEvt to missing value
	try
		set dupEvt to duplicate testEvt to dstCal
		set out to out & "  >>> duplicate FUNCIONA" & LF
		try
			set out to out & "  copia en: " & ((name of (calendar of dupEvt)) as text) & LF
		end try
	on error errMsg
		set out to out & "  >>> duplicate NO: " & errMsg & LF
	end try

	------------------------------------------------------------------
	-- Limpieza
	------------------------------------------------------------------
	set out to out & LF & "--- LIMPIEZA ---" & LF
	if dupEvt is not missing value then
		try
			delete dupEvt
			set out to out & "  copia borrada" & LF
		on error errMsg
			set out to out & "  !!! copia NO borrada: " & errMsg & LF
		end try
	end if
	try
		delete testEvt
		set out to out & "  evento de prueba borrado" & LF
	on error errMsg
		set out to out & "  !!! NO BORRADO: " & errMsg & LF
		set out to out & "  !!! busca \"" & testSubject & "\" y bórralo a mano" & LF
	end try

end tell

return out
