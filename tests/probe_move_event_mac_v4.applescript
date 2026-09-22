-- probe_move_event_mac_v4.applescript
--
-- Las v2 y v3 murieron igual: una variable no visible en la limpieza, y
-- con el aborto se perdió toda la salida. Dos cambios de planteamiento:
--
--   1. La limpieza NO usa variables. Barre por asunto ("ZZ-MCP-PROBE")
--      todos los calendarios. Sin referencias que se puedan perder.
--   2. Todo el cuerpo va dentro de un try. Pase lo que pase, el script
--      llega al `return` y devuelve lo que haya averiguado hasta ahí.
--
-- Uso:  osascript probe_move_event_mac_v4.applescript

set LF to linefeed
set out to "=== sonda de movimiento v4 ===" & LF
set testSubject to "ZZ-MCP-PROBE-BORRAR"

try
	tell application "Microsoft Outlook"

		set srcCal to calendar id 133 -- Docencia
		set dstCal to calendar id 135 -- Tutorías
		set srcName to (name of srcCal) as text
		set dstName to (name of dstCal) as text
		set out to out & "origen: " & srcName & " / destino: " & dstName & LF & LF

		set startD to (current date) + (1 * days)
		set time of startD to 3 * hours
		set endD to startD + (30 * minutes)

		--------------------------------------------------------------
		-- 1. Crear en un calendario concreto
		--------------------------------------------------------------
		set out to out & "--- 1. CREAR EN CALENDARIO CONCRETO ---" & LF
		set ev to make new calendar event at srcCal with properties {subject:testSubject, start time:startD, end time:endD}
		set out to out & "  creado en: " & ((name of (calendar of ev)) as text) & LF

		--------------------------------------------------------------
		-- 2. LA PREGUNTA QUE DECIDE EL DISEÑO
		--------------------------------------------------------------
		set out to out & LF & "--- 2. REASIGNAR CALENDARIO ---" & LF
		try
			set calendar of ev to dstCal
			set nowIn to (name of (calendar of ev)) as text
			set out to out & "  sin error; ahora en: " & nowIn & LF
			if nowIn is dstName then
				set out to out & "  >>> REASIGNACION REAL. Mover es limpio." & LF
			else
				set out to out & "  >>> IGNORADA EN SILENCIO." & LF
			end if
		on error errMsg
			set out to out & "  >>> NO SE PUEDE: " & errMsg & LF
		end try

		--------------------------------------------------------------
		-- 3. Plan B: duplicar
		--------------------------------------------------------------
		set out to out & LF & "--- 3. DUPLICATE ---" & LF
		try
			set cp to duplicate ev to dstCal
			set out to out & "  >>> FUNCIONA, copia en: " & ((name of (calendar of cp)) as text) & LF
		on error errMsg
			set out to out & "  >>> NO: " & errMsg & LF
		end try

		--------------------------------------------------------------
		-- 4. Campos de EVENTOS REALES (saltando los de prueba)
		--------------------------------------------------------------
		set out to out & LF & "--- 4. CAMPOS DE EVENTOS REALES ---" & LF

		repeat with cid in {133, 131}
			try
				set c to calendar id cid
				set out to out & LF & "CALENDARIO: " & ((name of c) as text) & LF
				set picked to missing value
				repeat with e in (calendar events of c)
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
					try
						set out to out & "  subject: " & (subject of picked) & LF
					end try
					try
						set out to out & "  class: " & ((class of picked) as text) & LF
					end try
					try
						set out to out & "  all day: " & ((all day flag of picked) as text) & LF
					end try
					try
						set out to out & "  location: " & ((location of picked) as text) & LF
					on error
						set out to out & "  location: NO" & LF
					end try

					try
						set na to (count of (attendees of picked))
						set out to out & "  n attendees: " & (na as text) & LF
						if na > 0 then
							set a1 to item 1 of (attendees of picked)
							try
								set out to out & "    class: " & ((class of a1) as text) & LF
							end try
							try
								set out to out & "    address: " & ((address of (email address of a1)) as text) & LF
							on error errMsg
								set out to out & "    address NO: " & errMsg & LF
							end try
						end if
					on error errMsg
						set out to out & "  attendees NO: " & errMsg & LF
					end try

					try
						set org to organizer of picked
						set out to out & "  organizer class: " & ((class of org) as text) & LF
						try
							set out to out & "  organizer addr: " & ((address of (email address of org)) as text) & LF
						end try
					on error errMsg
						set out to out & "  organizer NO: " & errMsg & LF
					end try

					try
						set cats to category of picked
						set out to out & "  category class: " & ((class of cats) as text) & LF
						try
							set out to out & "  n cats: " & ((count of cats) as text) & LF
							repeat with ct in cats
								try
									set out to out & "    cat: " & ((name of ct) as text) & LF
								end try
							end repeat
						end try
					on error errMsg
						set out to out & "  category NO: " & errMsg & LF
					end try

					try
						set out to out & "  recurrence class: " & ((class of (recurrence of picked)) as text) & LF
					on error errMsg
						set out to out & "  recurrence NO: " & errMsg & LF
					end try
				end if
			on error errMsg
				set out to out & "  ERROR calendario: " & errMsg & LF
			end try
		end repeat

	end tell
on error errMsg
	set out to out & LF & "!!! ERROR INESPERADO: " & errMsg & LF
end try

------------------------------------------------------------------
-- LIMPIEZA sin variables: barrido por asunto en TODOS los calendarios.
------------------------------------------------------------------
set out to out & LF & "--- LIMPIEZA (barrido por asunto) ---" & LF
try
	tell application "Microsoft Outlook"
		set nDeleted to 0
		repeat with c in calendars
			try
				repeat with e in (calendar events of c)
					try
						if (subject of e) starts with "ZZ-MCP-PROBE" then
							delete e
							set nDeleted to nDeleted + 1
						end if
					end try
				end repeat
			end try
		end repeat
		set out to out & "  eventos de prueba borrados: " & (nDeleted as text) & LF
	end tell
on error errMsg
	set out to out & "  barrido falló: " & errMsg & LF
	set out to out & "  busca \"ZZ-MCP-PROBE\" y borra a mano" & LF
end try

return out
