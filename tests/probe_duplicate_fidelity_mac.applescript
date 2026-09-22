-- Mide si `duplicate` conserva los campos de un evento al copiarlo a otro
-- calendario. Es el paso que decide si move_event puede existir.
--
-- Seguridad: el unico asistente es malawito@gmail.com, la otra cuenta del
-- propio usuario. Nada sale hacia terceros.
--
-- Lecciones de v2/v3/v4 aplicadas:
--   * todo en UN solo bloque tell, cada paso en su propio try
--   * la limpieza borra por id explicito: delete (calendar event id N of calendar id M)
--   * los ids se van guardando en toDelete segun se crean

on describeEvent(e)
	set LF to linefeed
	set r to ""
	tell application "Microsoft Outlook"
		try
			set r to r & "      subject: " & (subject of e) & LF
		end try
		try
			set r to r & "      calendar: " & ((name of (calendar of e)) as text) & " (" & ((id of (calendar of e)) as text) & ")" & LF
		end try
		try
			set r to r & "      start: " & ((start time of e) as string) & LF
		end try
		try
			set r to r & "      location: " & ((location of e) as text) & LF
		end try
		try
			set r to r & "      allday: " & ((all day flag of e) as text) & LF
		end try
		try
			set r to r & "      body: [" & ((plain text content of e) as text) & "]" & LF
		on error
			set r to r & "      body: <ilegible>" & LF
		end try
		try
			set r to r & "      organizer: " & ((organizer of e) as text) & LF
		on error
			set r to r & "      organizer: <ilegible>" & LF
		end try
		try
			set na to count of (attendees of e)
			set r to r & "      n attendees: " & (na as text) & LF
			repeat with a in (attendees of e)
				set aTxt to "?"
				try
					set aTxt to (address of (email address of a)) as text
				on error
					try
						set aTxt to (name of (email address of a)) as text
					on error
						set aTxt to "<no legible>"
					end try
				end try
				set r to r & "         - " & ((class of a) as text) & " : " & aTxt & LF
			end repeat
		on error errMsg
			set r to r & "      attendees NO: " & errMsg & LF
		end try
		try
			set cats to category of e
			set r to r & "      n cats: " & ((count of cats) as text)
			repeat with ct in cats
				try
					set r to r & " [" & ((name of ct) as text) & "]"
				end try
			end repeat
			set r to r & LF
		on error errMsg
			set r to r & "      category NO: " & errMsg & LF
		end try
		try
			if (recurrence of e) is missing value then
				set r to r & "      recurrence: NO" & LF
			else
				set r to r & "      recurrence: SI" & LF
			end if
		on error errMsg
			set r to r & "      recurrence ilegible: " & errMsg & LF
		end try
	end tell
	return r
end describeEvent

set LF to linefeed
set out to "=== SONDA DE FIDELIDAD DE duplicate ===" & LF
set toDelete to {}

tell application "Microsoft Outlook"

	------------------------------------------------------------------
	-- TEST A: evento sintetico con asistente propio, categoria, sitio y cuerpo
	------------------------------------------------------------------
	set out to out & LF & "--- TEST A: evento con asistente (malawito@gmail.com) ---" & LF

	try
		set sd to (current date) + (3 * days)
		set time of sd to 6 * hours
		set ed to sd + (45 * minutes)

		set origEv to make new calendar event at calendar id 133 with properties {subject:"ZZ-MCP-FIDELITY-BORRAR", start time:sd, end time:ed, location:"Despacho de pruebas", content:"<html><body>linea 1<br>linea 2</body></html>"}
		set origId to id of origEv
		set toDelete to toDelete & {{133, origId}}
		set out to out & "  creado original id " & (origId as text) & LF

		-- asistente: la otra cuenta del propio usuario
		try
			make new required attendee at origEv with properties {email address:{address:"malawito@gmail.com"}}
			set out to out & "  asistente anadido" & LF
		on error errMsg
			set out to out & "  asistente NO se pudo anadir: " & errMsg & LF
		end try

		-- categoria: reutiliza una existente, no crea ninguna
		try
			if (count of categories) > 0 then
				set category of origEv to {item 1 of categories}
				set out to out & "  categoria puesta: " & ((name of (item 1 of categories)) as text) & LF
			else
				set out to out & "  no hay categorias en el perfil, se omite" & LF
			end if
		on error errMsg
			set out to out & "  categoria NO: " & errMsg & LF
		end try

		set out to out & LF & "   ORIGINAL:" & LF & my describeEvent(origEv)

		-- ids de destino ANTES de duplicar, para localizar la copia despues
		set beforeIds to {}
		try
			repeat with e in (calendar events of calendar id 135)
				set beforeIds to beforeIds & {id of e}
			end repeat
		end try

		try
			duplicate origEv to calendar id 135
			set out to out & "  duplicate ejecutado (no devuelve resultado)" & LF
		on error errMsg
			set out to out & "  duplicate FALLA: " & errMsg & LF
		end try

		-- localizar la copia por diferencia de ids
		set copyId to missing value
		try
			repeat with e in (calendar events of calendar id 135)
				set thisId to id of e
				if thisId is not in beforeIds then
					set copyId to thisId
					exit repeat
				end if
			end repeat
		end try

		if copyId is missing value then
			set out to out & "  >>> NO aparecio copia en 135" & LF
		else
			set toDelete to toDelete & {{135, copyId}}
			set out to out & LF & "   COPIA (id " & (copyId as text) & "):" & LF & my describeEvent(calendar event id copyId of calendar id 135)
		end if
	on error errMsg
		set out to out & "  !!! TEST A abortado: " & errMsg & LF
	end try

	------------------------------------------------------------------
	-- TEST B: evento REAL recurrente, 0 asistentes (no envia nada)
	------------------------------------------------------------------
	set out to out & LF & "--- TEST B: recurrencia sobre evento real 12716 (0 asistentes) ---" & LF
	try
		set out to out & "   ORIGINAL (NO se toca):" & LF & my describeEvent(calendar event id 12716 of calendar id 133)

		set beforeIds2 to {}
		repeat with e in (calendar events of calendar id 135)
			set beforeIds2 to beforeIds2 & {id of e}
		end repeat

		try
			duplicate (calendar event id 12716 of calendar id 133) to calendar id 135
			set out to out & "  duplicate ejecutado" & LF
		on error errMsg
			set out to out & "  duplicate FALLA: " & errMsg & LF
		end try

		set copyId2 to missing value
		repeat with e in (calendar events of calendar id 135)
			set thisId to id of e
			if thisId is not in beforeIds2 then
				set copyId2 to thisId
				exit repeat
			end if
		end repeat

		if copyId2 is missing value then
			set out to out & "  >>> NO aparecio copia en 135" & LF
		else
			set toDelete to toDelete & {{135, copyId2}}
			set out to out & LF & "   COPIA (id " & (copyId2 as text) & "):" & LF & my describeEvent(calendar event id copyId2 of calendar id 135)
		end if
	on error errMsg
		set out to out & "  !!! TEST B abortado: " & errMsg & LF
	end try

	------------------------------------------------------------------
	-- LIMPIEZA por id explicito. NUNCA toca el 12716.
	------------------------------------------------------------------
	set out to out & LF & "--- LIMPIEZA ---" & LF
	repeat with pair in toDelete
		set pcid to item 1 of pair
		set peid to item 2 of pair
		if peid is 12716 then
			set out to out & "  SALTADO el original real 12716" & LF
		else
			try
				delete (calendar event id peid of calendar id pcid)
				set out to out & "  borrado " & (peid as text) & " de cal " & (pcid as text) & LF
			on error errMsg
				set out to out & "  FALLO al borrar " & (peid as text) & ": " & errMsg & LF
			end try
		end if
	end repeat

end tell
return out
