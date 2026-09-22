-- probe_calendars_mac.applescript
-- SOLO LECTURA. No crea, modifica ni borra nada.
--
-- Descubre qué expone el diccionario AppleScript de Outlook para Mac sobre
-- cuentas y calendarios, que es lo que outlook-desktop-mcp necesita para
-- poder trabajar con Docencia / Research / Tutorías.
--
-- Uso:  osascript probe_calendars_mac.applescript
-- Outlook tiene que estar abierto.

set LF to linefeed
set out to "=== sonda de calendarios: outlook-desktop-mcp ===" & LF

tell application "Microsoft Outlook"

	try
		set out to out & "Outlook version: " & (version as text) & LF
	on error errMsg
		set out to out & "Outlook version: ERROR " & errMsg & LF
	end try

	------------------------------------------------------------------
	-- Cuentas
	------------------------------------------------------------------
	set out to out & LF & "--- CUENTAS ---" & LF

	try
		set accs to exchange accounts
		set out to out & "exchange accounts: " & (count of accs) & LF
		repeat with a in accs
			set aName to "?"
			set aMail to "?"
			try
				set aName to (name of a)
			end try
			try
				set aMail to (email address of a)
			end try
			set out to out & "  exchange | " & aName & " | " & aMail & LF
		end repeat
	on error errMsg
		set out to out & "  exchange accounts ERROR: " & errMsg & LF
	end try

	try
		set accs to imap accounts
		repeat with a in accs
			set out to out & "  imap     | " & (name of a) & LF
		end repeat
	on error errMsg
		set out to out & "  imap accounts ERROR: " & errMsg & LF
	end try

	try
		set accs to pop accounts
		repeat with a in accs
			set out to out & "  pop      | " & (name of a) & LF
		end repeat
	on error errMsg
		set out to out & "  pop accounts ERROR: " & errMsg & LF
	end try

	------------------------------------------------------------------
	-- Calendarios
	--
	-- LA PREGUNTA CLAVE nº1: ¿existe una colección `calendars` y aparecen
	-- ahí Docencia, Research y Tutorías?
	------------------------------------------------------------------
	set out to out & LF & "--- CALENDARIOS ---" & LF

	try
		set cals to calendars
		set out to out & "count: " & (count of cals) & LF
		repeat with c in cals
			set cid to "?"
			set cname to "?"
			set nEv to "?"
			try
				set cid to (id of c as text)
			end try
			try
				set cname to (name of c)
			end try
			try
				set nEv to (count of (calendar events of c)) as text
			end try
			set out to out & "  CAL | id=" & cid & " | " & cname & " | eventos=" & nEv & LF
		end repeat
	on error errMsg
		set out to out & "  calendars ERROR: " & errMsg & LF
		set out to out & "  (si falla aquí, el diccionario no expone calendarios y" & LF
		set out to out & "   no hay nada que hacer por esta vía)" & LF
	end try

	------------------------------------------------------------------
	-- ¿Se puede LEER a qué calendario pertenece un evento?
	--
	-- LA PREGUNTA CLAVE nº2. Sin esto no puedo ni siquiera informar de
	-- dónde está cada evento.
	------------------------------------------------------------------
	set out to out & LF & "--- EVENTO -> CALENDARIO (lectura) ---" & LF

	try
		set evs to calendar events
		set out to out & "eventos en el calendario por defecto: " & (count of evs) & LF
		if (count of evs) is 0 then
			set out to out & "  sin eventos que sondear" & LF
		else
			set e to item 1 of evs
			try
				set out to out & "  id: " & (id of e as text) & LF
			end try
			try
				set out to out & "  subject: " & (subject of e) & LF
			end try
			try
				set out to out & "  calendar of event: " & (name of (calendar of e)) & LF
				set out to out & "  >>> LEGIBLE" & LF
			on error errMsg
				set out to out & "  >>> NO LEGIBLE: " & errMsg & LF
			end try
		end if
	on error errMsg
		set out to out & "  ERROR: " & errMsg & LF
	end try

end tell

return out
