-- probe_calendar_accounts_mac.applescript
-- SOLO LECTURA. No crea, modifica ni borra nada.
--
-- La sonda 1 listó 23 calendarios pero no dijo a qué cuenta pertenece cada
-- uno, y hay nombres repetidos ("Calendario" x3) y dos sin nombre. Esta
-- sonda busca la forma de atribuir cada calendario a su cuenta.
--
-- Uso:  osascript probe_calendar_accounts_mac.applescript

set LF to linefeed
set out to "=== sonda de atribución calendario -> cuenta ===" & LF

tell application "Microsoft Outlook"

	------------------------------------------------------------------
	-- Vía A: ¿los calendarios cuelgan de cada cuenta?
	------------------------------------------------------------------
	set out to out & LF & "--- VIA A: calendars of <account> ---" & LF
	try
		repeat with a in exchange accounts
			set aName to (name of a)
			set out to out & "CUENTA: " & aName & LF
			try
				set acals to calendars of a
				set out to out & "  calendars of account: " & (count of acals) & LF
				repeat with c in acals
					set cid to "?"
					set cname to "?"
					try
						set cid to (id of c as text)
					end try
					try
						set cname to (name of c) as text
					end try
					set out to out & "    id=" & cid & " | " & cname & LF
				end repeat
			on error errMsg
				set out to out & "  NO SOPORTADO: " & errMsg & LF
			end try
		end repeat
	on error errMsg
		set out to out & "ERROR: " & errMsg & LF
	end try

	------------------------------------------------------------------
	-- Vía B: ¿cada calendario tiene una propiedad que apunte a su cuenta?
	-- Volcamos las propiedades completas de unos cuantos calendarios
	-- representativos para ver qué campos existen de verdad.
	------------------------------------------------------------------
	set out to out & LF & "--- VIA B: properties of <calendar> ---" & LF
	set probeIds to {131, 133, 135, 146, 205, 102}
	repeat with pid in probeIds
		try
			set c to calendar id pid
			set out to out & LF & "CAL id=" & (pid as text) & LF
			try
				set out to out & "  properties: " & ((properties of c) as text) & LF
			on error
				-- Los records no siempre se pueden convertir a texto; probamos
				-- campo a campo los candidatos más probables.
				try
					set out to out & "  name: " & ((name of c) as text) & LF
				end try
				try
					set out to out & "  account: " & ((name of (account of c)) as text) & LF
				on error errMsg
					set out to out & "  account: NO EXPUESTO (" & errMsg & ")" & LF
				end try
				try
					set out to out & "  folder: " & ((name of (folder of c)) as text) & LF
				on error errMsg
					set out to out & "  folder: NO EXPUESTO" & LF
				end try
			end try
		on error errMsg
			set out to out & "CAL id=" & (pid as text) & " ERROR: " & errMsg & LF
		end try
	end repeat

	------------------------------------------------------------------
	-- Vía C: propiedades de un evento. Necesito saber qué campos puedo
	-- leer para clasificar (organizador, asistentes, recurrencia,
	-- categorías) y si el evento apunta a su calendario.
	------------------------------------------------------------------
	set out to out & LF & "--- VIA C: properties of <calendar event> ---" & LF
	try
		set e to item 1 of (calendar events of (calendar id 133))
		set out to out & "evento de muestra de Docencia:" & LF
		try
			set out to out & "  subject: " & (subject of e) & LF
		end try
		try
			set out to out & "  calendar: " & (name of (calendar of e)) & LF
		end try
		try
			set out to out & "  properties: " & ((properties of e) as text) & LF
		on error
			set out to out & "  (properties no convertible a texto; campos sueltos:)" & LF
			try
				set out to out & "  organizer: " & ((name of (organizer of e)) as text) & LF
			on error
				set out to out & "  organizer: NO EXPUESTO" & LF
			end try
			try
				set out to out & "  n attendees: " & ((count of (attendees of e)) as text) & LF
			on error
				set out to out & "  attendees: NO EXPUESTO" & LF
			end try
			try
				set out to out & "  recurrence: " & ((recurrence of e) as text) & LF
			on error
				set out to out & "  recurrence: NO EXPUESTO" & LF
			end try
			try
				set out to out & "  category: " & ((name of (category of e)) as text) & LF
			on error
				set out to out & "  category: NO EXPUESTO" & LF
			end try
			try
				set out to out & "  is all day: " & ((all day flag of e) as text) & LF
			on error
				set out to out & "  all day flag: NO EXPUESTO" & LF
			end try
		end try
	on error errMsg
		set out to out & "  ERROR: " & errMsg & LF
	end try

end tell

return out
