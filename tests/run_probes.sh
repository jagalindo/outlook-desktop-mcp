#!/bin/bash
# Lanza todas las sondas pendientes de una vez y vuelca la salida a un
# archivo, para no tener que copiar y pegar nada.
#
# Uso:  bash tests/run_probes.sh
# Outlook tiene que estar abierto.

cd "$(dirname "$0")/.." || exit 1
OUT="tests/probe_output.txt"

{
  echo "########################################################"
  echo "# SONDA: movimiento entre calendarios reales (v2)"
  echo "# ATENCION: crea y borra un evento de prueba"
  echo "########################################################"
  osascript tests/probe_move_event_mac_v2.applescript 2>&1

  echo
  echo "########################################################"
  echo "# SONDA: atribucion calendario -> cuenta (solo lectura)"
  echo "########################################################"
  osascript tests/probe_calendar_accounts_mac.applescript 2>&1
} | tee "$OUT"

echo
echo "--------------------------------------------------------"
echo "Salida guardada en $OUT"
echo "No hace falta que copies nada: puedo leer ese archivo yo."
echo "--------------------------------------------------------"
