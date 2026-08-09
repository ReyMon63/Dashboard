#!/bin/bash
# Runner local del Visor Ejecutivo (Opcion 2) para la Mac.
# Copia los insumos vivos a inputs/ y ejecuta el motor.
#
# Uso:
#   ./run_local.sh              -> diario: proyecta el mes en curso y publica
#   ./run_local.sh --dry        -> arma index.html pero NO hace git push
#   ./run_local.sh --noproject  -> publica hasta el ultimo mes cerrado (sin proyectar)
#   ./run_local.sh --force      -> republica aunque ya se haya publicado hoy
# (los flags se pasan tal cual a engine/run.py)
set -e
BASE="/Users/reymon/Desktop/Proyectos/visor_bot"
cd "$BASE"
AVANCE_SRC="/Users/reymon/Downloads/Avance de Ventas.xlsx"
PARAM_SRC="/Users/reymon/Library/CloudStorage/GoogleDrive-rrivas.work@gmail.com/Mi unidad/Visor Ejecutivo/Parámetros Base.xlsx"
SEG_SRC="$BASE/assets/Segmentacion.xlsx"
# Segmentación VIVA (tu archivo de Drive sincronizado); si existe, se usa esa en vez de la copia.
# Es solo lectura: el runner la copia HACIA inputs/, nunca modifica tu archivo.
SEG_LIVE="/Users/reymon/Library/CloudStorage/GoogleDrive-rrivas.work@gmail.com/Mi unidad/Avance Ventas Bot/Segmentación.xlsx"
[ -f "$SEG_LIVE" ] && SEG_SRC="$SEG_LIVE"

mkdir -p inputs logs
[ -f "$AVANCE_SRC" ] || { echo "FALTA Avance: $AVANCE_SRC"; exit 3; }
[ -f "$PARAM_SRC" ]  || { echo "FALTA Parametros: $PARAM_SRC"; exit 3; }
[ -f "$SEG_SRC" ]    || { echo "FALTA Segmentacion: $SEG_SRC"; exit 3; }

cp "$AVANCE_SRC" inputs/Avance_Ventas.xlsx
cp "$PARAM_SRC"  inputs/Parametros.xlsx
cp "$SEG_SRC"    inputs/Segmentacion.xlsx

# Cierre de mes: si el bot dejó 'Cierre de Ventas.xlsx', se revisa en esta corrida.
CIERRE_SRC="/Users/reymon/Downloads/Cierre de Ventas.xlsx"
if [ -f "$CIERRE_SRC" ]; then
  cp "$CIERRE_SRC" inputs/Cierre_Ventas.xlsx
  export CIERRE_SRC
  echo ">> Cierre de Ventas detectado: se revisará el cierre de mes."
else
  rm -f inputs/Cierre_Ventas.xlsx
fi

export VISOR_GH_TOKEN="$(cat "$BASE/.gh_token")"
# Carpeta de los Históricos "(con estrellas)" en tu Drive (para actualizarlos al cerrar mes)
export HIST_DIR="/Users/reymon/Library/CloudStorage/GoogleDrive-rrivas.work@gmail.com/Mi unidad/Visor Ejecutivo"
echo "== Visor diario $(date '+%Y-%m-%d %H:%M') =="
python3 engine/run.py "$@"
