#!/usr/bin/env bash
# GridMats Solar Walkway proposal template: build the workbook and deck for one client.
#   ./build.sh reliance-jio        (reads ../clients/reliance-jio.json)
# Needs python3 + openpyxl, LibreOffice, node + package.json deps, and PPTX_SKILL / RECALC set
# (PPTX_SKILL = pptx skill folder; RECALC = recalc.py from the xlsx skill).
set -euo pipefail
cd "$(dirname "$0")"
ID="${1:?usage: build.sh <client-id>}"
CFG="../clients/$ID.json"
if [[ "$ID" == _* ]]; then OUTDIR="../output/template"; NAME="TEMPLATE"; else
  OUTDIR="../output/$ID"
  NAME=$(python3 -c "import json,sys,re;c=json.load(open('$CFG'));print(re.sub(r'[^A-Za-z0-9]+','_',c['client_short']).strip('_'))")
fi
mkdir -p "$OUTDIR"
XL="$OUTDIR/GridMats_Solar_Walkway_Model_${NAME}.xlsx"
PPT="$OUTDIR/GridMats_Solar_Walkway_Proposal_${NAME}.pptx"
python3 build_sw_model.py "$CFG" "$XL"
python3 "${RECALC:?set RECALC}" "$XL" 90
python3 readout.py "$XL" "$OUTDIR/outputs.json" PilotKWp > /dev/null
python3 sw_chartdata.py "$XL" "$OUTDIR/chartdata.json" > /dev/null
[ -d node_modules ] || npm install --silent
NODE_PATH="$PWD/node_modules" node build_sw_deck.js "$CFG" "$OUTDIR/outputs.json" "$OUTDIR/chartdata.json" "$PPT"
echo "Built: $XL"
echo "Built: $PPT"
