#!/usr/bin/env bash
# Rebuild the workbook, recalculate it, then build the deck from the model's results.
# Needs: python3 + openpyxl, LibreOffice (for recalculation), node + the packages in package.json,
# and PPTX_SKILL / RECALC pointing at the pptx skill folder and the xlsx recalc script.
set -euo pipefail
cd "$(dirname "$0")"
OUT=..
python3 build_model.py "$OUT/GridMats_Reliance_Solar_CostBenefit_Model.xlsx"
python3 "${RECALC:?set RECALC to the path of recalc.py from the xlsx skill}" "$OUT/GridMats_Reliance_Solar_CostBenefit_Model.xlsx" 90
python3 readout.py "$OUT/GridMats_Reliance_Solar_CostBenefit_Model.xlsx" outputs.json Park_NPV > /dev/null
python3 chartdata.py "$OUT/GridMats_Reliance_Solar_CostBenefit_Model.xlsx" chartdata.json > /dev/null
[ -d node_modules ] || npm install --silent
NODE_PATH="$PWD/node_modules" node build_deck.js "$OUT/GridMats_Reliance_Solar_Pilot_Proposal.pptx"
