# GridMats® Solar Walkway proposal template

A reusable deck + cost-benefit workbook for GridMats Solar Walkway pilots. Each client gets one settings file; one command produces both files with matching numbers. **First client: Reliance Jio** (25 m² pilot, Reliance Corporate Park, Navi Mumbai).

## Files

| Path | What it is |
|---|---|
| `output/reliance-jio/GridMats_Solar_Walkway_Proposal_Reliance_Jio.pptx` | 18-slide pilot proposal for Reliance Jio Infocomm Ltd, with speaker notes |
| `output/reliance-jio/GridMats_Solar_Walkway_Model_Reliance_Jio.xlsx` | 15-sheet formula model behind every number in that deck |
| `output/template/…_TEMPLATE.pptx` / `.xlsx` | The same pair with `[placeholders]`, for a quick look at a new client |
| `clients/reliance-jio.json` | Reliance Jio settings (site, tariff, pilot size, co-funding, "why" points, site requirements) |
| `clients/_new-client.json` | Blank settings to copy for the next client |
| `assets/` | Concept renders, logo, client logos and founder photo from the Client/Partner Deck v2 |
| `build/` | Generators: `build_sw_model.py` (workbook), `build_sw_deck.js` (deck), `build.sh` (both) |

## Making a proposal for a new client

1. Copy `clients/_new-client.json` to `clients/<client-id>.json` and fill it in (client and entity, site, city sunshine profile, landed ₹/kWh from the bill, pilot length × width, co-funding, why-points, site requirements).
2. Run `build/build.sh <client-id>` (needs Python + openpyxl, LibreOffice, Node; set `PPTX_SKILL` and `RECALC` as noted in the script).
3. Files land in `output/<client-id>/`.

No scripts? Open the Excel model and change the yellow cells on the **Client** sheet: every number recalculates. Edit the deck text by hand.

## Workbook sheets

Cover · **Client** (all client inputs, including site soil CBR) · **Product & Costs** (GridMats product, loss stack, ground-preparation rates) · **Cost Build-up** (**GridMats internal**: actual unit costs, margin, cost-plus price; remove before sending the workbook to a client) · **Module Tech** (PERC / TOPCon / HJT library, 2026) · Irradiance (monthly sunshine, 9 cities) · Energy · Pilot BoQ (list vs GridMats contribution vs client pays) · Pilot Economics · Rollout Economics · Comparison (rooftop, carport, pavers + rooftop) · Sensitivity · Scale-up · Scorecard · Sources & Notes.

## How the price is built (Cost Build-up)

Unit costs are per 40 W piece, four pieces per m² (GridMats actuals, 7 Oct 2026): module ₹1,200 · glass ₹1,600 · GridMats grid ₹125 · MLP lumber ₹150 = ₹3,075 a piece, **₹12,300/m²** (₹1,143/sq ft, ₹77/Wp) before margin, labour, wiring, inverter and management. Those five are **assumptions to confirm** (yellow cells): wiring ₹120/piece, inverter ₹5/Wp, labour ₹150/piece, management 8% at rollout, margin 25%, and a 0 / 10 / 20% volume cost-down at rollout (Conservative / Base / Optimistic). The client price is cost × (1 + management) ÷ (1 − margin); the GridMats grid is sold at its published ladder price.

Ground preparation follows the site's soaked CBR: compacted subgrade, geotextile, **100 mm WMM if CBR is under 8, otherwise 50 mm**, 20 mm M-sand bedding, edge restraint. Rates on 'Product & Costs' (geotextile ₹55/m² and WMM ₹2,000/m³ are assumptions).

## Reliance Jio base case (ex-GST)

- Pilot: 25 m², 100 × 40 W TOPCon modules, 4.0 kWp, about 5,170 kWh a year (performance ratio 73% after seven losses). Soil CBR assumed 5% (100 mm WMM): confirm at the site walk.
- Reliance Jio pays ₹3.26 L (₹1,211/sq ft); GridMats contributes ₹4.55 L (50% of the PV layer plus valued services). Payback on Jio's outlay 6.9 years; on full cost 10.7 years. GridMats' own net investment is about ₹3.3 L (Cost Build-up, section 5).
- Campus rollout (1,000 m²): ₹1,860/sq ft installed (₹20,020/m²), 9.6-year payback, 9% IRR, ₹12.46/kWh levelised against ₹10/kWh grid power. A 6-year payback needs about ₹1,160/sq ft: module + glass + lumber cost would have to fall from ₹2,950 to about ₹1,570 a piece (Cost Build-up, section 6).
- Rooftop solar is cheaper per kWp; the walkway case is where roofs are full and a walkway is being built anyway.

All Solar Walkway yield, life and upkeep figures are estimates: the product is pre-commercial and the pilot measures them.
