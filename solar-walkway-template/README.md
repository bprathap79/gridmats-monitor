# GridMats® Solar Walkway proposal template

A reusable deck + cost-benefit workbook for GridMats Solar Walkway pilots. Each client gets one settings file; one command produces both files with matching numbers. **First client: Reliance Jio** (25 m² pilot, Reliance Corporate Park, Navi Mumbai).

## Files

| Path | What it is |
|---|---|
| `output/reliance-jio/GridMats_Solar_Walkway_Proposal_Reliance_Jio.pptx` | 18-slide pilot proposal for Reliance Jio Infocomm Ltd, with speaker notes |
| `output/reliance-jio/GridMats_Solar_Walkway_Model_Reliance_Jio.xlsx` | 14-sheet formula model behind every number in that deck |
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

Cover · **Client** (all client inputs) · **Product & Costs** (GridMats product, prices, loss stack) · **Module Tech** (PERC / TOPCon / HJT library, 2026) · Irradiance (monthly sunshine, 9 cities) · Energy · Pilot BoQ (list vs GridMats contribution vs client pays) · Pilot Economics · Rollout Economics · Comparison (rooftop, carport, pavers + rooftop) · Sensitivity · Scale-up · Scorecard · Sources & Notes.

## Reliance Jio base case (ex-GST)

- Pilot: 25 m², 100 × 40 W TOPCon modules, 4.0 kWp, about 5,170 kWh a year (performance ratio 73% after seven losses).
- Reliance Jio pays ₹2.22 L (₹825/sq ft); GridMats contributes ₹3.69 L. Payback on Jio's outlay 4.6 years; on full cost 6.8 years.
- Campus rollout (1,000 m²): ₹1,081/sq ft installed, 5.3-year payback, 20% IRR, ₹7.22/kWh levelised against ₹10/kWh grid power.
- Rooftop solar is cheaper per kWp; the walkway case is where roofs are full and a walkway is being built anyway.

All Solar Walkway yield, life and upkeep figures are estimates: the product is pre-commercial and the pilot measures them.
