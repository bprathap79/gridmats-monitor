import json, sys, openpyxl
f = sys.argv[1]
wb = openpyxl.load_workbook(f, data_only=True)
d = {}
ws = wb["Solar Park 100MW"]
for r in range(14, 60):
    a = ws[f"A{r}"].value
    if a == "Cumulative (₹ lakh)":
        d["cum_lakh"] = [ws.cell(row=r, column=c).value for c in range(3, 29)]
    if a == "Extra energy from cleaner modules (MWh)":
        d["mwh_y1"] = ws.cell(row=r, column=4).value
    if a == "Value of extra energy (₹ lakh)":
        d["energy_y1_lakh"] = ws.cell(row=r, column=4).value
ws = wb["Sensitivity"]
d["tornado"] = [{"driver": ws[f"A{r}"].value, "lo": ws[f"B{r}"].value, "hi": ws[f"C{r}"].value,
                 "npv_lo": ws[f"D{r}"].value, "npv_hi": ws[f"E{r}"].value, "swing": ws[f"F{r}"].value} for r in range(7, 16)]
ws = wb["Scale-up"]
d["scale"] = {ws[f"A{r}"].value: [ws[f"{c}{r}"].value for c in "BCDE"] for r in range(5, 15)}
ws = wb["Pilot Offer"]
d["pilot_items"] = []
for r in range(5, 25):
    a = ws[f"A{r}"].value
    if a and ws[f"E{r}"].value is not None:
        d["pilot_items"].append({"item": a, "qty": ws[f"B{r}"].value, "unit": ws[f"C{r}"].value,
                                 "list": ws[f"E{r}"].value, "offer": ws[f"G{r}"].value})
json.dump(d, open(sys.argv[2], "w"), indent=1, default=str)
print(json.dumps(d, default=str)[:3000])
