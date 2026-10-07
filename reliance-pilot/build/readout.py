import sys, json, openpyxl
f = sys.argv[1]
wb = openpyxl.load_workbook(f, data_only=True)
out = {}
for name, dn in wb.defined_names.items():
    for sheet, ref in dn.destinations:
        v = wb[sheet][ref.replace("$", "")].value
        out[name] = v
json.dump(out, open(sys.argv[2], "w"), indent=1, default=str)
keys = sys.argv[3].split(",") if len(sys.argv) > 3 else sorted(out)
for k in keys:
    v = out.get(k)
    print(f"{k:22s} {v:,.4f}" if isinstance(v, (int, float)) else f"{k:22s} {v}")
