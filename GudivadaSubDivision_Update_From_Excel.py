from pathlib import Path
import openpyxl, json, re, html

BASE = Path(__file__).resolve().parent
XLSX = BASE / "Gudivada_Net_Accounts_Latest.xlsx"
INDEX = BASE / "index.html"
wb = openpyxl.load_workbook(XLSX, data_only=True)

def read_rows(sheet):
    ws = wb[sheet]
    h = [c.value for c in ws[1]]
    return [dict(zip(h, r)) for r in ws.iter_rows(min_row=2, values_only=True) if r[0]]

summary = {}
for row in wb["Summary"].iter_rows(min_row=2, values_only=True):
    if row[0] is not None:
        summary[str(row[0])] = row[1]

opened = int(summary.get("Accounts Opened") or 0)
closed = int(summary.get("Accounts Closed") or 0)
net_total = int(summary.get("Net Accounts") or 0)
if opened - closed != net_total:
    raise RuntimeError(f"Summary mismatch: Opened {opened} - Closed {closed} != Net {net_total}")

t = INDEX.read_text(encoding="utf-8")
m = re.search(r'const D=(\{.*?\});', t, re.S)
if not m:
    raise RuntimeError("const D data block not found")

old = json.loads(m.group(1))
parent = {x["Office Name"]: x.get("Parent SO", "") for x in old.get("all", [])}

def enrich(rows):
    return [dict(x, **{"Parent SO": parent.get(x["Office Name"], x.get("Parent SO", ""))}) for x in rows]

bo = enrich(read_rows("BO Performance"))
so = enrich(read_rows("SO Performance"))
allr = enrich(read_rows("All Offices"))

combined = []
for s in so:
    children = [b for b in bo if b.get("Parent SO") == s["Office Name"]]
    prop = (s.get("Proportionate Target") or 0) + sum(b.get("Proportionate Target") or 0 for b in children)
    net = (s.get("Net Accounts") or 0) + sum(b.get("Net Accounts") or 0 for b in children)
    target = (s.get("Annual Target") or 0) + sum(b.get("Annual Target") or 0 for b in children)
    combined.append({
        "SO Name": s["Office Name"], "BOs": len(children),
        "Combined Target": target, "Proportionate Target": prop,
        "Net Accounts": net, "Achievement %": (net / prop * 100 if prop else 0),
        "Yet to Achieve": max(prop - net, 0)
    })

D = {
    "bo": bo, "so_own": so, "so_comb": combined, "all": allr,
    "summary": {"Accounts Opened": opened, "Accounts Closed": closed, "Net Accounts": net_total}
}
t = t[:m.start()] + "const D=" + json.dumps(D, ensure_ascii=False, separators=(",", ":")) + ";" + t[m.end():]

date = summary.get("Report Date")
if hasattr(date, "strftime"):
    date = date.strftime("%d-%m-%Y")
elif date:
    date = str(date)
if date:
    t = re.sub(r'As on Date:\s*\d{2}-\d{2}-\d{4}', f'As on Date: {date}', t)
INDEX.write_text(t, encoding="utf-8")

CSS = """@page{size:A4 portrait;margin:10mm}
*{box-sizing:border-box}
body{font-family:Arial,Helvetica,sans-serif;margin:0;color:#17202a;font-size:10pt}
h1{text-align:center;font-size:18pt;margin:0 0 3mm;color:#173f5f}
.meta{text-align:center;font-size:10pt;font-weight:600;margin-bottom:5mm}
table{width:100%;border-collapse:collapse;table-layout:auto}
thead{display:table-header-group}
tr{page-break-inside:avoid}
th{background:#173f5f !important;color:#ffffff !important;padding:7px 5px;border:1px solid #173f5f;font-size:9pt;text-align:center;white-space:normal;-webkit-print-color-adjust:exact;print-color-adjust:exact}
td{padding:5px;border:1px solid #bfc5ca;font-size:9pt;vertical-align:top;white-space:normal;overflow-wrap:anywhere;-webkit-print-color-adjust:exact;print-color-adjust:exact}
th:first-child,td:first-child{width:8%;text-align:center;white-space:nowrap}
tr.g td{background:#c6efce !important}
tr.y td{background:#ffeb9c !important}
tr.o td{background:#fce4d6 !important}
tr.r td{background:#ffc7ce !important}
.actions{margin-bottom:8px}
.actions button{padding:6px 10px;margin-right:5px}
@media print{.actions{display:none}}"""

def report(title, rows, cols):
    rows = sorted(rows, key=lambda r: float(r.get("Achievement %", 0) or 0), reverse=True)
    th = "".join("<th>%s</th>" % html.escape(str(a)) for a, _ in cols)
    trs = []
    for i, r in enumerate(rows, 1):
        p = float(r.get("Achievement %", 0) or 0)
        cls = "g" if p >= 100 else "y" if p >= 80 else "o" if p >= 60 else "r"
        vals = []
        for a, k in cols:
            if k == "__rank__":
                v = i
            elif k == "__yet_open__":
                v = max(float(r.get("Proportionate Target", 0) or 0) - float(r.get("Net Accounts", 0) or 0), 0)
            else:
                v = r.get(k, "")
            if k == "Achievement %":
                v = f"{float(v or 0):.1f}%"
            elif isinstance(v, (int, float)):
                v = f"{v:,.0f}"
            vals.append("<td>%s</td>" % html.escape(str(v)))
        trs.append('<tr class="%s">%s</tr>' % (cls, "".join(vals)))
    return (
        '<!doctype html><html><head><meta charset="utf-8"><title>%s</title><style>%s</style></head>'
        '<body><div class="actions"><button onclick="window.print()">Print / Save PDF</button>'
        '<button onclick="window.close()">Close</button></div><h1>%s</h1>'
        '<div class="meta"><b>As on Date: %s</b></div>'
        '<table><thead><tr>%s</tr></thead><tbody>%s</tbody></table></body></html>'
        % (html.escape(title), CSS, html.escape(title), html.escape(date or "22-08-2026"), th, "".join(trs))
    )

common_cols = [
    ("Sl. No.", "__rank__"), ("Office Name", "Office Name"), ("Type", "Office Type"),
    ("Opened", "Accounts Opened"), ("Closed", "Accounts Closed"),
    ("Net Accounts", "Net Accounts"), ("Proportionate Target", "Proportionate Target"),
    ("Achievement", "Achievement %"), ("Yet to Open", "__yet_open__")
]
combined_cols = [
    ("Sl. No.", "__rank__"), ("SO Name", "SO Name"), ("BOs", "BOs"),
    ("Combined Target", "Combined Target"), ("Proportionate Target", "Proportionate Target"),
    ("Net Accounts", "Net Accounts"), ("Achievement", "Achievement %"), ("Yet to Open", "__yet_open__")
]

(BASE / "BO_Performance_Report.html").write_text(report("Gudivada Sub Division - BO Performance Report", bo, common_cols), encoding="utf-8")
(BASE / "SO_Performance_Report.html").write_text(report("Gudivada Sub Division - SO Performance Report", so, common_cols), encoding="utf-8")
(BASE / "SO_HO_Performance_Report.html").write_text(
    report("Gudivada Sub Division - SO + HO Performance Report", so + [r for r in allr if r.get("Office Type") == "HO"], common_cols),
    encoding="utf-8"
)
(BASE / "SO_Consolidated_Report.html").write_text(
    report("Gudivada Sub Division - SO-wise Consolidated Performance Report", combined, combined_cols),
    encoding="utf-8"
)

print("Dashboard and ALL standalone reports updated.")
print(f"Authoritative Summary: Opened={opened}, Closed={closed}, Net={net_total}")
print("Reports: BO + SO + SO/HO + SO-wise Consolidated")
