from pathlib import Path
import openpyxl, json, re, html

BASE=Path(__file__).resolve().parent
XLSX=BASE/"Gudivada_Net_Accounts_Latest.xlsx"
INDEX=BASE/"index.html"
wb=openpyxl.load_workbook(XLSX,data_only=True)

def read_rows(sheet):
    ws=wb[sheet]; h=[c.value for c in ws[1]]
    return [dict(zip(h,r)) for r in ws.iter_rows(min_row=2,values_only=True) if r[0]]

t=INDEX.read_text(encoding="utf-8")
m=re.search(r'const D=(\{.*?\});',t,re.S)
if not m: raise RuntimeError("const D data block not found")
old=json.loads(m.group(1))
parent={x["Office Name"]:x.get("Parent SO","") for x in old.get("all",[])}
def enrich(rows):
    return [dict(x,**{"Parent SO":parent.get(x["Office Name"],x.get("Parent SO",""))}) for x in rows]

bo=enrich(read_rows("BO Performance")); so=enrich(read_rows("SO Performance")); allr=enrich(read_rows("All Offices"))
combined=[]
for s in so:
    children=[b for b in bo if b.get("Parent SO")==s["Office Name"]]
    prop=(s.get("Proportionate Target") or 0)+sum(b.get("Proportionate Target") or 0 for b in children)
    net=(s.get("Net Accounts") or 0)+sum(b.get("Net Accounts") or 0 for b in children)
    target=(s.get("Annual Target") or 0)+sum(b.get("Annual Target") or 0 for b in children)
    combined.append({"SO Name":s["Office Name"],"BOs":len(children),"Combined Target":target,
        "Proportionate Target":prop,"Net Accounts":net,"Achievement %":(net/prop*100 if prop else 0),
        "Yet to Achieve":max(prop-net,0)})
D={"bo":bo,"so_own":so,"so_comb":combined,"all":allr}
t=t[:m.start()]+"const D="+json.dumps(D,ensure_ascii=False,separators=(",",":"))+";"+t[m.end():]
date=None
try:
    for row in wb["Summary"].iter_rows(min_row=2,values_only=True):
        if row[0]=="Report Date" and hasattr(row[1],"strftime"): date=row[1].strftime("%d-%m-%Y"); break
except Exception: pass
if date: t=re.sub(r'As on Date:\s*\d{2}-\d{2}-\d{4}',f'As on Date: {date}',t)
INDEX.write_text(t,encoding="utf-8")

def report(title,rows,cols):
    rows=sorted(rows,key=lambda r:float(r.get("Achievement %",0)),reverse=True)
    th="".join(f"<th>{a}</th>" for a,k in cols); trs=[]
    for i,r in enumerate(rows,1):
        p=float(r.get("Achievement %",0)); cls="g" if p>=100 else "y" if p>=80 else "o" if p>=60 else "r"
        vals=[]
        for a,k in cols:
            v=i if k=="__rank__" else max(float(r.get("Proportionate Target",0))-float(r.get("Net Accounts",0)),0) if k=="__yet_open__" else r.get(k,"")
            if k=="Achievement %": v=f"{float(v or 0):.1f}%"
            elif isinstance(v,(int,float)): v=f"{v:,.0f}"
            vals.append(f"<td>{html.escape(str(v))}</td>")
        trs.append(f'<tr class="{cls}">{"".join(vals)}</tr>')
    css="@page{size:A4 landscape;margin:8mm}body{font-family:Arial;margin:0;color:#17202a}h1{color:#173f5f}table{width:100%;border-collapse:collapse;font-size:8.5px}th{background:#173f5f;color:white;padding:5px}td{padding:4px;border:1px solid #ddd}tr.g td{background:#c6efce}tr.y td{background:#ffeb9c}tr.o td{background:#fce4d6}tr.r td{background:#ffc7ce}.actions{margin-bottom:10px}@media print{.actions{display:none}}"
    return f'<!doctype html><html><head><meta charset="utf-8"><style>{css}</style></head><body><div class="actions"><button onclick="window.print()">Print / Save PDF</button><button onclick="window.close()">Close</button></div><h1>{html.escape(title)}</h1><p><b>As on Date: {date or "22-08-2026"}</b> | Gudivada Sub Division</p><table><thead><tr>{th}</tr></thead><tbody>{"".join(trs)}</tbody></table></body></html>'

cols=[("Rank","__rank__"),("Office","Office Name"),("Type","Office Type"),("Opened","Accounts Opened"),("Closed","Accounts Closed"),("Net","Net Accounts"),("Proportionate Target","Proportionate Target"),("Achievement","Achievement %"),("Yet to Open","__yet_open__")]
(BASE/"SO_HO_Performance_Report.html").write_text(report("Gudivada Sub Division — SO + HO Performance",so+[r for r in allr if r.get("Office Type")=="HO"],cols),encoding="utf-8")
print("Dashboard and SO+HO report updated.")
