from pathlib import Path
import openpyxl, json, re

BASE = Path(__file__).resolve().parent
XLSX = BASE / "Gudivada_Net_Accounts_Latest.xlsx"

wb = openpyxl.load_workbook(XLSX, data_only=True)

def read_rows(sheet):
    ws=wb[sheet]
    h=[c.value for c in ws[1]]
    return [dict(zip(h,r)) for r in ws.iter_rows(min_row=2, values_only=True) if r[0]]

# Preserve BO -> parent SO mapping from the current dashboard.
index_path=BASE/"index.html"
index_text=index_path.read_text(encoding="utf-8")
m=re.search(r'const D=(\{.*?\});',index_text,re.S)
if not m:
    raise RuntimeError("const D data block not found in index.html")
old=json.loads(m.group(1))
parent={x["Office Name"]:x.get("Parent SO","") for x in old.get("all",[])}

def enrich(items):
    out=[]
    for x in items:
        y=dict(x)
        y["Parent SO"]=parent.get(y["Office Name"], y.get("Parent SO",""))
        out.append(y)
    return out

bo=enrich(read_rows("BO Performance"))
so=enrich(read_rows("SO Performance"))
allr=enrich(read_rows("All Offices"))

combined=[]
for s in so:
    children=[b for b in bo if b.get("Parent SO")==s["Office Name"]]
    target=(s.get("Annual Target") or 0)+sum(b.get("Annual Target") or 0 for b in children)
    prop=(s.get("Proportionate Target") or 0)+sum(b.get("Proportionate Target") or 0 for b in children)
    net=(s.get("Net Accounts") or 0)+sum(b.get("Net Accounts") or 0 for b in children)
    combined.append({
        "SO Name":s["Office Name"],"BOs":len(children),
        "Combined Target":target,"Proportionate Target":prop,
        "Net Accounts":net,"Achievement %":(net/prop*100 if prop else 0),
        "Yet to Achieve":max(prop-net,0)
    })

data={"bo":bo,"so_own":so,"so_comb":combined,"all":allr}
blob=json.dumps(data,ensure_ascii=False,separators=(",",":"))

# Update the dashboard data block.
new_index=index_text[:m.start()]+"const D="+blob+";"+index_text[m.end():]

# Update the displayed report date from Summary.
report_date=None
try:
    for row in wb["Summary"].iter_rows(min_row=2,values_only=True):
        if row[0]=="Report Date":
            v=row[1]
            if hasattr(v,"strftime"): report_date=v.strftime("%d-%m-%Y")
            break
except Exception:
    pass
if report_date:
    new_index=re.sub(r'As on Date:\s*\d{2}-\d{2}-\d{4}',f'As on Date: {report_date}',new_index)

index_path.write_text(new_index,encoding="utf-8")

# Update data/date in each standalone report by replacing a marked JSON block
# when present; otherwise the page is regenerated from the existing template
# by replacing its embedded D block if included.
for name in ["BO_Performance_Report.html","SO_Performance_Report.html","SO_Consolidated_Report.html"]:
    p=BASE/name
    if not p.exists():
        continue
    t=p.read_text(encoding="utf-8")
    t=re.sub(r'const D=(\{.*?\});', "const D="+blob+";", t, flags=re.S)
    if report_date:
        t=re.sub(r'As on Date:\s*\d{2}-\d{2}-\d{4}',f'As on Date: {report_date}',t)
    p.write_text(t,encoding="utf-8")

print("Dashboard and standalone BO/SO/Consolidated PDF report pages updated successfully.")
