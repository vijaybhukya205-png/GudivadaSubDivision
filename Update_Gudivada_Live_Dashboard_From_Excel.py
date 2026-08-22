from pathlib import Path
import openpyxl, json, re

BASE=Path(__file__).resolve().parent
XLSX=BASE/"Gudivada_Net_Accounts_Latest.xlsx"
HTML=BASE/"index.html"

wb=openpyxl.load_workbook(XLSX,data_only=True)
def read_rows(sheet):
    ws=wb[sheet]
    h=[c.value for c in ws[1]]
    return [dict(zip(h,r)) for r in ws.iter_rows(min_row=2,values_only=True) if r[0]]

text=HTML.read_text(encoding="utf-8")
m=re.search(r'const D=(\{.*?\});',text,re.S)
if not m:
    raise RuntimeError("const D data block not found in index.html")
old=json.loads(m.group(1))
parent={x["Office Name"]:x.get("Parent SO","") for x in old.get("all",[])}

def enrich(items):
    out=[]
    for x in items:
        y=dict(x)
        y["Parent SO"]=parent.get(y["Office Name"],"")
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
    combined.append({"SO Name":s["Office Name"],"BOs":len(children),
        "Combined Target":target,"Proportionate Target":prop,"Net Accounts":net,
        "Achievement %":(net/prop*100 if prop else 0),"Yet to Achieve":max(prop-net,0)})

data={"bo":bo,"so_own":so,"so_comb":combined,"all":allr}
text=text[:m.start()]+("const D="+json.dumps(data,ensure_ascii=False,separators=(",",":"))+";")+text[m.end():]

try:
    for row in wb["Summary"].iter_rows(min_row=2,values_only=True):
        if row[0]=="Report Date":
            v=row[1]
            if hasattr(v,"strftime"): d=v.strftime("%d-%m-%Y")
            else:
                sv=str(v); d=f"{sv[8:10]}-{sv[5:7]}-{sv[:4]}"
            text=re.sub(r'As on Date:\s*\d{2}-\d{2}-\d{4}',f'As on Date: {d}',text)
            break
except Exception:
    pass

HTML.write_text(text,encoding="utf-8")
print("LIVE DASHBOARD DATA UPDATED FROM Gudivada_Net_Accounts_Latest.xlsx")
