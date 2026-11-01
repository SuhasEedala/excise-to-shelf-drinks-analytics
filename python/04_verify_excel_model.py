"""Step 4 - independent check of the Excel model.

Re-computes every SKU's excise, price build-up and shelf price in plain Python (no Excel formulas)
and compares the result with the values Excel calculated. Open and save the workbook in Excel first
so the cached values are current.
Run: python 04_verify_excel_model.py ../excel/ProofPoint_Excise_to_Shelf_Simulator.xlsx
"""
import math, sys, warnings
from pathlib import Path
warnings.filterwarnings("ignore")
from openpyxl import load_workbook
path = sys.argv[1] if len(sys.argv) > 1 else Path(__file__).resolve().parent.parent / "excel" / "ProofPoint_Excise_to_Shelf_Simulator.xlsx"
wb=load_workbook(path,data_only=True)
S=wb["SKU Master"]; A=wb["Assumptions"]; E=wb["Excise Rates"]; P=wb["Price Engine"]
cats={A.cell(r,2).value:dict(dm=A.cell(r,3).value,rm=A.cell(r,4).value,cogs=A.cell(r,5).value,el=A.cell(r,6).value) for r in range(8,13)}
g={A.cell(r,2).value:A.cell(r,3).value for r in range(17,25)}
GST=g["GST rate"]; WET=g["Wine Equalisation Tax (WET)"]; FREE=g["Beer excise-free alcohol allowance"]; CDS=g["Container deposit cost per container"]; CPI=g["Half-yearly excise indexation (forecast)"]
# independent rate card: actual Aug-26 + indexation
base={"BEER_PKG_LOW":55.83,"BEER_PKG_MID":65.03,"BEER_PKG_FULL":65.03,"BEER_DRT_FULL":43.39,"SPIRITS":110.15,"BRANDY":102.87,"WET":0,"NONE":0}
rates={1:base}
rates[2]={k:(v if k=="BEER_DRT_FULL" else round(v*(1+CPI),2)) for k,v in base.items()}
rates[3]={k:round(rates[2][k]*(1+CPI),2) for k in base}
# compare rate card
errs=0
for i,code in enumerate(base):
    r=21+i
    for p in (1,2,3):
        x=E.cell(r,3+p).value
        if abs(x-rates[p][code])>1e-9: errs+=1; print("RATE MISMATCH",code,p,x,rates[p][code])
def code(base_,abv):
    if base_=="Beer": return "BEER_PKG_LOW" if abv<=.03 else ("BEER_PKG_MID" if abv<=.035 else "BEER_PKG_FULL")
    return {"Spirits":"SPIRITS","Brandy":"BRANDY","WET":"WET"}.get(base_,"NONE")
maxdiff=0
for i in range(24):
    r=7+i
    sku,name,brand,cat,base_,units,lpu,abv,pp,cdsn,cdsf,cliff,vol=[S.cell(r,c).value for c in range(2,15)]
    L=units*lpu; c=code(base_,abv)
    for p in (1,2,3):
        rate=rates[p][c]
        ex=L*max(abv-FREE,0)*rate if base_=="Beer" else L*abv*rate
        cds=cdsn*CDS if (cdsf=="Now" or p==3) else 0
        w=(pp+ex+cds)*(1+cats[cat]["dm"]); wet=WET*w if base_=="WET" else 0
        raw=(w+wet)*(1+cats[cat]["rm"])*(1+GST); shelf=math.ceil(raw)-0.01
        c0=6+(p-1)*9; pr=8+i
        xl_ex=P.cell(pr,c0+1).value; xl_raw=P.cell(pr,c0+6).value; xl_sh=P.cell(pr,c0+7).value
        d=max(abs(xl_ex-ex),abs(xl_raw-raw),abs(xl_sh-shelf)); maxdiff=max(maxdiff,d)
        if d>1e-6: errs+=1; print("MISMATCH",sku,p,xl_ex,ex,xl_raw,raw,xl_sh,shelf)
print("Price engine max abs diff:",maxdiff,"errors:",errs)
# Cliff scanner statuses
C=wb["Cliff Scanner"]; from collections import Counter
st=Counter(C.cell(13+i,10).value for i in range(24)); print("Cliff statuses",st)
print("Summary table",[ (C.cell(r,3).value,C.cell(r,4).value) for r in range(8,12)])
# Simulator KPIs current selection
M=wb["Simulator"]; print("Simulator sel:",M["C7"].value,M["C8"].value,"shelf",M["G7"].value,"gov share",M["G9"].value,"mult",M["G10"].value)
# multiplier check beer
print("Beer $1 excise multiplier =",(1.12)*(1.20)*1.1, " spirits:",1.18*1.25*1.1, " wine (WET):",1.2*1.29*1.32*1.1)
L=wb["ABV Lab"]; print("ABV verdict:",[(L.cell(r,2).value,L.cell(r,3).value) for r in range(14,18)])
X=wb["Scenario Planner"]; print("Scenario KPIs:",[(X.cell(r,10).value,X.cell(r,13).value) for r in range(7,11)])
# check errors in any cell
bad=[]
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for cl in row:
            if isinstance(cl.value,str) and cl.value in ("#REF!","#DIV/0!","#N/A","#VALUE!","#NAME?","#NUM!","#NULL!"): bad.append((ws.title,cl.coordinate,cl.value))
print("Excel error cells:",len(bad),bad[:10])
