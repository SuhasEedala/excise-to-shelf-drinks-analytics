"""Step 2 - turn the weekly synthetic dataset into the monthly CSVs used by Power BI and Tableau.

Input : ProofPoint_PowerBI_Data.xlsx (written by 01_generate_synthetic_data.py)
Output: data/powerbi/*.csv and data/tableau/*.csv
"""
from pathlib import Path
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent
OUT_PBI = ROOT.parent / "data" / "powerbi"
OUT_TAB = ROOT.parent / "data" / "tableau"
OUT_PBI.mkdir(parents=True, exist_ok=True)
OUT_TAB.mkdir(parents=True, exist_ok=True)
LAST_DAY = "2026-09-30"  # last full month of sales

x = pd.read_excel(ROOT / "ProofPoint_PowerBI_Data.xlsx", sheet_name=None)

# ---- FactSales: weekly -> monthly, keeping promo detail needed for promo ROI ----
f = x["FactSales"]
f["MonthStart"] = f.WeekStart.values.astype("datetime64[M]")
promo = f.PromoFlag == 1
full_gp_pp = (f.GrossProfit + f.PromoDiscount) / f.Packs
f["PromoPacks"] = np.where(promo, f.Packs, 0)
f["PromoBaselinePacks"] = np.where(promo, f.BaselinePacks, 0)
f["PromoGrossSales"] = np.where(promo, f.GrossSales, 0)
f["PromoIncrementalGP"] = np.where(promo, f.GrossProfit - f.BaselinePacks * full_gp_pp, 0)
f["PromoWeeks"] = promo.astype(int)
f["ShelfXPacks"] = f.EstShelfPrice * f.Packs
sums = ["Packs", "BaselinePacks", "GrossSales", "PromoDiscount", "Rebates", "NetSales", "Excise", "DepositLevy",
        "WET", "NetRevenueExTax", "COGS", "GrossProfit", "LitresAlcohol", "PromoPacks", "PromoBaselinePacks",
        "PromoGrossSales", "PromoIncrementalGP", "PromoWeeks", "ShelfXPacks"]
g = f.groupby(["MonthStart", "CustomerID", "SKU"], as_index=False)[sums].sum()
g["EstShelfPrice"] = (g.ShelfXPacks / g.Packs).round(2)
g = g.drop(columns=["ShelfXPacks"])
for c in g.columns:
    if g[c].dtype == float and c != "EstShelfPrice":
        g[c] = g[c].round(2)
g["MonthStart"] = g.MonthStart.dt.strftime("%Y-%m-%d")
g.to_csv(OUT_PBI / "FactSales.csv", index=False)

# ---- dimensions ----
prod, cust = x["DimProduct"], x["DimCustomer"]
prod.to_csv(OUT_PBI / "DimProduct.csv", index=False)
cust.to_csv(OUT_PBI / "DimCustomer.csv", index=False)

d = x["DimDate"]
for c in ("Date", "MonthStart", "WeekCommencing"):
    d[c] = pd.to_datetime(d[c]).dt.strftime("%Y-%m-%d")
d = d[d.Date <= LAST_DAY].copy()
d["FY"] = d.FY.replace({"FY27": "FY27 YTD"})  # FY27 is only Jul-Sep 2026
d.to_csv(OUT_PBI / "DimDate.csv", index=False)

# Real published rates only. Light beer (<=3%) rates used inside the synthetic sales are estimates,
# so they are not shown as "real" rates.
e = x["DimExciseRates"]
e["EffectiveDate"] = pd.to_datetime(e.EffectiveDate).dt.strftime("%Y-%m-%d")
e = e[e.TaxCode != "BEER_PKG_LOW"]
e.to_csv(OUT_PBI / "DimExciseRates.csv", index=False)
e.to_csv(OUT_TAB / "Excise_Rates_History.csv", index=False)

# ---- Tableau: one flat monthly table ----
t = (g.merge(prod[["SKU", "Product", "Category", "Brand", "PricePoint"]], on="SKU", how="left")
      .merge(cust[["CustomerID", "Customer", "Channel", "ChannelGroup", "Suburb", "Region", "Latitude", "Longitude"]],
             on="CustomerID", how="left"))
t["TaxesLevies"] = t.Excise + t.DepositLevy + t.WET
t.to_csv(OUT_TAB / "Southerly_Sales_Monthly.csv", index=False)
print("FactSales rows:", len(g))
