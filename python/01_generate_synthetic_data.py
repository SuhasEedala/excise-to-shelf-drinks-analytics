"""Step 1 - generate the synthetic (but calibrated) weekly sales dataset for PROOF POINT.
Run: python 01_generate_synthetic_data.py  (takes a few minutes; seed is fixed so output is repeatable)
Fictional company: Southerly Drinks Co. (NSW brand owner + distributor).
Real inputs: ATO/ABF excise rates by effective date, WET 29%, beer 1.15% allowance, GST.
"""
import math, random
from datetime import date, timedelta
import numpy as np
import pandas as pd

rng = np.random.default_rng(42)
random.seed(42)

# ---------------- excise history (real; <=3% packaged back-cast via index factors) ----------------
EXC = pd.DataFrame([
    # eff date,  pkg_low, pkg_full, drt_full, spirits, brandy, factor, source
    (date(2024, 2, 5), 51.63, 60.12, 42.37, 101.85, 95.12, 1.018, "ABF ACN 2024/02"),
    (date(2024, 8, 5), 52.66, 61.32, 43.22, 103.89, 97.02, 1.020, "ABF ACN 2024/29"),
    (date(2025, 2, 3), 52.87, 61.57, 43.39, 104.31, 97.41, 1.004, "ABF ACN 2025/02"),
    (date(2025, 8, 4), 53.72, 62.56, 43.39, 105.98, 98.97, 1.016, "ABF ACN 2025/19 - draught frozen"),
    (date(2026, 2, 2), 54.74, 63.75, 43.39, 107.99, 100.85, 1.019, "ABF ACN 2026/04"),
    (date(2026, 8, 3), 55.83, 65.03, 43.39, 110.15, 102.87, 1.020, "ABF ACN 2026/30"),
], columns=["EffectiveDate", "BEER_PKG_LOW", "BEER_PKG_FULL", "BEER_DRT_FULL", "SPIRITS", "BRANDY", "IndexFactor", "Source"])
EXC["BEER_PKG_MID"] = EXC["BEER_PKG_FULL"]

def rates_on(d):
    row = EXC[EXC.EffectiveDate <= d].iloc[-1]
    return row

GST, WET, BEER_FREE, CDS = 0.10, 0.29, 0.0115, 0.12
CAT = {  # dm, rm, cogs, elasticity, promo k, trend/yr, rebate n/a
    "Beer":    dict(dm=.12, rm=.20, cogs=.55, el=-0.8, k=3.0, trend=-0.04),
    "RTD":     dict(dm=.15, rm=.25, cogs=.50, el=-1.1, k=3.5, trend=0.01),
    "Spirits": dict(dm=.18, rm=.25, cogs=.40, el=-0.9, k=2.5, trend=-0.02),
    "Wine":    dict(dm=.20, rm=.32, cogs=.60, el=-0.7, k=2.5, trend=-0.06),
    "No/Low":  dict(dm=.15, rm=.30, cogs=.55, el=-1.2, k=2.0, trend=0.30),
    "Draught": dict(dm=.10, rm=.00, cogs=.55, el=-0.5, k=0.0, trend=-0.03),
}
# id, name, brand, category, taxbase, units, L/unit, abv, pp(Aug-26), cds containers, cds from, price point, annual packs
SKUS = [
    ('B01', 'Southerly Lager 24x375ml Can', 'Southerly Brewing', 'Beer', 'Beer', 24, .375, .046, 17.40, 24, 'Now', 60, 180000),
    ('B02', 'Coastline Pale Ale 16x375ml Can', 'Southerly Brewing', 'Beer', 'Beer', 16, .375, .052, 12.40, 16, 'Now', 45, 60000),
    ('B03', 'Southerly Mid 24x375ml Can', 'Southerly Brewing', 'Beer', 'Beer', 24, .375, .035, 15.20, 24, 'Now', 50, 90000),
    ('B04', 'Harbour Hazy IPA 4x440ml Can', 'Southerly Brewing', 'Beer', 'Beer', 4, .44, .065, 6.40, 4, 'Now', 25, 40000),
    ('B05', 'Southerly Light 24x375ml Bottle', 'Southerly Brewing', 'Beer', 'Beer', 24, .375, .027, 14.80, 24, 'Now', 45, 30000),
    ('B06', 'Old Quay Lager 30x375ml Block', 'Southerly Brewing', 'Beer', 'Beer', 30, .375, .042, 20.00, 30, 'Now', 65, 70000),
    ('R01', 'Southerly Vodka Soda 10x330ml Can', 'Southerly Spirits', 'RTD', 'Spirits', 10, .33, .045, 10.60, 10, 'Now', 45, 120000),
    ('R02', 'Southerly Bourbon & Cola 10x375ml Can', 'Southerly Spirits', 'RTD', 'Spirits', 10, .375, .048, 10.50, 10, 'Now', 50, 80000),
    ('R03', 'Margarita Seltzer 4x355ml Can', 'Southerly Spirits', 'RTD', 'Spirits', 4, .355, .07, 4.06, 4, 'Now', 25, 50000),
    ('R04', 'Lime Hard Seltzer 10x330ml Can', 'Southerly Spirits', 'RTD', 'Spirits', 10, .33, .04, 9.40, 10, 'Now', 40, 60000),
    ('S01', 'Quayside Dry Gin 700ml', 'Southerly Spirits', 'Spirits', 'Spirits', 1, .7, .40, 9.30, 1, 'From 2027', 70, 25000),
    ('S02', 'Southerly Vodka 700ml', 'Southerly Spirits', 'Spirits', 'Spirits', 1, .7, .37, 5.20, 1, 'From 2027', 55, 40000),
    ('S03', 'Harbour Spiced Rum 700ml', 'Southerly Spirits', 'Spirits', 'Spirits', 1, .7, .35, 6.30, 1, 'From 2027', 55, 30000),
    ('S04', 'Southerly Single Malt 700ml', 'Southerly Spirits', 'Spirits', 'Spirits', 1, .7, .43, 22.00, 1, 'From 2027', 90, 6000),
    ('S05', 'Night Shift Coffee Liqueur 700ml', 'Southerly Spirits', 'Spirits', 'Spirits', 1, .7, .20, 7.20, 1, 'From 2027', 40, 12000),
    ('S06', 'Old Quay Brandy 700ml', 'Southerly Spirits', 'Spirits', 'Brandy', 1, .7, .37, 5.40, 1, 'From 2027', 50, 15000),
    ('W01', 'Southerly Estate Shiraz 750ml', 'Southerly Estate', 'Wine', 'WET', 1, .75, .14, 8.83, 1, 'From 2027', 20, 45000),
    ('W02', 'Southerly Estate Sauv Blanc 750ml', 'Southerly Estate', 'Wine', 'WET', 1, .75, .125, 6.10, 1, 'From 2027', 15, 60000),
    ('W03', 'Southerly Prosecco 750ml', 'Southerly Estate', 'Wine', 'WET', 1, .75, .11, 8.00, 1, 'From 2027', 20, 30000),
    ('W04', 'Southerly Cask Red 4L', 'Southerly Estate', 'Wine', 'WET', 1, 4, .13, 7.00, 0, 'From 2027', 15, 35000),
    ('W05', 'Southerly Reserve Cabernet 750ml', 'Southerly Estate', 'Wine', 'WET', 1, .75, .145, 16.50, 1, 'From 2027', 40, 8000),
    ('N01', 'Southerly Zero Lager 6x330ml Can', 'Southerly Brewing', 'No/Low', 'None', 6, .33, .004, 9.80, 6, 'Now', 20, 50000),
    ('N02', 'Botanic Zero Gin Alternative 700ml', 'Southerly Spirits', 'No/Low', 'None', 1, .7, 0, 19.50, 1, 'From 2027', 40, 10000),
    ('N03', 'Southerly Zero Sparkling 750ml', 'Southerly Estate', 'No/Low', 'None', 1, .75, 0, 6.20, 1, 'From 2027', 15, 20000),
    ('K01', 'Southerly Lager 50L Keg', 'Southerly Brewing', 'Draught', 'Draught', 1, 50, .046, 160.00, 0, 'n/a', 0, 9000),
    ('K02', 'Coastline Pale Ale 50L Keg', 'Southerly Brewing', 'Draught', 'Draught', 1, 50, .052, 175.00, 0, 'n/a', 0, 4500),
]
prod = pd.DataFrame(SKUS, columns=["SKU", "Product", "Brand", "Category", "TaxBase", "UnitsPerPack", "LitresPerUnit", "ABV",
                                   "ProducerPriceAug26", "DepositContainers", "DepositStarts", "PricePoint", "PlanAnnualPacks"])
prod["LitresPerPack"] = prod.UnitsPerPack * prod.LitresPerUnit
prod["StdDrinksPerPack"] = (prod.LitresPerPack * prod.ABV / 0.012674).round(1)
prod["PackType"] = np.where(prod.Category == "Draught", "Keg", np.where(prod.UnitsPerPack > 1, "Multipack", "Bottle/Cask"))
prod["CategoryOrder"] = prod.Category.map({"Beer": 1, "RTD": 2, "Spirits": 3, "Wine": 4, "No/Low": 5, "Draught": 6})

def tax_code(base, abv):
    if base == "Beer":
        return "BEER_PKG_LOW" if abv <= .03 else ("BEER_PKG_MID" if abv <= .035 else "BEER_PKG_FULL")
    return {"Spirits": "SPIRITS", "Brandy": "BRANDY", "Draught": "BEER_DRT_FULL"}.get(base, "NONE")
prod["TaxCode"] = [tax_code(b, a) for b, a in zip(prod.TaxBase, prod.ABV)]

def excise(p, d):
    r = rates_on(d)
    code = p.TaxCode
    if code == "NONE":
        return 0.0
    rate = r[code]
    if p.TaxBase in ("Beer", "Draught"):
        return p.LitresPerPack * max(p.ABV - BEER_FREE, 0) * rate
    return p.LitresPerPack * p.ABV * rate

def producer_price(p, d):
    # list price steps up 2.5% each 1 July; Aug-26 values are the FY27 level
    fy_steps = 0 if d >= date(2026, 7, 1) else (1 if d >= date(2025, 7, 1) else 2)
    return p.ProducerPriceAug26 / (1.025 ** fy_steps)

def shelf(p, d):
    c = CAT[p.Category]
    ex = excise(p, d)
    cds = p.DepositContainers * CDS if p.DepositStarts == "Now" else 0
    w = (producer_price(p, d) + ex + cds) * (1 + c["dm"])
    wet = WET * w if p.TaxBase == "WET" else 0
    return (w + wet) * (1 + c["rm"]) * (1 + GST)

# ---------------- customers (fictional outlets, real NSW suburbs) ----------------
SUB = {
    "Drummoyne": (-33.852, 151.154, "Inner West"), "Balmain": (-33.858, 151.179, "Inner West"), "Newtown": (-33.897, 151.179, "Inner West"),
    "Marrickville": (-33.911, 151.155, "Inner West"), "Leichhardt": (-33.883, 151.157, "Inner West"), "Stanmore": (-33.894, 151.164, "Inner West"),
    "Breakfast Point": (-33.843, 151.110, "Inner West"), "Burwood": (-33.877, 151.104, "Inner West"), "Strathfield": (-33.873, 151.094, "Inner West"),
    "Bondi": (-33.891, 151.274, "Eastern Suburbs"), "Randwick": (-33.914, 151.241, "Eastern Suburbs"), "Surry Hills": (-33.886, 151.211, "Eastern Suburbs"),
    "Sydney CBD": (-33.868, 151.209, "Sydney CBD"), "Manly": (-33.797, 151.285, "Northern Beaches"), "Dee Why": (-33.751, 151.286, "Northern Beaches"),
    "Chatswood": (-33.797, 151.183, "North Shore"), "North Sydney": (-33.839, 151.207, "North Shore"), "Hornsby": (-33.704, 151.099, "North Shore"),
    "Ryde": (-33.815, 151.104, "North Shore"), "Parramatta": (-33.815, 151.001, "Western Sydney"), "Blacktown": (-33.771, 150.906, "Western Sydney"),
    "Penrith": (-33.751, 150.694, "Western Sydney"), "Liverpool": (-33.920, 150.923, "Western Sydney"), "Bankstown": (-33.917, 151.034, "Western Sydney"),
    "Castle Hill": (-33.729, 151.004, "Hills District"), "Cronulla": (-34.058, 151.152, "Sutherland Shire"), "Sutherland": (-34.031, 151.058, "Sutherland Shire"),
    "Hurstville": (-33.967, 151.102, "St George"), "Newcastle": (-32.927, 151.776, "Hunter"), "Maitland": (-32.733, 151.557, "Hunter"),
    "Wollongong": (-34.425, 150.893, "Illawarra"), "Gosford": (-33.425, 151.342, "Central Coast"), "The Entrance": (-33.339, 151.498, "Central Coast"),
    "Orange": (-33.284, 149.100, "Central West"), "Bathurst": (-33.419, 149.577, "Central West"),
}
subs = list(SUB.keys())
cust = []
cid = 1
def add(name, banner, channel, group, suburb, size):
    global cid
    lat, lon, region = SUB[suburb]
    cust.append((f"C{cid:03d}", name, banner, channel, group, suburb, region, "NSW", lat + rng.normal(0, .003), lon + rng.normal(0, .003), size))
    cid += 1
for s in rng.choice(subs, 20, replace=False):
    add(f"Bottle Bay {s}", "Bottle Bay", "National Chain", "Off-premise", s, rng.choice(["Large", "Medium"], p=[.6, .4]))
for s in rng.choice(subs, 15, replace=False):
    add(f"CellarMax {s}", "CellarMax", "National Chain", "Off-premise", s, rng.choice(["Large", "Medium"], p=[.5, .5]))
ind_names = ["Harbourside Cellars", "The Corner Grog", "Paddle & Pour", "Local Liquor Co-op", "Wattle Wine Store", "Seaside Liquor",
             "Old Mill Liquor", "Gully Cellars", "Two Bays Liquor", "Fig Tree Drinks", "Lantern Lane Liquor", "Ridgeway Cellars",
             "Sandstone Liquor", "Bluegum Bottleshop", "Kingfisher Cellars", "Railway Parade Liquor", "Silver Gull Wines", "Hilltop Grog",
             "Banksia Bottles", "Ferry Wharf Cellars"]
for n in ind_names:
    add(f"{n}", "Independent", "Independent Bottle Shop", "Off-premise", rng.choice(subs), rng.choice(["Medium", "Small"], p=[.4, .6]))
pubs = ["The Southerly Arms", "Jacaranda Hotel", "The Tidewater Inn", "Ironbark Hotel", "The Lamplighter", "Wharfside Hotel",
        "The Copper Kettle", "Northshore Tavern", "The Pelican Hotel", "Stockman's Rest Hotel"]
for n in pubs:
    add(n, "Independent venue", "Pub & Hotel", "On-premise", rng.choice(subs), rng.choice(["Large", "Medium"], p=[.4, .6]))
for n in ["Harbour Leagues Club", "Bayview Bowls Club", "Westside RSL", "Parklands Sports Club", "Valley Workers Club"]:
    add(n, "Club", "Club", "On-premise", rng.choice(subs), rng.choice(["Large", "Medium"], p=[.5, .5]))
for n in ["Saltbush Dining", "Laneway Bar & Kitchen", "Olive & Ember", "The Rooftop Room", "Mezze House"]:
    add(n, "Independent venue", "Restaurant & Bar", "On-premise", rng.choice(subs), "Small")
for n, s in [("SipDirect Online", "Sydney CBD"), ("Doorstep Drinks Online", "Parramatta"), ("Cellar2Door Online", "North Sydney")]:
    add(n, n.replace(" Online", ""), "Online", "Online", s, "Large")
custdf = pd.DataFrame(cust, columns=["CustomerID", "Customer", "Banner", "Channel", "ChannelGroup", "Suburb", "Region", "State",
                                     "Latitude", "Longitude", "SizeTier"])
REBATE = {"National Chain": .08, "Independent Bottle Shop": .04, "Pub & Hotel": .06, "Club": .06, "Restaurant & Bar": .03, "Online": .05}
SIZE = {"Large": 1.6, "Medium": 1.0, "Small": 0.55}
CH_SHARE = {"National Chain": .45, "Independent Bottle Shop": .25, "Pub & Hotel": .10, "Club": .05, "Restaurant & Bar": .03, "Online": .12}

# ranging
def ranged(c, p):
    ch = c.Channel
    if p.Category == "Draught":
        return ch in ("Pub & Hotel", "Club") or (ch == "Restaurant & Bar" and p.SKU == "K02")
    if ch == "National Chain":
        return True
    if ch == "Independent Bottle Shop":
        return p.SKU not in ("B05", "S04", "W05", "N02") or rng.random() < .4
    if ch in ("Pub & Hotel", "Club"):
        return p.Category in ("RTD", "Spirits", "No/Low") or p.SKU in ("W01", "W02", "W03", "B04")
    if ch == "Restaurant & Bar":
        return p.Category in ("Wine", "Spirits") or p.SKU in ("N01", "N03", "B04")
    if ch == "Online":
        return p.SKU not in ("W04", "B05")
    return False

# ---------------- calendar ----------------
weeks = [date(2024, 7, 1) + timedelta(weeks=i) for i in range(0, 118)]
weeks = [w for w in weeks if w <= date(2026, 9, 28)]
SEAS = {1: 1.15, 2: .97, 3: 1.0, 4: .96, 5: .9, 6: .88, 7: .86, 8: .86, 9: .93, 10: 1.0, 11: 1.1, 12: 1.45}
SEAS_SPIRIT = {1: 1.0, 2: .92, 3: .95, 4: .98, 5: 1.0, 6: 1.05, 7: 1.08, 8: 1.05, 9: .97, 10: .98, 11: 1.05, 12: 1.35}

# promo calendar per SKU x promo group (chains together; independents separately)
promo = {}
MECH = ["Price off", "Multi-buy", "Catalogue feature"]
for p in prod.itertuples():
    if p.Category == "Draught":
        continue
    for grp in ("Chain", "Indep", "Online"):
        cal = {}
        i = int(rng.integers(0, 8))
        while i < len(weeks):
            dur = int(rng.integers(2, 5))
            depth = float(rng.choice([.10, .15, .20, .25], p=[.25, .35, .25, .15]))
            mech = str(rng.choice(MECH))
            for j in range(i, min(i + dur, len(weeks))):
                cal[j] = (depth, mech)
            if i + dur < len(weeks):
                cal.setdefault(i + dur, (0.0, "Post-promo dip"))
            i += dur + int(rng.integers(7, 13))
        promo[(p.SKU, grp)] = cal

# per-customer weights so that SKU annual volume ~ plan split by channel share
rows = []
base_shelf = {p.SKU: shelf(p, weeks[0]) for p in prod.itertuples()}
for p in prod.itertuples():
    c = CAT[p.Category]
    elig = [cc for cc in custdf.itertuples() if ranged(cc, p)]
    if not elig:
        continue
    by_ch = {}
    for cc in elig:
        by_ch.setdefault(cc.Channel, []).append(cc)
    shares = {ch: CH_SHARE[ch] for ch in by_ch} if p.Category != "Draught" else {ch: {"Pub & Hotel": .7, "Club": .25, "Restaurant & Bar": .05}[ch] for ch in by_ch}
    tot = sum(shares.values())
    for ch, lst in by_ch.items():
        ch_annual = p.PlanAnnualPacks * shares[ch] / tot
        wsum = sum(SIZE[x.SizeTier] for x in lst)
        for cc in lst:
            base_week = ch_annual * SIZE[cc.SizeTier] / wsum / 52 * 1.06  # FY25 level a bit above plan (decline since)
            grp = "Chain" if ch == "National Chain" else ("Indep" if ch == "Independent Bottle Shop" else ("Online" if ch == "Online" else None))
            for wi, w in enumerate(weeks):
                yrs = (w - weeks[0]).days / 365.25
                seas = (SEAS_SPIRIT if p.Category == "Spirits" else SEAS)[w.month]
                if p.Category == "No/Low":
                    seas = 0.6 * seas + 0.4 * (1.3 if w.month == 1 else 1.0)  # Dry-July/Jan effect light
                trend = (1 + c["trend"]) ** yrs
                sh = shelf(p, w) if p.Category != "Draught" else (producer_price(p, w) + excise(p, w))
                sh0 = base_shelf[p.SKU] if p.Category != "Draught" else (producer_price(p, weeks[0]) + excise(p, weeks[0]))
                price_eff = (sh / sh0) ** c["el"]
                cliff = 1.0
                if p.PricePoint and (math.ceil(sh) - 0.01) > p.PricePoint:
                    cliff = 0.94
                baseline = base_week * seas * trend * price_eff * cliff * float(rng.lognormal(0, 0.12))
                depth, mech, uplift = 0.0, "None", 1.0
                if grp:
                    pr = promo[(p.SKU, grp)].get(wi)
                    if pr:
                        depth, mech = pr
                        uplift = 0.9 if mech == "Post-promo dip" else 1 + depth * c["k"] * float(rng.uniform(.8, 1.2))
                packs = baseline * uplift
                if p.Category == "Draught":
                    packs = max(packs, 0)
                packs_i = int(round(packs)) if packs >= 1 else int(rng.random() < packs)
                base_i = baseline
                if packs_i == 0:
                    continue
                pp = producer_price(p, w)
                ex = excise(p, w)
                cds = p.DepositContainers * CDS if p.DepositStarts == "Now" else 0.0
                list_ex_wet = (pp + ex + cds) * (1 + c["dm"])
                wet = WET * list_ex_wet if p.TaxBase == "WET" else 0.0
                list_price = list_ex_wet + wet
                gross = packs_i * list_price
                promo_disc = gross * depth if mech not in ("None", "Post-promo dip") else 0.0
                rebate = (gross - promo_disc) * REBATE[ch]
                freight = 0.35 if p.UnitsPerPack > 1 else 0.20
                if p.Category == "Draught":
                    freight = 6.0
                cogs = packs_i * (pp * c["cogs"] + freight)
                rows.append((w, cc.CustomerID, p.SKU, packs_i, round(base_i, 2), round(list_price, 4), round(gross, 2),
                             round(promo_disc, 2), round(rebate, 2), round(packs_i * ex, 2), round(packs_i * cds, 2),
                             round(packs_i * wet, 2), round(cogs, 2), 1 if mech not in ("None", "Post-promo dip") else 0,
                             depth if mech not in ("None", "Post-promo dip") else 0.0, mech, round(sh, 2)))

fact = pd.DataFrame(rows, columns=["WeekStart", "CustomerID", "SKU", "Packs", "BaselinePacks", "ListPricePerPack", "GrossSales",
                                   "PromoDiscount", "Rebates", "Excise", "DepositLevy", "WET", "COGS", "PromoFlag", "PromoDepth",
                                   "PromoMechanic", "EstShelfPrice"])
fact["NetSales"] = (fact.GrossSales - fact.PromoDiscount - fact.Rebates).round(2)
fact["NetRevenueExTax"] = (fact.NetSales - fact.Excise - fact.DepositLevy - fact.WET).round(2)
fact["GrossProfit"] = (fact.NetRevenueExTax - fact.COGS).round(2)
fact = fact.merge(prod[["SKU", "LitresPerPack", "ABV"]], on="SKU")
fact["LitresAlcohol"] = (fact.Packs * fact.LitresPerPack * fact.ABV).round(3)
fact = fact.drop(columns=["LitresPerPack", "ABV"])

# ---------------- date table ----------------
d0, d1 = date(2024, 7, 1), date(2026, 10, 4)
dates = pd.DataFrame({"Date": pd.date_range(d0, d1, freq="D")})
dates["Year"] = dates.Date.dt.year
dates["MonthNo"] = dates.Date.dt.month
dates["Month"] = dates.Date.dt.strftime("%b")
dates["MonthYear"] = dates.Date.dt.strftime("%b %Y")
dates["MonthStart"] = dates.Date.values.astype("datetime64[M]")
dates["YearMonthKey"] = dates.Year * 100 + dates.MonthNo
dates["FY"] = np.where(dates.MonthNo >= 7, "FY" + ((dates.Year + 1) % 100).astype(str).str.zfill(2), "FY" + (dates.Year % 100).astype(str).str.zfill(2))
dates["FYNo"] = np.where(dates.MonthNo >= 7, dates.Year + 1, dates.Year)
dates["FYMonthNo"] = ((dates.MonthNo - 7) % 12) + 1
dates["FYQuarter"] = "Q" + (((dates.FYMonthNo - 1) // 3) + 1).astype(str)
dates["WeekCommencing"] = (dates.Date - pd.to_timedelta(dates.Date.dt.weekday, unit="D"))
def period_label(d):
    r = EXC[EXC.EffectiveDate <= d.date()].iloc[-1]
    return r.EffectiveDate.strftime("%b %Y") + " rates"
dates["ExcisePeriod"] = [period_label(d) for d in dates.Date]
dates["ExciseIncreaseDay"] = dates.Date.dt.date.isin(list(EXC.EffectiveDate)).astype(int)
dates["Season"] = dates.MonthNo.map({12: "Summer", 1: "Summer", 2: "Summer", 3: "Autumn", 4: "Autumn", 5: "Autumn",
                                     6: "Winter", 7: "Winter", 8: "Winter", 9: "Spring", 10: "Spring", 11: "Spring"})

# excise long table for charts
exc_long = EXC.melt(id_vars=["EffectiveDate", "IndexFactor", "Source"],
                    value_vars=["BEER_PKG_LOW", "BEER_PKG_FULL", "BEER_DRT_FULL", "SPIRITS", "BRANDY"],
                    var_name="TaxCode", value_name="RatePerLAL")
exc_long["TaxClass"] = exc_long.TaxCode.map({"BEER_PKG_LOW": "Packaged beer <=3%", "BEER_PKG_FULL": "Packaged beer >3%",
                                             "BEER_DRT_FULL": "Draught beer >3.5%", "SPIRITS": "Spirits & RTDs", "BRANDY": "Brandy"})
exc_long["Basis"] = np.where(exc_long.TaxCode == "BEER_PKG_LOW", "Back-cast from Aug-26 rate using index factors", "Published ABF/ATO rate")
exc_long["EffectiveDate"] = pd.to_datetime(exc_long.EffectiveDate)

prod_out = prod.drop(columns=["PlanAnnualPacks"])
fact["WeekStart"] = pd.to_datetime(fact.WeekStart)
print("fact rows", len(fact))
print(fact.groupby(fact.WeekStart.dt.year)[["Packs", "GrossSales", "NetSales", "GrossProfit", "Excise"]].sum().round(0))
fy = np.where(fact.WeekStart.dt.month >= 7, fact.WeekStart.dt.year + 1, fact.WeekStart.dt.year)
print(fact.groupby([fy, fact.merge(prod, on="SKU").Category.values])["Packs"].sum().unstack().round(0))

from pathlib import Path
out = Path(__file__).resolve().parent / "ProofPoint_PowerBI_Data.xlsx"
with pd.ExcelWriter(out, engine="openpyxl") as xw:
    fact.to_excel(xw, sheet_name="FactSales", index=False)
    prod_out.to_excel(xw, sheet_name="DimProduct", index=False)
    custdf.to_excel(xw, sheet_name="DimCustomer", index=False)
    dates.to_excel(xw, sheet_name="DimDate", index=False)
    exc_long.to_excel(xw, sheet_name="DimExciseRates", index=False)
print("saved", out)
