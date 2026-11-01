"""PROOF POINT - Excise-to-Shelf Price Simulator (Excel engine builder)."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule, DataBarRule
from openpyxl.chart import BarChart, LineChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as CL

# ---------- palette ----------
BURG, GOLD, CREAM, INK, GREY = "5B1A2E", "D9A441", "FBF7F0", "2B2B2B", "8A8A8A"
LIGHT_BURG, LIGHT_GOLD, INPUT_FILL = "F3E6EA", "F8EDD6", "FFF2CC"
WHITE = "FFFFFF"
F = "Arial"

def font(sz=10, bold=False, color=INK, italic=False):
    return Font(name=F, size=sz, bold=bold, color=color, italic=italic)

def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)

thin = Side(style="thin", color="D9D2C5")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
BOTTOM = Border(bottom=Side(style="medium", color=GOLD))

FMT_USD = '$#,##0.00;($#,##0.00);"-"'
FMT_USD0 = '$#,##0;($#,##0);"-"'
FMT_PCT = '0.0%;(0.0%);"-"'
FMT_ABV = '0.0%'
FMT_NUM = '#,##0;(#,##0);"-"'
FMT_X = '0.00"x"'

wb = Workbook()

SHEETS = ["Start Here", "Simulator", "Cliff Scanner", "ABV Lab", "Scenario Planner",
          "Price Engine", "SKU Master", "Excise Rates", "Assumptions"]
ws_map = {}
for i, n in enumerate(SHEETS):
    ws_map[n] = wb.active if i == 0 else wb.create_sheet(n)
    ws_map[n].title = n
TAB = {"Start Here": BURG, "Simulator": GOLD, "Cliff Scanner": GOLD, "ABV Lab": GOLD,
       "Scenario Planner": GOLD, "Price Engine": GREY, "SKU Master": GREY,
       "Excise Rates": GREY, "Assumptions": GREY}


def base_sheet(ws, title, subtitle, widths):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = TAB[ws.title]
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.column_dimensions["A"].width = 2.5
    for r in range(1, 4):
        for c in range(1, 30):
            ws.cell(r, c).fill = fill(BURG)
    ws["B1"].value = "PROOF POINT"
    ws["B1"].font = font(9, True, GOLD)
    ws["B2"].value = title
    ws["B2"].font = font(18, True, WHITE)
    ws["B3"].value = subtitle
    ws["B3"].font = font(10, False, "E9DCCB", italic=True)
    ws.row_dimensions[2].height = 28
    # nav bar row 4
    for c in range(1, 30):
        ws.cell(4, c).fill = fill(LIGHT_GOLD)
    if ws.title != "Start Here":
        cell = ws.cell(4, 2)
        cell.value = '=HYPERLINK("#\'Start Here\'!A1","◄  Back to Start Here")'
        cell.font = font(9, True, BURG)
    else:
        ws.cell(4, 2).value = "Southerly Drinks Co. (fictional)  ·  Real ATO excise rates effective 3 Aug 2026"
        ws.cell(4, 2).font = font(9, True, BURG)
    ws.row_dimensions[4].height = 18


def _lbl_txt(color, sz=900):
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, RichTextProperties
    cp = CharacterProperties(sz=sz, b=True, solidFill=color)
    return RichText(bodyPr=RichTextProperties(), p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def hdr(ws, row, col, labels, color=BURG, fc=WHITE, wrap=True, height=30):
    for i, lab in enumerate(labels):
        c = ws.cell(row, col + i, lab)
        c.font = font(9, True, fc)
        c.fill = fill(color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
        c.border = BOX
    ws.row_dimensions[row].height = height


def section(ws, ref, text):
    ws[ref].value = text
    ws[ref].font = font(12, True, BURG)
    ws[ref].border = BOTTOM


def name(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


def inp(c, fmt=None):
    c.font = font(10, True, "0000FF")
    c.fill = fill(INPUT_FILL)
    c.border = BOX
    if fmt:
        c.number_format = fmt


def calc(c, fmt=None, bold=False, color=INK):
    c.font = font(10, bold, color)
    c.border = BOX
    if fmt:
        c.number_format = fmt


def link(c, fmt=None):
    c.font = font(10, False, "008000")
    c.border = BOX
    if fmt:
        c.number_format = fmt

# =====================================================================
# ASSUMPTIONS
# =====================================================================
A = ws_map["Assumptions"]
base_sheet(A, "Assumptions", "Every editable driver lives here. Blue = input you can change.",
           {"B": 34, "C": 16, "D": 16, "E": 16, "F": 16, "G": 60})
section(A, "B6", "Category economics (illustrative - replace with company data)")
hdr(A, 7, 2, ["Category", "Distributor margin", "Retailer margin", "COGS % of producer price",
              "Own-price elasticity", "Notes"])
cats = [("Beer", .12, .20, .55, -0.8), ("RTD", .15, .25, .50, -1.1), ("Spirits", .18, .25, .40, -0.9),
        ("Wine", .20, .32, .60, -0.7), ("No/Low", .15, .30, .55, -1.2)]
for i, (cname, dm, rm, cogs, el) in enumerate(cats):
    r = 8 + i
    calc(A.cell(r, 2, cname), bold=True)
    for j, v in enumerate([dm, rm, cogs]):
        inp(A.cell(r, 3 + j, v), FMT_PCT)
    inp(A.cell(r, 6, el), "0.00")
A.cell(8, 7, "Margins are mark-ups on cost (incl. tax), which is why excise gets 'margined' down the chain.").font = font(9, italic=True, color=GREY)
A.cell(9, 7, "Elasticities are illustrative planning values; RTD & No/Low assumed most price-sensitive.").font = font(9, italic=True, color=GREY)
A.cell(10, 7, "COGS % = producer cost of goods as a share of its ex-tax selling price.").font = font(9, italic=True, color=GREY)
name("CAT_LIST", "Assumptions!$B$8:$B$12")
name("CAT_DM", "Assumptions!$C$8:$C$12")
name("CAT_RM", "Assumptions!$D$8:$D$12")
name("CAT_COGS", "Assumptions!$E$8:$E$12")
name("CAT_EL", "Assumptions!$F$8:$F$12")

section(A, "B15", "Tax & regulatory drivers")
hdr(A, 16, 2, ["Driver", "Value", "Unit", "", "", "Source / note"])
glob = [
    ("GST rate", 0.10, FMT_PCT, "%", "GST_RATE", "ATO - GST applies to the final shelf price."),
    ("Wine Equalisation Tax (WET)", 0.29, FMT_PCT, "% of wholesale value", "WET_RATE", "ATO - WET is 29% of the last wholesale value of wine."),
    ("Beer excise-free alcohol allowance", 0.0115, '0.00%', "ABV points", "BEER_FREE", "ATO - beer excise is charged on (ABV - 1.15%)."),
    ("Container deposit cost per container", 0.12, FMT_USD, "$ / container", "CDS_COST", "10c refund + scheme admin (assumed). Wine & spirit bottles join NSW/SA/NT schemes mid-late 2027."),
    ("Half-yearly excise indexation (forecast)", 0.018, FMT_PCT, "% per half-year", "CPI_HALF", "Assumption. Recent ATO factors: 1.004, 1.016, 1.019, 1.020."),
    ("Price-point 'cliff' volume penalty", 0.06, FMT_PCT, "extra volume loss", "CLIFF_PEN", "Assumed extra loss when a pack crosses its defended price point."),
    ("ABV backlash coefficient", 0.015, FMT_PCT, "loss per (0.1pt cut)^2", "BACKLASH_K", "Calibrated loosely to VB's 2009 ABV cut (reversed 2012 after sales fell)."),
    ("Pure alcohol per standard drink", 0.012674, '0.000000', "litres", "STD_DRINK_L", "10 g ethanol / 0.789 g/ml = 12.67 ml (Australian standard drink)."),
]
for i, (lab, v, fmt, unit, nm, note) in enumerate(glob):
    r = 17 + i
    calc(A.cell(r, 2, lab), bold=True)
    inp(A.cell(r, 3, v), fmt)
    A.cell(r, 4, unit).font = font(9, color=GREY)
    A.cell(r, 7, note).font = font(9, italic=True, color=GREY)
    name(nm, f"Assumptions!$C${r}")

section(A, "B27", "Pricing periods")
hdr(A, 28, 2, ["Period", "Period #", "Excise status", "", "", "Note"])
periods = [("Aug 2026 (current)", "Actual ATO rates from 3 Aug 2026"),
           ("Feb 2027 (forecast)", "Indexed by forecast factor; draught beer still frozen"),
           ("Aug 2027 (forecast)", "Indexed again; draught freeze ends; wine/spirit bottles join container deposit schemes")]
for i, (p, note) in enumerate(periods):
    r = 29 + i
    calc(A.cell(r, 2, p), bold=True)
    calc(A.cell(r, 3, i + 1))
    A.cell(r, 4, "Actual" if i == 0 else "Forecast").font = font(9, color=GREY)
    A.cell(r, 7, note).font = font(9, italic=True, color=GREY)
name("PERIODS", "Assumptions!$B$29:$B$31")
A.freeze_panes = "A5"

# =====================================================================
# EXCISE RATES
# =====================================================================
E = ws_map["Excise Rates"]
base_sheet(E, "Excise Rates", "Real ATO / ABF rates (A$ per litre of pure alcohol) + forecast rate card",
           {"B": 24, "C": 34, "D": 16, "E": 16, "F": 16, "G": 14, "H": 44})
section(E, "B6", "Indexation history - actual rates")
hdr(E, 7, 2, ["Effective date", "Packaged beer >3.5%", "Draught beer >3.5%", "Spirits & RTDs", "Index factor", "", "Source"])
hist = [
    ("Aug 2023", 59.06, 41.62, 100.05, 1.022, "ABF Customs Notice 2023/33"),
    ("Feb 2024", 60.12, 42.37, 101.85, 1.018, "ABF Customs Notice 2024/02"),
    ("Aug 2024", 61.32, 43.22, 103.89, 1.020, "ABF Customs Notice 2024/29"),
    ("Feb 2025", 61.57, 43.39, 104.31, 1.004, "ABF Customs Notice 2025/02"),
    ("Aug 2025", 62.56, 43.39, 105.98, 1.016, "ABF Customs Notice 2025/19 (draught frozen)"),
    ("Feb 2026", 63.75, 43.39, 107.99, 1.019, "ABF Customs Notice 2026/04; ATO rates page"),
    ("Aug 2026", 65.03, 43.39, 110.15, 1.020, "ABF Customs Notice 2026/30; ATO rates page"),
]
for i, row in enumerate(hist):
    r = 8 + i
    calc(E.cell(r, 2, row[0]), bold=True)
    for j in range(3):
        inp(E.cell(r, 3 + j, row[1 + j]), FMT_USD)
        E.cell(r, 3 + j).fill = fill(WHITE)
    inp(E.cell(r, 6, row[4]), "0.000")
    E.cell(r, 6).fill = fill(WHITE)
    E.cell(r, 8, row[5]).font = font(9, italic=True, color=GREY)
r = 15
calc(E.cell(r, 2, "Change Aug-23 to Aug-26"), bold=True)
for j in range(3):
    calc(E.cell(r, 3 + j, f"={CL(3+j)}14/{CL(3+j)}8-1"), FMT_PCT, bold=True, color=BURG)
E.cell(16, 2, "Draught beer indexation is paused from Aug 2025 to Aug 2027 - packaged beer kept rising.").font = font(9, italic=True, color=GREY)

section(E, "B19", "Forecast rate card (drives the whole model)")
hdr(E, 20, 2, ["Tax code", "Description", "Aug 2026 (current)", "Feb 2027 (forecast)", "Aug 2027 (forecast)", "", "Basis"])
card = [
    ("BEER_PKG_LOW", "Packaged beer <=3% ABV", 55.83, False, "ATO item 1.1"),
    ("BEER_PKG_MID", "Packaged beer >3% to 3.5%", 65.03, False, "ATO item 1.5"),
    ("BEER_PKG_FULL", "Packaged beer >3.5%", 65.03, False, "ATO item 1.10"),
    ("BEER_DRT_FULL", "Draught beer >3.5%", 43.39, True, "ATO item 1.11 - frozen to Aug 2027"),
    ("SPIRITS", "Spirits, RTDs & other beverages", 110.15, False, "ATO items 2, 3.2, 3.10"),
    ("BRANDY", "Brandy", 102.87, False, "ATO item 3.1"),
    ("WET", "Wine (taxed via WET, not excise)", 0, False, "See WET_RATE"),
    ("NONE", "Non-alcoholic (<=1.15%) / untaxed", 0, False, "No excise"),
]
for i, (code, desc, v, frozen, basis) in enumerate(card):
    r = 21 + i
    calc(E.cell(r, 2, code), bold=True)
    calc(E.cell(r, 3, desc))
    inp(E.cell(r, 4, v), FMT_USD)
    if frozen:
        calc(E.cell(r, 5, f"=D{r}"), FMT_USD)
        calc(E.cell(r, 6, f"=ROUND(E{r}*(1+CPI_HALF),2)"), FMT_USD)
    else:
        calc(E.cell(r, 5, f"=ROUND(D{r}*(1+CPI_HALF),2)"), FMT_USD)
        calc(E.cell(r, 6, f"=ROUND(E{r}*(1+CPI_HALF),2)"), FMT_USD)
    E.cell(r, 8, basis).font = font(9, italic=True, color=GREY)
name("TAX_CODES", "'Excise Rates'!$B$21:$B$28")
name("RATE_CARD", "'Excise Rates'!$D$21:$F$28")

ch = LineChart()
ch.title = "Excise per litre of alcohol: packaged beer vs draught vs spirits"
ch.y_axis.title = "A$ per LAL"
ch.height, ch.width = 7.5, 17
data = Reference(E, min_col=3, max_col=5, min_row=7, max_row=14)
ch.add_data(data, titles_from_data=True)
ch.set_categories(Reference(E, min_col=2, min_row=8, max_row=14))
for s, col in zip(ch.series, [BURG, GOLD, "4F6D7A"]):
    s.graphicalProperties.line.solidFill = col
    s.graphicalProperties.line.width = 28000
    s.smooth = False
ch.y_axis.majorGridlines = None
ch.legend.position = "b"
E.add_chart(ch, "J6")
E.freeze_panes = "A5"

# =====================================================================
# SKU MASTER
# =====================================================================
S = ws_map["SKU Master"]
base_sheet(S, "SKU Master", "Southerly Drinks Co. - FICTIONAL portfolio of 24 SKUs (synthetic prices, real tax rules)",
           {"B": 8, "C": 38, "D": 20, "E": 11, "F": 10, "G": 9, "H": 10, "I": 9, "J": 13, "K": 11, "L": 11, "M": 12, "N": 13, "O": 10, "P": 11})
skus = [
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
]
NS = len(skus)
R0 = 7
R1 = R0 + NS - 1
hdr(S, 6, 2, ["SKU", "Product", "Brand", "Category", "Tax base", "Units / pack", "Litres / unit",
              "ABV", "Producer price ex-tax ($/pack)", "Deposit containers", "Deposit starts",
              "Price point to defend ($)", "Annual volume (packs)", "Litres / pack", "Std drinks / pack"], height=42)
for i, s in enumerate(skus):
    r = R0 + i
    for j, v in enumerate(s):
        c = S.cell(r, 2 + j, v)
        if j in (0, 1, 2, 3, 4, 10):
            calc(c)
        else:
            fmt = {5: FMT_NUM, 6: '0.000', 7: FMT_ABV, 8: FMT_USD, 9: FMT_NUM, 11: FMT_USD0, 12: FMT_NUM}[j]
            inp(c, fmt)
    calc(S.cell(r, 15, f"=G{r}*H{r}"), '0.00')
    calc(S.cell(r, 16, f"=O{r}*I{r}/STD_DRINK_L"), '0.0')
    if i % 2:
        for c in range(2, 6):
            S.cell(r, c).fill = fill(CREAM)
S.cell(R1 + 2, 2, "Fictional company & synthetic commercial data for portfolio purposes. Tax rates and rules are real (ATO/ABF, Aug 2026).").font = font(9, italic=True, color=GREY)
S.freeze_panes = S.cell(R0, 4)
S.auto_filter.ref = f"B6:P{R1}"
for nm, col in [("SKU_ID", "B"), ("SKU_NAME", "C"), ("SKU_CAT", "E"), ("SKU_BASE", "F"), ("SKU_UNITS", "G"),
                ("SKU_LPU", "H"), ("SKU_ABV", "I"), ("SKU_PP", "J"), ("SKU_CDSN", "K"), ("SKU_CDSFROM", "L"),
                ("SKU_CLIFF", "M"), ("SKU_VOL", "N"), ("SKU_LPACK", "O"), ("SKU_STD", "P")]:
    name(nm, f"'SKU Master'!${col}${R0}:${col}${R1}")
# excise-eligible SKUs for ABV Lab are the first 16 rows (beer, RTD, spirits)
name("ABV_SKUS", f"'SKU Master'!$C${R0}:$C${R0+15}")

# =====================================================================
# PRICE ENGINE
# =====================================================================
P = ws_map["Price Engine"]
widths = {"B": 7, "C": 34, "D": 9, "E": 17}
for c in range(6, 6 + 27):
    widths[CL(c)] = 10.5
base_sheet(P, "Price Engine", "Excise -> deposit -> distributor -> WET -> retailer -> GST -> shelf, for every SKU and period",
           widths)
blocks = ["Rate $/LAL", "Excise $", "Deposit $", "Wholesale $", "WET $", "Retail ex-GST $", "Shelf (raw) $", "Shelf price $", "Tax share"]
NB = len(blocks)
hdr(P, 6, 2, ["", "", "", ""], color=WHITE)
for p in range(3):
    c0 = 6 + p * NB
    P.merge_cells(start_row=6, start_column=c0, end_row=6, end_column=c0 + NB - 1)
    cell = P.cell(6, c0, periods[p][0])
    cell.font = font(10, True, WHITE)
    cell.fill = fill([BURG, "7A2E45", "9C4A60"][p])
    cell.alignment = Alignment(horizontal="center")
hdr(P, 7, 2, ["SKU", "Product", "Category", "Tax code"] + blocks * 3, height=32)
for i in range(NS):
    r = 8 + i
    sr = R0 + i
    link(P.cell(r, 2, f"='SKU Master'!B{sr}"))
    link(P.cell(r, 3, f"='SKU Master'!C{sr}"))
    link(P.cell(r, 4, f"='SKU Master'!E{sr}"))
    calc(P.cell(r, 5, f"=IF('SKU Master'!F{sr}=\"Beer\",IF('SKU Master'!I{sr}<=0.03,\"BEER_PKG_LOW\",IF('SKU Master'!I{sr}<=0.035,\"BEER_PKG_MID\",\"BEER_PKG_FULL\")),IF('SKU Master'!F{sr}=\"Spirits\",\"SPIRITS\",IF('SKU Master'!F{sr}=\"Brandy\",\"BRANDY\",IF('SKU Master'!F{sr}=\"WET\",\"WET\",\"NONE\"))))"))
    for p in range(3):
        c0 = 6 + p * NB
        L = lambda k: f"{CL(c0 + k)}{r}"
        dm = f"INDEX(CAT_DM,MATCH($D{r},CAT_LIST,0))"
        rm = f"INDEX(CAT_RM,MATCH($D{r},CAT_LIST,0))"
        calc(P.cell(r, c0, f"=INDEX(RATE_CARD,MATCH($E{r},TAX_CODES,0),{p+1})"), FMT_USD)
        calc(P.cell(r, c0 + 1, f"=IF('SKU Master'!$F{sr}=\"Beer\",'SKU Master'!$O{sr}*MAX('SKU Master'!$I{sr}-BEER_FREE,0),'SKU Master'!$O{sr}*'SKU Master'!$I{sr})*{L(0)}"), FMT_USD)
        calc(P.cell(r, c0 + 2, f"=IF(OR('SKU Master'!$L{sr}=\"Now\",{p+1}>=3),'SKU Master'!$K{sr}*CDS_COST,0)"), FMT_USD)
        calc(P.cell(r, c0 + 3, f"=('SKU Master'!$J{sr}+{L(1)}+{L(2)})*(1+{dm})"), FMT_USD)
        calc(P.cell(r, c0 + 4, f"=IF('SKU Master'!$F{sr}=\"WET\",WET_RATE*{L(3)},0)"), FMT_USD)
        calc(P.cell(r, c0 + 5, f"=({L(3)}+{L(4)})*(1+{rm})"), FMT_USD)
        calc(P.cell(r, c0 + 6, f"={L(5)}*(1+GST_RATE)"), FMT_USD)
        calc(P.cell(r, c0 + 7, f"=CEILING({L(6)},1)-0.01"), FMT_USD, bold=True, color=BURG)
        calc(P.cell(r, c0 + 8, f"=({L(1)}+{L(4)}+{L(6)}-{L(5)})/{L(6)}"), FMT_PCT)
P.freeze_panes = "F8"
PE0, PE1 = 8, 8 + NS - 1
for p in range(3):
    c0 = 6 + p * NB
    name(f"PE_EX{p+1}", f"'Price Engine'!${CL(c0+1)}${PE0}:${CL(c0+1)}${PE1}")
    name(f"PE_CDS{p+1}", f"'Price Engine'!${CL(c0+2)}${PE0}:${CL(c0+2)}${PE1}")
    name(f"PE_RAW{p+1}", f"'Price Engine'!${CL(c0+6)}${PE0}:${CL(c0+6)}${PE1}")
    name(f"PE_SHELF{p+1}", f"'Price Engine'!${CL(c0+7)}${PE0}:${CL(c0+7)}${PE1}")
P.cell(PE1 + 2, 2, "Retail convention: shelf price = next whole dollar minus 1c (e.g. $59.99). Tax share = (excise + WET + GST) / shelf.").font = font(9, italic=True, color=GREY)

# =====================================================================
# SIMULATOR
# =====================================================================
M = ws_map["Simulator"]
base_sheet(M, "Excise-to-Shelf Simulator", "Pick a product and a tax period - see exactly who takes what from every pack sold",
           {"B": 30, "C": 16, "D": 13, "E": 3, "F": 36, "G": 14, "H": 12, "I": 3, "J": 12, "K": 12, "L": 12, "M": 12, "N": 12, "O": 12, "R": 3, "S": 22, "T": 16})
section(M, "B6", "1 · Choose")
ctrl = [("Product", skus[0][1], None), ("Tax period", periods[0][0], None),
        ("Producer price change", 0, FMT_PCT), ("ABV override (blank = actual)", None, FMT_ABV),
        ("Distributor margin override", None, FMT_PCT), ("Retailer margin override", None, FMT_PCT)]
for i, (lab, v, fmt) in enumerate(ctrl):
    r = 7 + i
    calc(M.cell(r, 2, lab), bold=True)
    c = M.cell(r, 3, v)
    inp(c, fmt)
    M.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4) if i < 2 else None
M["C7"].alignment = Alignment(horizontal="left")
dv1 = DataValidation(type="list", formula1="=SKU_NAME", allow_blank=False)
dv2 = DataValidation(type="list", formula1="=PERIODS", allow_blank=False)
dv3 = DataValidation(type="decimal", operator="between", formula1="-0.5", formula2="0.5")
dv3.error, dv3.errorTitle = "Enter a % between -50% and +50%", "Out of range"
dv4 = DataValidation(type="decimal", operator="between", formula1="0", formula2="0.6", allow_blank=True)
for dv in (dv1, dv2, dv3, dv4):
    M.add_data_validation(dv)
dv1.add("C7"); dv2.add("C8"); dv3.add("C9"); dv4.add("C10"); dv4.add("C11"); dv4.add("C12")
M["D9"].value = "← try +5% or -5%"
M["D9"].font = font(8, italic=True, color=GREY)

# hidden-ish lookups (col J..K block 'Under the hood')
section(M, "S6", "Under the hood (lookups)")
uh = [
    ("Row #", "=MATCH($C$7,SKU_NAME,0)", "0"),
    ("Period #", "=MATCH($C$8,PERIODS,0)", "0"),
    ("SKU ID", "=INDEX(SKU_ID,K7)", None),
    ("Category", "=INDEX(SKU_CAT,K7)", None),
    ("Tax base", "=INDEX(SKU_BASE,K7)", None),
    ("Litres / pack", "=INDEX(SKU_LPACK,K7)", "0.00"),
    ("ABV used", "=IF(ISNUMBER($C$10),$C$10,INDEX(SKU_ABV,K7))", FMT_ABV),
    ("Tax code", "=IF(K11=\"Beer\",IF(K13<=0.03,\"BEER_PKG_LOW\",IF(K13<=0.035,\"BEER_PKG_MID\",\"BEER_PKG_FULL\")),IF(K11=\"Spirits\",\"SPIRITS\",IF(K11=\"Brandy\",\"BRANDY\",IF(K11=\"WET\",\"WET\",\"NONE\"))))", None),
    ("Rate $/LAL", "=INDEX(RATE_CARD,MATCH(K14,TAX_CODES,0),K8)", FMT_USD),
    ("Distributor margin", "=IF(ISNUMBER($C$11),$C$11,INDEX(CAT_DM,MATCH(K10,CAT_LIST,0)))", FMT_PCT),
    ("Retailer margin", "=IF(ISNUMBER($C$12),$C$12,INDEX(CAT_RM,MATCH(K10,CAT_LIST,0)))", FMT_PCT),
    ("Price point to defend", "=INDEX(SKU_CLIFF,K7)", FMT_USD0),
    ("Std drinks / pack", "=K12*K13/STD_DRINK_L", "0.0"),
]
for i, (lab, f, fmt) in enumerate(uh):
    r = 7 + i
    M.cell(r, 19, lab).font = font(9, color=GREY)
    calc(M.cell(r, 20, f), fmt)
    M.cell(r, 20).font = font(9, color=GREY)

section(M, "B15", "2 · Price build-up (per pack)")
hdr(M, 16, 2, ["Component", "$ per pack", "% of shelf"])
build = [
    ("Producer price (ex-tax)", f"=INDEX(SKU_PP,K7)*(1+$C$9)"),
    ("Excise duty", "=IF(K11=\"Beer\",K12*MAX(K13-BEER_FREE,0),K12*K13)*K15"),
    ("Container deposit", "=IF(OR(INDEX(SKU_CDSFROM,K7)=\"Now\",K8>=3),INDEX(SKU_CDSN,K7)*CDS_COST,0)"),
    ("Distributor margin", "=(C17+C18+C19)*K16"),
    ("Wine Equalisation Tax", "=IF(K11=\"WET\",WET_RATE*(C17+C18+C19+C20),0)"),
    ("Retailer margin", "=(C17+C18+C19+C20+C21)*K17"),
    ("GST", "=SUM(C17:C22)*GST_RATE"),
    ("Price-ending rounding", "=CEILING(SUM(C17:C23),1)-0.01-SUM(C17:C23)"),
]
for i, (lab, f) in enumerate(build):
    r = 17 + i
    calc(M.cell(r, 2, lab))
    calc(M.cell(r, 3, f), FMT_USD)
    calc(M.cell(r, 4, f"=C{r}/$C$25"), FMT_PCT)
    if lab in ("Excise duty", "Wine Equalisation Tax", "GST"):
        for c in (2, 3, 4):
            M.cell(r, c).fill = fill(LIGHT_BURG)
calc(M.cell(25, 2, "Shelf price"), bold=True)
calc(M.cell(25, 3, "=SUM(C17:C24)"), FMT_USD, bold=True, color=BURG)
calc(M.cell(25, 4, "=C25/$C$25"), FMT_PCT, bold=True)
for c in (2, 3, 4):
    M.cell(25, c).fill = fill(LIGHT_GOLD)
M.cell(26, 2, "Shaded rows = government take").font = font(8, italic=True, color=GREY)

section(M, "F6", "3 · Headline numbers")
kpis = [
    ("Shelf price", "=C25", FMT_USD),
    ("Government take (excise+WET+GST)", "=C18+C21+C23", FMT_USD),
    ("Government share of shelf", "=G8/G7", FMT_PCT),
    ("$1 of excise becomes … at shelf", "=(1+K16)*IF(K11=\"WET\",1+WET_RATE,1)*(1+K17)*(1+GST_RATE)", FMT_X),
    ("Price per standard drink", "=C25/K19", FMT_USD),
    ("Price point to defend", "=K18", FMT_USD0),
    ("Headroom to price point", "=K18-C25", FMT_USD),
    ("Price-point status", "=IF(C25>K18,\"OVER the price point\",IF(K18-C25<1,\"On the edge (<$1)\",\"Safe\"))", None),
]
for i, (lab, f, fmt) in enumerate(kpis):
    r = 7 + i
    c1 = M.cell(r, 6, lab)
    c1.font = font(10, True, INK)
    c1.fill = fill(CREAM)
    c1.border = BOX
    c2 = M.cell(r, 7, f)
    calc(c2, fmt, bold=True, color=BURG)
    c2.font = font(12, True, BURG)
    c2.fill = fill(CREAM)
    M.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
    M.row_dimensions[r].height = 20
M.conditional_formatting.add("G14", FormulaRule(formula=['LEFT($G$14,4)="OVER"'], fill=fill("F4C7C3"), font=Font(name=F, bold=True, color="9C0006")))
M.conditional_formatting.add("G14", FormulaRule(formula=['LEFT($G$14,2)="On"'], fill=fill("FFEB9C"), font=Font(name=F, bold=True, color="9C5700")))
M.conditional_formatting.add("G14", FormulaRule(formula=['$G$14="Safe"'], fill=fill("C6EFCE"), font=Font(name=F, bold=True, color="006100")))
M.cell(15, 6, "Insight: distributor and retailer margins are charged on top of tax, so every excise rise is amplified before it reaches the shopper.").font = font(9, italic=True, color=BURG)
M.merge_cells("F15:H16")
M["F15"].alignment = Alignment(wrap_text=True, vertical="top")

# waterfall helper (rows 29-37)
section(M, "B28", "Chart data - waterfall")
hdr(M, 29, 2, ["Step", "Base (hidden)", "Value"])
steps = ["Producer", "Excise", "Deposit", "Distributor", "WET", "Retailer", "GST", "Rounding", "Shelf"]
for i, st in enumerate(steps):
    r = 30 + i
    calc(M.cell(r, 2, st))
    if st == "Shelf":
        calc(M.cell(r, 3, 0), FMT_USD)
        calc(M.cell(r, 4, "=C25"), FMT_USD)
    else:
        calc(M.cell(r, 3, "=0" if i == 0 else f"=SUM($C$17:C{16+i})"), FMT_USD)
        calc(M.cell(r, 4, f"=C{17+i}"), FMT_USD)
    for c in (2, 3, 4):
        M.cell(r, c).font = font(9, color=GREY)

wf = BarChart()
wf.type = "col"
wf.grouping = "stacked"
wf.overlap = 100
wf.title = "Where the shelf price comes from"
wf.add_data(Reference(M, min_col=3, max_col=4, min_row=29, max_row=38), titles_from_data=True)
wf.set_categories(Reference(M, min_col=2, min_row=30, max_row=38))
wf.series[0].graphicalProperties.noFill = True
wf.series[0].graphicalProperties.line.noFill = True
wf.series[1].graphicalProperties.solidFill = BURG
wf.series[1].dLbls = DataLabelList()
wf.series[1].dLbls.showVal = True
wf.series[1].dLbls.numFmt = '$0.00'
wf.series[1].dLbls.showSerName = False; wf.series[1].dLbls.position = 'inEnd'; wf.series[1].dLbls.txPr = _lbl_txt(GOLD); wf.series[1].dLbls.showCatName = False; wf.series[1].dLbls.showLegendKey = False
from openpyxl.chart.label import DataLabelList as _D
wf.legend = None
wf.y_axis.majorGridlines = None
wf.y_axis.numFmt = '$0'
wf.gapWidth = 40
wf.height, wf.width = 9, 16.5
M.add_chart(wf, "F18")

dn = DoughnutChart()
dn.title = "Who keeps your money"
section(M, "S22", "Chart data - split")
hdr(M, 23, 19, ["Party", "$ per pack"], height=None)
split = [("Producer", "=C17"), ("Distributor", "=C20"), ("Retailer", "=C22+C24"), ("Government", "=C18+C21+C23"), ("Deposit scheme", "=C19")]
for i, (pty, f) in enumerate(split):
    r = 24 + i
    calc(M.cell(r, 19, pty))
    calc(M.cell(r, 20, f), FMT_USD)
    for c in (19, 20):
        M.cell(r, c).font = font(9, color=GREY)
dn.add_data(Reference(M, min_col=20, min_row=23, max_row=28), titles_from_data=True)
dn.set_categories(Reference(M, min_col=19, min_row=24, max_row=28))
dn.holeSize = 55
from openpyxl.chart.series import DataPoint
for idx, col in enumerate([GOLD, "4F6D7A", "A3B18A", BURG, "BDB5A6"]):
    pt = DataPoint(idx=idx)
    pt.graphicalProperties.solidFill = col
    dn.series[0].dPt.append(pt)
dn.dataLabels = DataLabelList()
dn.dataLabels.showPercent = True; dn.dataLabels.showVal = False; dn.dataLabels.showCatName = False; dn.dataLabels.showSerName = False; dn.dataLabels.showLegendKey = False
dn.legend.position = "r"
dn.height, dn.width = 8.2, 13
M.add_chart(dn, "J6")
import re as _re
for _row in M.iter_rows():
    for _c in _row:
        if isinstance(_c.value, str) and _c.value.startswith("=") and _c.column != 20 or (isinstance(_c.value, str) and _c.value.startswith("=") and _c.column == 20 and _c.row < 22):
            _c.value = _re.sub(r"(?<![A-Z$_])K(\d+)\b", r"T\1", _c.value)
M.freeze_panes = "A5"

# =====================================================================
# CLIFF SCANNER
# =====================================================================
C = ws_map["Cliff Scanner"]
base_sheet(C, "Price-Point Cliff Scanner", "Which packs get pushed over the price shoppers won't pay - and when",
           {"B": 7, "C": 36, "D": 10, "E": 12, "F": 12, "G": 12, "H": 12, "I": 12, "J": 18, "K": 44})
section(C, "B6", "Portfolio summary")
summ = [("Already over", "Already over"), ("Crosses Feb-27", "Crosses Feb-27"), ("Crosses Aug-27", "Crosses Aug-27"), ("Safe", "Safe")]
hdr(C, 7, 2, ["", "Status", "# SKUs", "Annual packs", "% of volume"])
C.merge_cells("B7:B7")
SR0 = 13
SR1 = SR0 + NS - 1
for i, (lab, key) in enumerate(summ):
    r = 8 + i
    calc(C.cell(r, 3, lab), bold=True)
    calc(C.cell(r, 4, f'=COUNTIF($J${SR0}:$J${SR1},"{key}")'), FMT_NUM)
    calc(C.cell(r, 5, f'=SUMIF($J${SR0}:$J${SR1},"{key}",SKU_VOL)'), FMT_NUM)
    calc(C.cell(r, 6, f'=E{r}/SUM($E$8:$E$11)'), FMT_PCT)
hdr(C, 12, 2, ["SKU", "Product", "Category", "Price point", "Aug 2026 shelf", "Feb 2027 shelf", "Aug 2027 shelf",
               "Headroom today", "Status", "Recommended action"])
for i in range(NS):
    r = SR0 + i
    pr = PE0 + i
    sr = R0 + i
    link(C.cell(r, 2, f"='SKU Master'!B{sr}"))
    link(C.cell(r, 3, f"='SKU Master'!C{sr}"))
    link(C.cell(r, 4, f"='SKU Master'!E{sr}"))
    link(C.cell(r, 5, f"='SKU Master'!M{sr}"), FMT_USD0)
    for p in range(3):
        link(C.cell(r, 6 + p, f"=INDEX(PE_SHELF{p+1},{i+1})"), FMT_USD)
    calc(C.cell(r, 9, f"=E{r}-F{r}"), FMT_USD)
    calc(C.cell(r, 10, f'=IF(F{r}>E{r},"Already over",IF(G{r}>E{r},"Crosses Feb-27",IF(H{r}>E{r},"Crosses Aug-27","Safe")))'), bold=True)
    calc(C.cell(r, 11, f'=IF(J{r}="Already over","Reset: new pack size or re-tier the price point",IF(J{r}="Crosses Feb-27","Act now: test ABV trim / pack re-engineering (ABV Lab)",IF(J{r}="Crosses Aug-27","Plan: excise + deposit double-hit - lock promo calendar","Hold: monitor next indexation")))'))
    C.cell(r, 11).font = font(9, color=INK)
rng = f"J{SR0}:J{SR1}"
C.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Already over"'], fill=fill("F4C7C3"), font=Font(name=F, bold=True, color="9C0006")))
C.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Crosses Feb-27"'], fill=fill("FCD5B4"), font=Font(name=F, bold=True, color="974706")))
C.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Crosses Aug-27"'], fill=fill("FFEB9C"), font=Font(name=F, bold=True, color="9C5700")))
C.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Safe"'], fill=fill("C6EFCE"), font=Font(name=F, bold=True, color="006100")))
C.conditional_formatting.add(f"I{SR0}:I{SR1}", DataBarRule(start_type="min", end_type="max", color=GOLD, showValue=True))
for p in range(3):
    col = CL(6 + p)
    C.conditional_formatting.add(f"{col}{SR0}:{col}{SR1}", FormulaRule(formula=[f"{col}{SR0}>$E{SR0}"], font=Font(name=F, bold=True, color="C00000")))
C.auto_filter.ref = f"B12:K{SR1}"
C.freeze_panes = f"D{SR0}"
bc = BarChart()
bc.type = "bar"
bc.title = "SKUs by price-point status"
bc.add_data(Reference(C, min_col=4, min_row=7, max_row=11), titles_from_data=True)
bc.set_categories(Reference(C, min_col=3, min_row=8, max_row=11))
bc.series[0].graphicalProperties.solidFill = BURG
bc.series[0].dLbls = DataLabelList(); bc.series[0].dLbls.showVal = True; bc.series[0].dLbls.showSerName = False; bc.series[0].dLbls.showCatName = False; bc.series[0].dLbls.showLegendKey = False
bc.legend = None
bc.y_axis.majorGridlines = None
bc.x_axis.scaling.orientation = "maxMin"
bc.height, bc.width = 5.5, 12
C.add_chart(bc, "M6")
C.row_dimensions[5].height = 6

# =====================================================================
# ABV LAB
# =====================================================================
L = ws_map["ABV Lab"]
base_sheet(L, "ABV Lab - the VB question", "Excise is charged per litre of alcohol. Trim the ABV, save tax... but how much before shoppers notice?",
           {"B": 40, "C": 16, "D": 11, "E": 12, "F": 12, "G": 12, "H": 18, "I": 14, "J": 14, "K": 14, "L": 14})
section(L, "B6", "Choose")
labs = [("Product (beer / RTD / spirits)", skus[0][1], None), ("Tax period", periods[0][0], None),
        ("ABV step size", 0.001, '0.0%')]
for i, (lab, v, fmt) in enumerate(labs):
    r = 7 + i
    calc(L.cell(r, 2, lab), bold=True)
    inp(L.cell(r, 3, v), fmt)
L.merge_cells("C7:F7"); L.merge_cells("C8:D8")
dva = DataValidation(type="list", formula1="=ABV_SKUS"); L.add_data_validation(dva); dva.add("C7")
dvb = DataValidation(type="list", formula1="=PERIODS"); L.add_data_validation(dvb); dvb.add("C8")
dvc = DataValidation(type="list", formula1='"0.001,0.002,0.005,0.01"'); L.add_data_validation(dvc); dvc.add("C9")
L["E9"].value = "0.1% for beer & RTD · 0.5-1% for spirits"
L["E9"].font = font(8, italic=True, color=GREY)

ab = [("Row #", "=MATCH($C$7,SKU_NAME,0)", "0"), ("Period #", "=MATCH($C$8,PERIODS,0)", "0"),
      ("Category", "=INDEX(SKU_CAT,I7)", None), ("Tax base", "=INDEX(SKU_BASE,I7)", None),
      ("Litres / pack", "=INDEX(SKU_LPACK,I7)", "0.00"), ("Current ABV", "=INDEX(SKU_ABV,I7)", FMT_ABV),
      ("Producer price", "=INDEX(SKU_PP,I7)", FMT_USD), ("COGS %", "=INDEX(CAT_COGS,MATCH(I9,CAT_LIST,0))", FMT_PCT),
      ("Elasticity", "=INDEX(CAT_EL,MATCH(I9,CAT_LIST,0))", "0.00"), ("Annual packs", "=INDEX(SKU_VOL,I7)", FMT_NUM),
      ("Deposit $/pack", "=IF(OR(INDEX(SKU_CDSFROM,I7)=\"Now\",I8>=3),INDEX(SKU_CDSN,I7)*CDS_COST,0)", FMT_USD),
      ("Chain multiplier", "=(1+INDEX(CAT_DM,MATCH(I9,CAT_LIST,0)))*(1+INDEX(CAT_RM,MATCH(I9,CAT_LIST,0)))*(1+GST_RATE)", '0.000')]
section(L, "H6", "Under the hood")
for i, (lab, f, fmt) in enumerate(ab):
    r = 7 + i
    L.cell(r, 8, lab).font = font(9, color=GREY)
    calc(L.cell(r, 9, f), fmt)
    L.cell(r, 9).font = font(9, color=GREY)

section(L, "B20", "ABV grid - two strategies")
hdr(L, 21, 2, ["ABV", "Tax code", "Rate $/LAL", "Excise / pack", "Tax saved / pack", "Backlash volume loss",
               "A: Shelf price (pass saving on)", "A: Annual packs", "A: Gross profit", "B: Annual packs (hold shelf)", "B: Gross profit"], height=44)
G0 = 22
offs = list(range(-10, 6))
for i, k in enumerate(offs):
    r = G0 + i
    calc(L.cell(r, 2, f"=MAX($I$12+({k})*$C$9,0)"), '0.0%', bold=(k == 0))
    calc(L.cell(r, 3, f"=IF($I$10=\"Beer\",IF(B{r}<=0.03,\"BEER_PKG_LOW\",IF(B{r}<=0.035,\"BEER_PKG_MID\",\"BEER_PKG_FULL\")),IF($I$10=\"Spirits\",\"SPIRITS\",IF($I$10=\"Brandy\",\"BRANDY\",\"NONE\")))"))
    calc(L.cell(r, 4, f"=INDEX(RATE_CARD,MATCH(C{r},TAX_CODES,0),$I$8)"), FMT_USD)
    calc(L.cell(r, 5, f"=IF($I$10=\"Beer\",$I$11*MAX(B{r}-BEER_FREE,0),$I$11*B{r})*D{r}"), FMT_USD)
    calc(L.cell(r, 6, f"=$E${G0+10}-E{r}"), FMT_USD)
    calc(L.cell(r, 7, f"=MIN(BACKLASH_K*(MAX($I$12-B{r},0)/0.001)^2,0.9)"), FMT_PCT)
    calc(L.cell(r, 8, f"=($I$13+E{r}+$I$17)*$I$18"), FMT_USD)
    calc(L.cell(r, 9, f"=$I$16*MAX(1+$I$15*(H{r}/$H${G0+10}-1),0)*(1-G{r})"), FMT_NUM)
    calc(L.cell(r, 10, f"=$I$13*(1-$I$14)*I{r}"), FMT_USD0)
    calc(L.cell(r, 11, f"=$I$16*(1-G{r})"), FMT_NUM)
    calc(L.cell(r, 12, f"=($I$13*(1-$I$14)+F{r})*K{r}"), FMT_USD0)
    if k == 0:
        for c in range(2, 13):
            L.cell(r, c).fill = fill(LIGHT_GOLD)
G1 = G0 + len(offs) - 1
for col in ("J", "L"):
    L.conditional_formatting.add(f"{col}{G0}:{col}{G1}", FormulaRule(formula=[f"{col}{G0}=MAX(${col}${G0}:${col}${G1})"], fill=fill("C6EFCE"), font=Font(name=F, bold=True, color="006100")))
L.cell(G1 + 1, 2, "Gold row = today's recipe. Green = best gross profit for each strategy. Backlash only applies to ABV cuts.").font = font(8, italic=True, color=GREY)

section(L, "B13", "Verdict")
verdict = [
    ("Best ABV - A (pass saving to shopper)", f"=INDEX($B${G0}:$B${G1},MATCH(MAX($J${G0}:$J${G1}),$J${G0}:$J${G1},0))", '0.0%'),
    ("Best ABV - B (hold shelf, bank saving)", f"=INDEX($B${G0}:$B${G1},MATCH(MAX($L${G0}:$L${G1}),$L${G0}:$L${G1},0))", '0.0%'),
    ("Extra gross profit vs today (best of A/B)", f"=MAX(MAX($J${G0}:$J${G1}),MAX($L${G0}:$L${G1}))-$J${G0+10}", FMT_USD0),
    ("Excise saved per pack at best B", f"=INDEX($F${G0}:$F${G1},MATCH(MAX($L${G0}:$L${G1}),$L${G0}:$L${G1},0))", FMT_USD),
]
for i, (lab, f, fmt) in enumerate(verdict):
    r = 14 + i
    c1 = L.cell(r, 2, lab); c1.font = font(10, True); c1.fill = fill(CREAM); c1.border = BOX
    c2 = L.cell(r, 3, f); calc(c2, fmt); c2.font = font(12, True, BURG); c2.fill = fill(CREAM)
L.cell(18, 2, "Lesson from VB (cut from 4.9% to 4.8% in 2007 and 4.6% in 2009; restored to 4.9% in 2012 after losing share): the tax saving is real, but drinkers notice.").font = font(8, italic=True, color=GREY)

lc = LineChart()
lc.title = "Annual gross profit by ABV"
lc.add_data(Reference(L, min_col=10, min_row=21, max_row=G1), titles_from_data=True)
lc.add_data(Reference(L, min_col=12, min_row=21, max_row=G1), titles_from_data=True)
lc.set_categories(Reference(L, min_col=2, min_row=G0, max_row=G1))
lc.series[0].graphicalProperties.line.solidFill = GOLD
lc.series[1].graphicalProperties.line.solidFill = BURG
for s in lc.series:
    s.graphicalProperties.line.width = 28000
lc.y_axis.numFmt = '$#,##0'
lc.y_axis.majorGridlines = None
lc.x_axis.number_format = '0.0%'
lc.legend.position = "b"
lc.height, lc.width = 8, 15
L.add_chart(lc, "N6")
L.freeze_panes = "A5"

# =====================================================================
# SCENARIO PLANNER
# =====================================================================
X = ws_map["Scenario Planner"]
base_sheet(X, "Scenario Planner", "Next excise rise: pass it on, split it, or absorb it? Profit impact by category.",
           {"B": 36, "C": 14, "D": 13, "E": 13, "F": 13, "G": 13, "H": 13, "I": 13, "J": 13, "K": 13, "L": 13, "M": 13, "N": 13, "O": 13, "P": 13, "Q": 13})
section(X, "B6", "Choose")
strategies = [("Full pass-through (100%)", 1.0), ("Split 50/50", 0.5), ("Absorb the rise (0%)", 0.0), ("Custom", None)]
calc(X.cell(7, 2, "Compare today against"), bold=True); inp(X.cell(7, 3, periods[1][0]))
calc(X.cell(8, 2, "Pricing strategy"), bold=True); inp(X.cell(8, 3, strategies[0][0]))
calc(X.cell(9, 2, "Custom pass-through %"), bold=True); inp(X.cell(9, 3, 0.75), FMT_PCT)
X.merge_cells("C7:D7"); X.merge_cells("C8:D8")
for i, (sname, sv) in enumerate(strategies):
    X.cell(7 + i, 7, sname).font = font(9, color=GREY)
    c = X.cell(7 + i, 8, sv if sv is not None else "=C9")
    c.font = font(9, color=GREY); c.number_format = FMT_PCT
dvp = DataValidation(type="list", formula1='"Feb 2027 (forecast),Aug 2027 (forecast)"'); X.add_data_validation(dvp); dvp.add("C7")
dvs = DataValidation(type="list", formula1="=$G$7:$G$10"); X.add_data_validation(dvs); dvs.add("C8")
calc(X.cell(10, 2, "Period # / pass-through used"), bold=True)
calc(X.cell(10, 3, "=MATCH(C7,PERIODS,0)"), "0")
calc(X.cell(10, 4, "=INDEX($H$7:$H$10,MATCH(C8,$G$7:$G$10,0))"), FMT_PCT)

# SKU-level table
T0 = 24
T1 = T0 + NS - 1
hdr(X, T0 - 1, 2, ["Product", "Category", "Base packs", "Today shelf (raw)", "Target shelf (full)", "Tax rise / pack",
                    "Scenario shelf", "Crosses price point?", "Scenario packs", "Base GP", "Scenario GP",
                    "GP if 100%", "GP if 50%", "GP if 0%", "Best for this SKU", "Best GP"], height=44)
def gp_expr(r, pexpr):
    shelf_new = f"($E{r}+({pexpr})*($F{r}-$E{r}))"
    crossed = f"AND(CEILING({shelf_new},1)-0.01>INDEX(SKU_CLIFF,{r-T0+1}),CEILING($E{r},1)-0.01<=INDEX(SKU_CLIFF,{r-T0+1}))"
    vol = f"$D{r}*MAX(1+INDEX(CAT_EL,MATCH($C{r},CAT_LIST,0))*({shelf_new}/$E{r}-1),0)*(1-IF({crossed},CLIFF_PEN,0))"
    unit = f"(INDEX(SKU_PP,{r-T0+1})*(1-INDEX(CAT_COGS,MATCH($C{r},CAT_LIST,0)))-(1-({pexpr}))*$G{r})"
    return f"={unit}*{vol}"
for i in range(NS):
    r = T0 + i
    link(X.cell(r, 2, f"=INDEX(SKU_NAME,{i+1})"))
    link(X.cell(r, 3, f"=INDEX(SKU_CAT,{i+1})"))
    link(X.cell(r, 4, f"=INDEX(SKU_VOL,{i+1})"), FMT_NUM)
    link(X.cell(r, 5, f"=INDEX(PE_RAW1,{i+1})"), FMT_USD)
    calc(X.cell(r, 6, f"=IF($C$10=2,INDEX(PE_RAW2,{i+1}),INDEX(PE_RAW3,{i+1}))"), FMT_USD)
    calc(X.cell(r, 7, f"=IF($C$10=2,INDEX(PE_EX2,{i+1})+INDEX(PE_CDS2,{i+1}),INDEX(PE_EX3,{i+1})+INDEX(PE_CDS3,{i+1}))-INDEX(PE_EX1,{i+1})-INDEX(PE_CDS1,{i+1})"), FMT_USD)
    calc(X.cell(r, 8, f"=CEILING(E{r}+$D$10*(F{r}-E{r}),1)-0.01"), FMT_USD)
    calc(X.cell(r, 9, f"=IF(AND(H{r}>INDEX(SKU_CLIFF,{i+1}),CEILING(E{r},1)-0.01<=INDEX(SKU_CLIFF,{i+1})),\"Yes\",\"-\")"))
    calc(X.cell(r, 10, f"=D{r}*MAX(1+INDEX(CAT_EL,MATCH(C{r},CAT_LIST,0))*((E{r}+$D$10*(F{r}-E{r}))/E{r}-1),0)*(1-IF(I{r}=\"Yes\",CLIFF_PEN,0))"), FMT_NUM)
    calc(X.cell(r, 11, f"=INDEX(SKU_PP,{i+1})*(1-INDEX(CAT_COGS,MATCH(C{r},CAT_LIST,0)))*D{r}"), FMT_USD0)
    calc(X.cell(r, 12, gp_expr(r, "$D$10")), FMT_USD0)
    calc(X.cell(r, 13, gp_expr(r, "1")), FMT_USD0)
    calc(X.cell(r, 14, gp_expr(r, "0.5")), FMT_USD0)
    calc(X.cell(r, 15, gp_expr(r, "0")), FMT_USD0)
    calc(X.cell(r, 16, f'=IF(MAX(M{r}:O{r})-MIN(M{r}:O{r})<1,"No tax rise",INDEX({{"Pass 100%","Split 50/50","Absorb"}},MATCH(MAX(M{r}:O{r}),M{r}:O{r},0)))'), bold=True, color=BURG)
    calc(X.cell(r, 17, f"=MAX(M{r}:O{r})"), FMT_USD0)
X.conditional_formatting.add(f"P{T0}:P{T1}", FormulaRule(formula=[f'OR(P{T0}="Absorb",P{T0}="Split 50/50")'], fill=fill(LIGHT_GOLD), font=Font(name=F, bold=True, color=BURG)))
X.conditional_formatting.add(f"I{T0}:I{T1}", CellIsRule(operator="equal", formula=['"Yes"'], fill=fill("F4C7C3"), font=Font(name=F, bold=True, color="9C0006")))

section(X, "B12", "Results by category")
hdr(X, 13, 2, ["Category", "Base packs", "Scenario packs", "Volume change", "Base GP", "Scenario GP", "GP change",
               "GP if 100%", "GP if 50%", "GP if 0%", "Best strategy"], height=36)
for i, (cname, *_r) in enumerate(cats):
    r = 14 + i
    calc(X.cell(r, 2, cname), bold=True)
    calc(X.cell(r, 3, f"=SUMIF($C${T0}:$C${T1},B{r},$D${T0}:$D${T1})"), FMT_NUM)
    calc(X.cell(r, 4, f"=SUMIF($C${T0}:$C${T1},B{r},$J${T0}:$J${T1})"), FMT_NUM)
    calc(X.cell(r, 5, f"=D{r}/C{r}-1"), FMT_PCT)
    calc(X.cell(r, 6, f"=SUMIF($C${T0}:$C${T1},B{r},$K${T0}:$K${T1})"), FMT_USD0)
    calc(X.cell(r, 7, f"=SUMIF($C${T0}:$C${T1},B{r},$L${T0}:$L${T1})"), FMT_USD0)
    calc(X.cell(r, 8, f"=G{r}-F{r}"), FMT_USD0, bold=True)
    for j, col in enumerate(["M", "N", "O"]):
        calc(X.cell(r, 9 + j, f"=SUMIF($C${T0}:$C${T1},B{r},${col}${T0}:${col}${T1})"), FMT_USD0)
    calc(X.cell(r, 12, f'=IF(MAX(I{r}:K{r})-MIN(I{r}:K{r})<1,"No tax rise",INDEX({{"Pass 100%","Split 50/50","Absorb"}},MATCH(MAX(I{r}:K{r}),I{r}:K{r},0)))'), bold=True, color=BURG)
r = 19
calc(X.cell(r, 2, "Portfolio total"), bold=True)
for col in "CDFGHIJK":
    calc(X[f"{col}{r}"], FMT_USD0 if col in "FGHIJK" else FMT_NUM, bold=True)
    X[f"{col}{r}"].value = f"=SUM({col}14:{col}18)"
calc(X.cell(r, 5, "=D19/C19-1"), FMT_PCT, bold=True)
calc(X.cell(r, 12, '=INDEX({"Pass 100%","Split 50/50","Absorb"},MATCH(MAX(I19:K19),I19:K19,0))'), bold=True, color=BURG)
for c in range(2, 13):
    X.cell(19, c).fill = fill(LIGHT_GOLD)
X.conditional_formatting.add("H14:H19", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=F, bold=True, color="C00000")))
X.conditional_formatting.add("H14:H19", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=F, bold=True, color="006100")))
section(X, "J6", "Headline")
xk = [("Portfolio GP change", "=H19", FMT_USD0),
      ("Portfolio volume change", "=E19", FMT_PCT),
      ("SKUs pushed over a price point", f'=COUNTIF(I{T0}:I{T1},"Yes")', "0"),
      ("Smart mix uplift vs blanket pass-through", f"=SUM(Q{T0}:Q{T1})-I19", FMT_USD0)]
for i, (lab, f, fmt) in enumerate(xk):
    r = 7 + i
    c1 = X.cell(r, 10, lab); c1.font = font(10, True); c1.fill = fill(CREAM); c1.border = BOX
    X.merge_cells(start_row=r, start_column=10, end_row=r, end_column=12)
    c2 = X.cell(r, 13, f); calc(c2, fmt); c2.font = font(12, True, BURG); c2.fill = fill(CREAM)
X.cell(11, 10, "Smart mix = choose pass / split / absorb per SKU (see 'Best for this SKU' column).").font = font(8, italic=True, color=GREY)
X.cell(20, 2, "GP = producer gross profit (ex-tax price less COGS, less any tax rise absorbed). Volume responds to shelf price via category elasticity + cliff penalty.").font = font(8, italic=True, color=GREY)

sc = BarChart()
sc.type = "col"
sc.title = "Gross profit: today vs scenario"
sc.add_data(Reference(X, min_col=6, max_col=7, min_row=13, max_row=18), titles_from_data=True)
sc.set_categories(Reference(X, min_col=2, min_row=14, max_row=18))
sc.series[0].graphicalProperties.solidFill = "BDB5A6"
sc.series[1].graphicalProperties.solidFill = BURG
sc.y_axis.numFmt = '$#,##0'
sc.y_axis.majorGridlines = None
sc.legend.position = "b"
sc.height, sc.width = 7.5, 15
X.add_chart(sc, "N5")
X.freeze_panes = "A5"

# =====================================================================
# START HERE
# =====================================================================
H = ws_map["Start Here"]
base_sheet(H, "Excise-to-Shelf Price Simulator", "A commercial pricing tool for an Australian drinks company - built by Suhas Eedala",
           {"B": 4, "C": 30, "D": 90})
section(H, "C6", "The business problem")
H["C7"].value = ("Australians are drinking less every year, while alcohol excise rises automatically every February and August. "
                 "Margins are charged on top of tax, so each rise is amplified before it hits the shelf - pushing packs past the "
                 "price points shoppers won't cross. Which packs are at risk, and what should a brand or distributor do about it?")
H.merge_cells("C7:D9")
H["C7"].alignment = Alignment(wrap_text=True, vertical="top")
H["C7"].font = font(11)
section(H, "C11", "Market context (late 2026)")
facts = [
    ("Consumption", "Total alcohol consumption fell ~10% in the three years to 2024 (AIHW); FY24 saw the sharpest fall on record (IBISWorld) - via ABC News, Oct 2026."),
    ("Excise", "Packaged beer excise up 10.1% and spirits up 10.1% from Aug 2023 to Aug 2026 (ABF notices) - see Excise Rates tab."),
    ("Wine", "2.06bn litres of wine in storage (Jun-2025), ~262M litres above a healthy level (Wine Australia); FY26 exports 598M litres, a 22-year low."),
    ("No / Low", "Non-alcoholic beer is the fastest-growing segment - forecast to roughly double by 2028 (ABC News); beer at or below 1.15% ABV pays no excise."),
    ("Regulation", "Wine & spirit bottles join NSW, SA and NT container deposit schemes from mid-late 2027."),
]
for i, (k, v) in enumerate(facts):
    r = 12 + i
    H.cell(r, 3, k).font = font(10, True, BURG)
    H.cell(r, 4, v).font = font(10)
    H.cell(r, 4).alignment = Alignment(wrap_text=True, vertical="center")
    H.cell(r, 3).alignment = Alignment(vertical="center")
    H.row_dimensions[r].height = 28
section(H, "C18", "How to use it")
guide = [
    ("Simulator", "Pick any product + tax period. See the full price build-up, government share and price-point risk."),
    ("Cliff Scanner", "All 24 SKUs scanned against their defended price point across the next two excise rises."),
    ("ABV Lab", "The VB question: how far can you trim alcohol % to save excise before shoppers walk?"),
    ("Scenario Planner", "Pass the next rise on, split it or absorb it - gross-profit impact by category."),
    ("Price Engine", "The transparent calculation layer behind every number (audit trail)."),
    ("SKU Master / Excise Rates / Assumptions", "Data and drivers. Blue cells on yellow are inputs - change them and everything updates."),
]
for i, (k, v) in enumerate(guide):
    r = 19 + i
    c = H.cell(r, 3)
    tgt = k.split(" /")[0]
    c.value = f'=HYPERLINK("#\'{tgt}\'!A1","{k}")'
    c.font = font(10, True, BURG)
    H.cell(r, 4, v).font = font(10)
section(H, "C26", "Colour legend")
lg = [("Input you can change", INPUT_FILL, "0000FF"), ("Formula (do not overwrite)", WHITE, INK), ("Link from another sheet", WHITE, "008000"), ("Government take", LIGHT_BURG, INK)]
for i, (t, fc, tc) in enumerate(lg):
    r = 27 + i
    c = H.cell(r, 3, 123.45); c.number_format = "0.00"; c.fill = fill(fc); c.font = font(10, True, tc); c.border = BOX
    H.cell(r, 4, t).font = font(10)
section(H, "C32", "Data & honesty note")
H["C33"].value = ("Southerly Drinks Co. is a FICTIONAL company. SKU prices, volumes, margins and elasticities are synthetic, illustrative "
                  "assumptions. Excise rates, WET, GST and the beer 1.15% allowance are real (ATO / ABF, effective 3 Aug 2026). "
                  "Sources: ATO Excise duty rates for alcohol; ABF Customs Notices 2023/33-2026/30; ATO Calculating alcohol excise duty; "
                  "Wine Australia Export Report (Jun 2026); AGW container deposit update; ABC News 26 Oct 2026.")
H.merge_cells("C33:D36")
H["C33"].alignment = Alignment(wrap_text=True, vertical="top")
H["C33"].font = font(9, italic=True, color=GREY)

# print setup & finish
for ws in wb.worksheets:
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToHeight = 0
wb.active = 0
from pathlib import Path
out = str(Path(__file__).resolve().parent / "ProofPoint_Excise_to_Shelf_Simulator_build.xlsx")
wb.save(out)
print("saved", out)
