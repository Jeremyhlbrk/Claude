"""Build Inventory_KPIs_PowerBI.xlsx from the tidy CSVs in data/ (the same files the dashboard uses).

Usage: python pipeline/build_workbook.py [data_dir] [output.xlsx]
"""
import csv
import datetime as dt
import sys
from pathlib import Path

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parent.parent
DATA = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "data"
OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "Inventory_KPIs_PowerBI.xlsx"


def load(name, types):
    with open(DATA / f"{name}.csv", newline="") as f:
        rows = list(csv.DictReader(f))
    cols = list(types)
    out = []
    for r in rows:
        row = []
        for c in cols:
            v = r.get(c, "")
            t = types[c]
            if v in ("", "nan", "None"):
                row.append(None)
            elif t == "n":
                row.append(float(v))
            elif t == "i":
                row.append(int(float(v)))
            elif t == "d":
                row.append(dt.date.fromisoformat(v[:10]))
            else:
                row.append(v)
        out.append(row)
    return cols, out


F = "Arial"
base = Font(name=F, size=10); bold = Font(name=F, size=10, bold=True); h1 = Font(name=F, size=16, bold=True)
h2 = Font(name=F, size=12, bold=True); blue = Font(name=F, size=10, color="0000FF"); muted = Font(name=F, size=9, color="666666")
yellow = PatternFill("solid", fgColor="FFFF00"); head = PatternFill("solid", fgColor="1F3A5F")
hfont = Font(name=F, size=10, bold=True, color="FFFFFF"); wrap = Alignment(wrap_text=True, vertical="top")
MONEY = '$#,##0;($#,##0);-'; TURNS = '0.00"x"'; PCT = '+0.0%;-0.0%;0.0%'; DATE = "yyyy-mm-dd"; INT = "#,##0"

wb = openpyxl.Workbook()


def table(title, name, cols, rows, fmts=None, widths=None, inputs=(), notes=()):
    ws = wb.create_sheet(title)
    ws.append(cols)
    for r in rows:
        ws.append(r)
    n = len(rows) + 1
    t = Table(displayName=name, ref=f"A1:{get_column_letter(len(cols))}{n}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(t)
    fmts = fmts or {}
    for j, c in enumerate(cols, 1):
        f = fmts.get(c)
        font = blue if c in inputs else base
        for i in range(2, n + 1):
            cell = ws.cell(i, j)
            cell.font = font
            if f:
                cell.number_format = f
        ws.cell(1, j).font = bold
        ws.column_dimensions[get_column_letter(j)].width = (widths or {}).get(c, max(10, min(40, len(c) + 4)))
    if notes:
        nc = len(cols) + 2
        ws.cell(1, nc, "Notes").font = bold
        for i, t_ in enumerate(notes, 2):
            ws.cell(i, nc, t_).font = muted
        ws.column_dimensions[get_column_letter(nc)].width = 90
    ws.freeze_panes = "A2"
    return ws


# ---------------- data ----------------
bal_c, bal = load("balance_monthly", {"MonthStart": "d", "Org": "s", "Subinventory": "s", "SubType": "s", "Balance": "n"})
io_c, io = load("inout_monthly", {"MonthStart": "d", "Org": "s", "Subinventory": "s", "SubType": "s", "TrxType": "s", "Direction": "s", "Value": "n", "Quantity": "n", "Items": "i"})
rq_c, rq = load("requisitions", {"MonthStart": "d", "CreationDate": "d", "Org": "s", "DestSub": "s", "SubType": "s", "ReqNumber": "s", "LineNum": "s", "ItemCategory": "s", "ItemID": "s", "Qty": "n", "QtyDelivered": "n", "UnitPrice": "n", "Amount": "n", "Status": "s", "Cancelled": "s", "Buyer": "s", "Supplier": "s"})
po_c0, po = load("open_pos", {"AsOfDate": "d", "PONumber": "s", "LineNum": "s", "Org": "s", "ShipTo": "s", "Supplier": "s", "Buyer": "s", "ItemID": "s", "Description": "s", "PromisedDate": "d", "NeedByDate": "d", "DaysPastDue": "i", "AgeBucket": "s", "QtyOutstanding": "n", "UnitPrice": "n", "OpenAmount": "n"})
po_c = po_c0 + ["FirstLine"]
for i, r in enumerate(po, 2):
    r.append(f"=IF(COUNTIFS($A$2:$A{i},$A{i},$B$2:$B{i},$B{i})=1,1,0)")
tu_c, tu = load("turns", {"PeriodEnd": "d", "Level": "s", "Name": "s", "Org": "s", "TotalUsage": "n", "BalanceSum": "n", "AverageBalance": "n", "Turns": "n"})
an_c, an = load("annual", {"Year": "i", "LocationCode": "s", "Basis": "s", "Turns": "n", "TotalUsage": "n", "AverageBalance": "n"})

last_month = max(r[0] for r in bal)
trx_types = sorted({r[4] for r in io})
USAGE = {"92 - WO / AC Issue": ("Issue", 1), "93 - Dept Issue": ("Issue", 1), "Move Order Issue": ("Issue", 1),
         "07- Dept Return": ("Return", 1), "08 - WO / AC Return": ("Return", 1), "PO Receipt": ("PO receipt", 0)}

def trx_group(t):
    if t in USAGE:
        return USAGE[t]
    tl = t.lower()
    if "transfer" in tl or "intransit" in tl or "intr" in tl or "pick" in tl:
        return ("Transfer", 0)
    if "adjust" in tl or "cycle" in tl:
        return ("Adjustment", 0)
    if "vendor" in tl or "supplier" in tl:
        return ("Return to supplier", 0)
    return ("Other", 0)

# ---------------- sheets ----------------
rm = wb.active; rm.title = "Read Me"
sm = wb.create_sheet("Summary")

table("Locations", "Locations", ["Org", "LocationName", "MainSub", "SurplusSub", "SortOrder"],
      [["LAS", "Las Vegas", "2111", "2111S", 1], ["HND", "Henderson", "2151", "2151S", 2], ["SNV", "Southern Nevada division", None, None, 3], ["CO", "Company", None, None, 4]],
      widths={"LocationName": 26})
table("Balance Monthly", "BalanceMonthly", bal_c, bal, {"MonthStart": DATE, "Balance": MONEY}, inputs=("Balance",),
      notes=["Source: OBIEE 'Inventory Balance - MONTH' (one row per subinventory per month).",
             "SubType: Main = 2111 / 2151, Surplus = 2111S / 2151S, Other = every other subinventory in the org.",
             "Add older or newer months by pasting rows at the bottom; the table grows by itself."])
table("In Out", "InOut", io_c, io, {"MonthStart": DATE, "Value": MONEY, "Quantity": INT, "Items": INT}, inputs=("Value", "Quantity", "Items"),
      widths={"TrxType": 34},
      notes=["Source: OBIEE 'In Out Summary', totalled by month, subinventory and transaction type.",
             "Out values are negative, as in OBIEE.",
             "The export had no date column, so it is stamped October 2026. Confirm the date range it covers."])
table("Trx Types", "TrxTypes", ["TrxType", "Group", "CountsAsUsage"], [[t, *trx_group(t)] for t in trx_types],
      widths={"TrxType": 36, "Group": 18}, inputs=("Group", "CountsAsUsage"),
      notes=["Usage = issues less returns, where CountsAsUsage = 1. Change the 1s and 0s to change what counts as usage.",
             "PO receipts = the 'PO Receipt' type."])
table("Requisitions", "Requisitions", rq_c, rq, {"MonthStart": DATE, "CreationDate": DATE, "Qty": INT, "QtyDelivered": INT, "UnitPrice": MONEY, "Amount": MONEY},
      inputs=("Qty", "UnitPrice"), widths={"ItemCategory": 22, "Supplier": 30, "Buyer": 22},
      notes=["Source: OBIEE 'Purchases for a Given Period - Requisition Order With Price'.",
             "Amount = Qty x UnitPrice. KPIs count APPROVED lines with Cancelled = N."])
table("Open POs", "OpenPOs", po_c, po, {"AsOfDate": DATE, "PromisedDate": DATE, "NeedByDate": DATE, "QtyOutstanding": INT, "UnitPrice": MONEY, "OpenAmount": MONEY, "DaysPastDue": INT},
      inputs=("QtyOutstanding", "UnitPrice"), widths={"Supplier": 32, "Description": 40, "ShipTo": 22},
      notes=["Source: OBIEE 'OPEN POs Past Promised Date', all locations. Org is matched from Ship To.",
             "Each upload is a snapshot (AsOfDate). Keep old snapshots to see the trend.",
             "OpenAmount = QtyOutstanding x UnitPrice, or the line amount for service lines with no quantity.",
             "FirstLine = 1 on the first line of each PO in a snapshot (used to count POs). Copy it down for pasted rows."])
table("Turns", "Turns", tu_c, tu, {"PeriodEnd": DATE, "TotalUsage": MONEY, "BalanceSum": MONEY, "AverageBalance": MONEY, "Turns": TURNS},
      inputs=("TotalUsage", "BalanceSum", "AverageBalance", "Turns"),
      notes=["Source: OBIEE 'Inventory TURNS' by Organization, by Division and Company.",
             "Year to date Jan to PeriodEnd. BalanceSum is the sum of month-end balances for those months.",
             "OBIEE's AverageBalance divides BalanceSum by 12, not by the number of months, so it reads low mid-year.",
             "Turns = usage / (BalanceSum / 12), which equals annualised turns."])
table("Annual", "Annual", an_c, an, {"TotalUsage": MONEY, "AverageBalance": MONEY, "Turns": TURNS}, inputs=("Turns", "TotalUsage", "AverageBalance"),
      notes=["Source: 'KPIs - Nov Update 2026.xlsx'. Full-year turns history; 2026 is year to date."])

ca_rows = []
m = dt.date(2018, 1, 1)
while m <= dt.date(2027, 12, 1):
    i = len(ca_rows) + 2
    ca_rows.append([m, f"=YEAR(A{i})", f"=MONTH(A{i})", f'=TEXT(A{i},"mmm")', f'="Q"&ROUNDUP(C{i}/3,0)', f'=B{i}&" "&E{i}',
                    f"=DATE(B{i},(ROUNDUP(C{i}/3,0)-1)*3+1,1)", f'=TEXT(A{i},"yyyy-mm")'])
    m = dt.date(m.year + (m.month == 12), m.month % 12 + 1, 1)
table("Calendar", "Calendar", ["MonthStart", "Year", "MonthNumber", "MonthName", "Quarter", "YearQuarter", "QuarterStart", "YearMonth"], ca_rows,
      {"MonthStart": DATE, "QuarterStart": DATE})

# ---------------- Summary (live formulas) ----------------
sm["A1"] = "Inventory KPIs - Las Vegas and Henderson"; sm["A1"].font = h1
sm["A3"] = "Month"; sm["A3"].font = bold
sm["B3"] = last_month; sm["B3"].number_format = "mmm yyyy"; sm["B3"].fill = yellow; sm["B3"].font = blue
sm["B3"].comment = Comment("Type the first day of any month (e.g. 2026-09-01).", "Claude")
sm["C3"] = "Change this to compare another month."; sm["C3"].font = muted
sm["A4"] = "Quarter start"; sm["A4"].font = bold
sm["B4"] = "=DATE(YEAR(B3),(ROUNDUP(MONTH(B3)/3,0)-1)*3+1,1)"; sm["B4"].number_format = "mmm yyyy"
sm["A5"] = "Months in quarter to date"; sm["A5"].font = bold
sm["B5"] = "=MONTH(B3)-MONTH(B4)+1"
sm["A6"] = "Months in year to date"; sm["A6"].font = bold
sm["B6"] = "=MONTH(B3)"
for c in ("B4", "B5", "B6"):
    sm[c].font = base

hdr = ["Measure", "Location", "This month", "Last month", "MoM", "Same month LY", "YoY",
       "Qtr to date", "Same months last qtr", "QoQ", "Same qtr LY", "Qtr YoY", "Year to date", "Same months LY", "YTD YoY"]
R0 = 8
for j, h in enumerate(hdr, 1):
    c = sm.cell(R0, j, h); c.font = hfont; c.fill = head; c.alignment = wrap
sm.row_dimensions[R0].height = 30


SHEETS = {"BalanceMonthly": ("Balance Monthly", bal_c), "InOut": ("In Out", io_c), "Requisitions": ("Requisitions", rq_c),
          "OpenPOs": ("Open POs", po_c), "TrxTypes": ("Trx Types", ["TrxType", "Group", "CountsAsUsage"])}
MAXR = 50000  # ranges reach well past today's rows, so pasted rows are picked up


def rg(tbl, col):
    sheet, cols = SHEETS[tbl]
    L = get_column_letter(cols.index(col) + 1)
    return f"'{sheet}'!${L}$2:${L}${MAXR}"


def period_sum(kind, sub, loc, start, end):
    """Total of a measure for one org between two month starts (inclusive)."""
    if kind == "balance":
        crit = f'{rg("BalanceMonthly","MonthStart")},">="&{start},{rg("BalanceMonthly","MonthStart")},"<="&{end},{rg("BalanceMonthly","Org")},{loc}'
        if sub:
            crit += f',{rg("BalanceMonthly","SubType")},"{sub}"'
        return f'SUMIFS({rg("BalanceMonthly","Balance")},{crit})'
    m, o, v = rg("InOut", "MonthStart"), rg("InOut", "Org"), rg("InOut", "Value")
    if kind == "usage":
        return (f'-SUMPRODUCT({v}*({m}>={start})*({m}<={end})*({o}={loc}),'
                f'SUMIFS({rg("TrxTypes","CountsAsUsage")},{rg("TrxTypes","TrxType")},{rg("InOut","TrxType")}))')
    crit = f'{m},">="&{start},{m},"<="&{end},{o},{loc}'
    if kind == "receipts":
        return f'SUMIFS({v},{crit},{rg("InOut","TrxType")},"PO Receipt",{rg("InOut","Direction")},"In")'
    if kind == "in":
        return f'SUMIFS({v},{crit},{rg("InOut","Direction")},"In")'
    if kind == "out":
        return f'-SUMIFS({v},{crit},{rg("InOut","Direction")},"Out")'
    if kind == "req":
        rm_, ro = rg("Requisitions", "MonthStart"), rg("Requisitions", "Org")
        return (f'SUMIFS({rg("Requisitions","Amount")},{rm_},">="&{start},{rm_},"<="&{end},{ro},{loc},'
                f'{rg("Requisitions","Status")},"APPROVED",{rg("Requisitions","Cancelled")},"N")')
    raise ValueError(kind)


def has_data(kind, loc, start, end, months):
    """True when the first and last month of the period are loaded for this org, so comparisons stay blank until history exists."""
    tbl = {"balance": "BalanceMonthly", "req": "Requisitions"}.get(kind, "InOut")
    m, o = rg(tbl, "MonthStart"), rg(tbl, "Org")
    return f"AND(COUNTIFS({m},{start},{o},{loc})>0,COUNTIFS({m},{end},{o},{loc})>0)"


MEASURES = [("Inventory balance", "balance", None, True), ("Main stores balance", "balance", "Main", True),
            ("Surplus balance", "balance", "Surplus", True), ("Usage (net issues)", "usage", None, False),
            ("PO receipts $", "receipts", None, False), ("Value in", "in", None, False), ("Value out", "out", None, False),
            ("Requisitions $", "req", None, False)]
VALCOLS = (3, 4, 6, 8, 9, 11, 13, 14)
row = R0 + 1
for label, kind, sub, is_avg in MEASURES:
    for loc in ("LAS", "HND", "Both"):
        L = f"$B{row}"
        sm.cell(row, 1, label).font = bold if loc == "LAS" else base
        sm.cell(row, 2, loc).font = base
        B3, B4, B5, B6 = "$B$3", "$B$4", "$B$5", "$B$6"
        ys = "DATE(YEAR($B$3),1,1)"
        periods = {3: (B3, B3, 1), 4: (f"EDATE({B3},-1)", f"EDATE({B3},-1)", 1), 6: (f"EDATE({B3},-12)", f"EDATE({B3},-12)", 1),
                   8: (B4, B3, B5), 9: (f"EDATE({B4},-3)", f"EDATE({B3},-3)", B5), 11: (f"EDATE({B4},-12)", f"EDATE({B3},-12)", B5),
                   13: (ys, B3, B6), 14: (f"EDATE({ys},-12)", f"EDATE({B3},-12)", B6)}
        for col, (st, en, n) in periods.items():
            if loc == "Both":
                a, b = get_column_letter(col) + str(row - 2), get_column_letter(col) + str(row - 1)
                sm.cell(row, col, f'=IF(OR({a}="",{b}=""),"",{a}+{b})')
            else:
                s_ = period_sum(kind, sub, L, st, en)
                v = f"{s_}/{n}" if is_avg else s_
                sm.cell(row, col, f'=IF({has_data(kind, L, st, en, n)},{v},"")')
        for pc, a, b in ((5, "C", "D"), (7, "C", "F"), (10, "H", "I"), (12, "H", "K"), (15, "M", "N")):
            sm.cell(row, pc, f'=IF(OR({a}{row}="",{b}{row}=""),"",IF(N({b}{row})=0,"",{a}{row}/{b}{row}-1))')
        for j in range(3, 16):
            c = sm.cell(row, j); c.font = bold if loc == "Both" else base
            c.number_format = PCT if j in (5, 7, 10, 12, 15) else MONEY
        row += 1

row += 1
sm.cell(row, 1, "Overdue POs (latest upload)").font = h2; row += 1
for j, h in enumerate(["Location", "As of", "Lines", "POs", "Open $", "90+ days lines", "90+ days $"], 1):
    c = sm.cell(row, j, h); c.font = hfont; c.fill = head
row += 1
pa, po_o, pn, pam, pb = rg("OpenPOs", "AsOfDate"), rg("OpenPOs", "Org"), rg("OpenPOs", "PONumber"), rg("OpenPOs", "OpenAmount"), rg("OpenPOs", "AgeBucket")
for loc in ("LAS", "HND", "Both", "Company"):
    r = row
    sm.cell(r, 1, loc).font = bold if loc in ("Both", "Company") else base
    sm.cell(r, 2, f"=MAX({pa})").number_format = DATE
    if loc == "Both":
        for j in (3, 5, 6, 7):
            L_ = get_column_letter(j)
            sm.cell(r, j, f"={L_}{r-2}+{L_}{r-1}")
        sm.cell(r, 4, f"=D{r-2}+D{r-1}")
    else:
        oc = "" if loc == "Company" else f',{po_o},"{loc}"'
        om = "1" if loc == "Company" else f'({po_o}="{loc}")'
        sm.cell(r, 3, f"=COUNTIFS({pa},$B{r}{oc})")
        sm.cell(r, 4, f"=SUMIFS({rg('OpenPOs','FirstLine')},{pa},$B{r}{oc})")
        sm.cell(r, 5, f"=SUMIFS({pam},{pa},$B{r}{oc})")
        sm.cell(r, 6, f'=COUNTIFS({pa},$B{r},{pb},"90+"{oc})')
        sm.cell(r, 7, f'=SUMIFS({pam},{pa},$B{r},{pb},"90+"{oc})')
    for j, f in zip(range(2, 8), [DATE, INT, INT, MONEY, INT, MONEY]):
        sm.cell(r, j).number_format = f; sm.cell(r, j).font = bold if loc in ("Both", "Company") else base
    row += 1

row += 1
sm.cell(row, 1, "Turns (year to date, annualised)").font = h2; row += 1
for j, h in enumerate(["Level", "Period end", "Turns", "Last year (full)", "Change", "Usage YTD", "True avg balance"], 1):
    c = sm.cell(row, j, h); c.font = hfont; c.fill = head
row += 1
for org, label in (("LAS", "Las Vegas"), ("HND", "Henderson"), ("SNV", "Southern Nevada"), ("CO", "Company")):
    r = row
    sm.cell(r, 1, label).font = base; sm.cell(r, 9, org).font = muted
    pe = f"_xlfn.MAXIFS(Turns!$A$2:$A$500,Turns!$D$2:$D$500,$I{r})"
    sm.cell(r, 2, f"={pe}").number_format = "mmm yyyy"
    sm.cell(r, 3, f"=SUMIFS(Turns!$H$2:$H$500,Turns!$D$2:$D$500,$I{r},Turns!$A$2:$A$500,$B{r})").number_format = TURNS
    sm.cell(r, 4, f'=IFERROR(1/(1/SUMIFS(Annual!$D$2:$D$500,Annual!$B$2:$B$500,$I{r},Annual!$A$2:$A$500,YEAR($B{r})-1,Annual!$C$2:$C$500,"Jan-Dec")),"")').number_format = TURNS
    sm.cell(r, 5, f'=IF(D{r}="","",C{r}-D{r})').number_format = '+0.00"x";-0.00"x";0.00"x"'
    sm.cell(r, 6, f"=SUMIFS(Turns!$E$2:$E$500,Turns!$D$2:$D$500,$I{r},Turns!$A$2:$A$500,$B{r})").number_format = MONEY
    sm.cell(r, 7, f"=SUMIFS(Turns!$F$2:$F$500,Turns!$D$2:$D$500,$I{r},Turns!$A$2:$A$500,$B{r})/MONTH($B{r})").number_format = MONEY
    for j in range(2, 8):
        sm.cell(r, j).font = base
    row += 1
row += 1
for t in ["Balances: quarter and year columns are averages of month-end balances. Flows (usage, receipts, in, out, requisitions) are totals.",
          "Quarter to date is compared with the same months of the last quarter and of the same quarter last year.",
          "A blank means the earlier months are not loaded yet. Lower balance and surplus are better; higher turns are better.",
          "Column I of the Turns block holds the codes the formulas look up."]:
    sm.cell(row, 1, t).font = muted; row += 1
for col, w in zip("ABCDEFGHIJKLMNO", [22, 10] + [14] * 13):
    sm.column_dimensions[col].width = w
sm.freeze_panes = sm.cell(R0 + 1, 3)

# ---------------- KPI Measures (DAX) ----------------
km = wb.create_sheet("KPI Measures")
km.append(["Measure", "What it means", "Better when", "DAX (Power BI: Modeling > New measure)"])
B = "AVERAGEX ( VALUES ( 'Calendar'[MonthStart] ), CALCULATE ( SUM ( BalanceMonthly[Balance] ) ) )"
dax = [
    ("Balance", "Month-end inventory balance; over several months, the average month-end balance.", "Lower", f"Balance = {B}"),
    ("Surplus Balance", "Balance in surplus subinventories (2111S, 2151S).", "Lower", "Surplus Balance = CALCULATE ( [Balance], BalanceMonthly[SubType] = \"Surplus\" )"),
    ("Main Balance", "Balance in main stores (2111, 2151).", "Lower", "Main Balance = CALCULATE ( [Balance], BalanceMonthly[SubType] = \"Main\" )"),
    ("Usage", "Issues less returns, for types with CountsAsUsage = 1 in TrxTypes.", "-", "Usage = - SUMX ( InOut, InOut[Value] * RELATED ( TrxTypes[CountsAsUsage] ) )"),
    ("PO Receipts", "Value received on POs.", "-", "PO Receipts = CALCULATE ( SUM ( InOut[Value] ), InOut[TrxType] = \"PO Receipt\", InOut[Direction] = \"In\" )"),
    ("Items Received", "Items received on POs.", "-", "Items Received = CALCULATE ( SUM ( InOut[Items] ), InOut[TrxType] = \"PO Receipt\", InOut[Direction] = \"In\" )"),
    ("Value In", "All value in.", "-", "Value In = CALCULATE ( SUM ( InOut[Value] ), InOut[Direction] = \"In\" )"),
    ("Value Out", "All value out (shown positive).", "-", "Value Out = - CALCULATE ( SUM ( InOut[Value] ), InOut[Direction] = \"Out\" )"),
    ("Net In Out", "Value in minus value out.", "-", "Net In Out = [Value In] - [Value Out]"),
    ("Requisitions $", "Approved, uncancelled requisition dollars.", "-", "Requisitions $ = CALCULATE ( SUM ( Requisitions[Amount] ), Requisitions[Status] = \"APPROVED\", Requisitions[Cancelled] = \"N\" )"),
    ("Overdue PO Lines", "PO lines past promised date on the latest upload.", "Lower", "Overdue PO Lines = VAR d = MAX ( OpenPOs[AsOfDate] ) RETURN CALCULATE ( COUNTROWS ( OpenPOs ), OpenPOs[AsOfDate] = d )"),
    ("Overdue PO $", "Open dollars past promised date on the latest upload.", "Lower", "Overdue PO $ = VAR d = MAX ( OpenPOs[AsOfDate] ) RETURN CALCULATE ( SUM ( OpenPOs[OpenAmount] ), OpenPOs[AsOfDate] = d )"),
    ("Turns", "Year-to-date annualised turns from OBIEE.", "Higher", "Turns = VAR d = MAX ( Turns[PeriodEnd] ) RETURN CALCULATE ( AVERAGE ( Turns[Turns] ), Turns[PeriodEnd] = d )"),
    ("<any> PM", "Template: previous month. Swap [Balance] for any measure.", "-", "Balance PM = CALCULATE ( [Balance], DATEADD ( 'Calendar'[MonthStart], -1, MONTH ) )"),
    ("<any> MoM %", "Change vs previous month.", "-", "Balance MoM % = DIVIDE ( [Balance] - [Balance PM], [Balance PM] )"),
    ("<any> PY", "Same month last year.", "-", "Balance PY = CALCULATE ( [Balance], SAMEPERIODLASTYEAR ( 'Calendar'[MonthStart] ) )"),
    ("<any> YoY %", "Change vs same month last year.", "-", "Balance YoY % = DIVIDE ( [Balance] - [Balance PY], [Balance PY] )"),
    ("<any> QTD", "Quarter to date (average for balances, total for flows).", "-", "Balance QTD = CALCULATE ( [Balance], DATESQTD ( 'Calendar'[MonthStart] ) )"),
    ("<any> PQTD", "Same months of the previous quarter.", "-", "Balance PQTD = CALCULATE ( [Balance], DATEADD ( DATESQTD ( 'Calendar'[MonthStart] ), -1, QUARTER ) )"),
    ("<any> QoQ %", "Quarter to date vs same months last quarter.", "-", "Balance QoQ % = DIVIDE ( [Balance QTD] - [Balance PQTD], [Balance PQTD] )"),
    ("<any> QTD PY", "Same quarter last year.", "-", "Balance QTD PY = CALCULATE ( [Balance QTD], SAMEPERIODLASTYEAR ( 'Calendar'[MonthStart] ) )"),
    ("<any> YTD", "Year to date.", "-", "Balance YTD = CALCULATE ( [Balance], DATESYTD ( 'Calendar'[MonthStart] ) )"),
    ("<any> YTD PY", "Same months last year.", "-", "Balance YTD PY = CALCULATE ( [Balance YTD], SAMEPERIODLASTYEAR ( 'Calendar'[MonthStart] ) )"),
    ("<any> YTD YoY %", "Year to date vs same months last year.", "-", "Balance YTD YoY % = DIVIDE ( [Balance YTD] - [Balance YTD PY], [Balance YTD PY] )"),
]
for r in dax:
    km.append(list(r))
for rr in km.iter_rows():
    for c in rr:
        c.font = base; c.alignment = wrap
for c in km[1]:
    c.font = hfont; c.fill = head
for col, w in zip("ABCD", [20, 50, 12, 120]):
    km.column_dimensions[col].width = w
km.freeze_panes = "A2"

# ---------------- Read Me ----------------
lines = [
    ("Inventory KPIs - Power BI workbook", h1), ("", base),
    ("What is in this file", h2),
    ("Summary - live Excel view of every KPI for Las Vegas, Henderson and both: month, MoM, YoY, quarter to date, QoQ, quarter YoY, year to date. Change the yellow month.", base),
    ("Balance Monthly (BalanceMonthly) - month-end balance per subinventory, Jan 2018 onward.", base),
    ("In Out (InOut) - value in and out by month, subinventory and transaction type.", base),
    ("Trx Types (TrxTypes) - which transaction types count as usage. Edit the 1s and 0s.", base),
    ("Requisitions - requisition lines with amount.", base),
    ("Open POs (OpenPOs) - PO lines past promised date, one snapshot per upload date.", base),
    ("Turns - OBIEE year-to-date turns by org, division and company. Annual - full-year turns history.", base),
    ("Locations, Calendar - lookup tables for Power BI.", base),
    ("KPI Measures - DAX for every measure, plus MoM, YoY, QoQ and YTD templates.", base),
    ("", base), ("Adding more data (older or newer months)", h2),
    ("Easiest: send the same OBIEE exports to Claude for any date range. Claude runs pipeline/load_obiee.py and pipeline/build_workbook.py, which merge the new months and rebuild this file and the dashboard. A month sent again replaces the old one.", base),
    ("By hand: paste rows at the bottom of the matching table, using the same columns. Blue text = data from OBIEE.", base),
    ("Tip: add a Month (or Period) column to the In Out Summary and Turns reports in OBIEE, and export several months at once. Then no date has to be assumed.", base),
    ("", base), ("Connect to Power BI", h2),
    ("1. Save this file to OneDrive or SharePoint so Power BI Service can refresh it.", base),
    ("2. Power BI Desktop > Get data > Excel workbook > tick the tables (not the sheets): BalanceMonthly, InOut, TrxTypes, Requisitions, OpenPOs, Turns, Annual, Locations, Calendar.", base),
    ("3. Relationships: Calendar[MonthStart] 1-to-many to BalanceMonthly, InOut and Requisitions [MonthStart]; Locations[Org] 1-to-many to the Org column of each table; TrxTypes[TrxType] 1-to-many to InOut[TrxType].", base),
    ("4. Mark Calendar as a date table on MonthStart.", base),
    ("5. Add the measures from KPI Measures.", base),
    ("", base), ("Notes", h2),
    ("The In Out export had no date column. It is stamped October 2026; its totals look larger than one month, so confirm the date range.", base),
    ("Turns are year to date (Jan to Oct 2026) and annualised. OBIEE's Average Balance column divides by 12, so the true average is higher.", base),
    ("Overdue POs come from a company-wide report; Org is matched from Ship To (2111 Las Vegas, 2151 Henderson, else Other).", base),
]
for i, (t, f) in enumerate(lines, 1):
    c = rm.cell(i, 1, t); c.font = f; c.alignment = Alignment(wrap_text=True)
rm.column_dimensions["A"].width = 140

wb.move_sheet("Summary", offset=1 - wb.sheetnames.index("Summary"))
wb.active = 0
wb.save(OUT)
print("saved", OUT, [ws.title for ws in wb.worksheets])
