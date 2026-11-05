"""Load raw OBIEE exports into the tidy CSVs that feed the dashboard and the Power BI workbook.

Usage:
    python pipeline/load_obiee.py <export files...> [--month 2026-10] [--data data]

Each export is recognised by its columns, not its file name. New rows are merged into the CSVs
in --data: a month (or snapshot date) that is loaded again replaces the old rows for it, so
re-sending an export is safe and older history can be added at any time.

Exports without a date column (In/Out Summary, Inventory Turns) are stamped with --month,
which defaults to the month before today. Adding a Month column to those OBIEE reports
removes the need for --month.
"""
import argparse
import datetime as dt
import re
import sys
from pathlib import Path

import pandas as pd

ORG = {"Las Vegas": "LAS", "Henderson": "HND"}
MAIN_SUB = {"LAS": "2111", "HND": "2151"}


def org_from_text(s):
    s = str(s)
    if "2111" in s or "Las Vegas" in s:
        return "LAS"
    if "2151" in s or "Henderson" in s or "HND" in s:
        return "HND"
    return "Other"


def sub_type(org, sub):
    sub = str(sub)
    base = MAIN_SUB.get(org)
    if base is None or not sub.startswith(base):
        return "Other"
    suffix = sub[len(base):]
    return {"": "Main", "S": "Surplus"}.get(suffix, "Other")


def read_any(path):
    p = Path(path)
    if p.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(p, header=None)
    return pd.read_csv(p, header=None, encoding="utf-8-sig", dtype=str, keep_default_na=False)


def promote_header(raw, must_have):
    """Find the header row (OBIEE exports can have title rows above it)."""
    for i in range(min(10, len(raw))):
        row = [str(v).strip() for v in raw.iloc[i].tolist()]
        if all(any(m.lower() == c.lower() for c in row) for m in must_have):
            df = raw.iloc[i + 1:].copy()
            df.columns = row
            return df.reset_index(drop=True)
    return None


def num(s):
    return pd.to_numeric(pd.Series(s).astype(str).str.replace(r"[$,]", "", regex=True), errors="coerce")


def drop_footer(df, first_col):
    """Remove OBIEE footer lines like 'TOTAL RECORD COUNT' and the confidentiality notice."""
    c = df[first_col].astype(str)
    return df[~c.str.contains("TOTAL RECORD COUNT|CONFIDENTIAL", case=False, na=False)]


# ---------------- parsers ----------------

def parse_balance(raw):
    """Pivot: Division | Organization | Subinventory | Avg | one column per month."""
    hdr_i = next((i for i in range(6) if "Subinventory" in [str(v) for v in raw.iloc[i]]), None)
    if hdr_i is None:
        return None
    hdr = raw.iloc[hdr_i]
    month_cols = [c for c in raw.columns if re.fullmatch(r"[A-Z]{3}-\d{2}", str(hdr[c]).strip())]
    if not month_cols:
        return None
    body = raw.iloc[hdr_i + 1:].copy()
    body[raw.columns[1]] = body[raw.columns[1]].ffill()
    body = body[body[raw.columns[2]].notna() & (body[raw.columns[2]].astype(str).str.strip() != "")]
    out = []
    for org_name, g in body.groupby(raw.columns[1]):
        org = ORG.get(str(org_name).strip())
        if not org:
            continue
        # OBIEE gives each organisation its own set of period columns (sometimes out of order).
        # Each column's header label is its real month, so map by label and check for repeats.
        used = [c for c in month_cols if g[c].notna().any() and (g[c].astype(str).str.strip() != "").any()]
        months = pd.to_datetime([str(hdr[c]).strip() for c in used], format="%b-%y")
        if months.duplicated().any():
            print(f"  warning: {org_name} has repeated month columns: {sorted(set(months[months.duplicated()].strftime('%b-%y')))}")
        for c, m in zip(used, months):
            for _, r in g.iterrows():
                v = pd.to_numeric(r[c], errors="coerce")
                if pd.notna(v):
                    sub = str(r[raw.columns[2]]).strip()
                    out.append((m.strftime("%Y-%m-%d"), org, sub, sub_type(org, sub), round(float(v), 2)))
    df = pd.DataFrame(out, columns=["MonthStart", "Org", "Subinventory", "SubType", "Balance"])
    return "balance_monthly", df, ["MonthStart", "Org"]


def parse_inout(df, month):
    df = drop_footer(df, df.columns[0])
    df = df[df["IN or OUT"].astype(str).str.strip() != ""].copy()
    df["Org"] = df["Organization Name"].map(lambda s: ORG.get(str(s).strip(), "Other"))
    df["Direction"] = df["IN or OUT"].str.replace("Total ", "", regex=False).str.strip()
    df["Value"] = num(df["Value"]).values
    df["Quantity"] = num(df["Quantity"]).values
    if "Month" in df.columns:
        df["MonthStart"] = pd.to_datetime(df["Month"]).dt.to_period("M").dt.start_time.dt.strftime("%Y-%m-%d")
    else:
        df["MonthStart"] = month
    g = (df.groupby(["MonthStart", "Org", "Subinventory", "TRX Type", "Direction"], as_index=False)
           .agg(Value=("Value", "sum"), Quantity=("Quantity", "sum"), Items=("Item ID", "nunique")))
    g = g.rename(columns={"TRX Type": "TrxType"})
    g["SubType"] = [sub_type(o, s) for o, s in zip(g.Org, g.Subinventory)]
    g["Value"] = g["Value"].round(2)
    return "inout_monthly", g, ["MonthStart", "Org"]


def parse_turns(df, month, level):
    df = drop_footer(df, df.columns[0])
    name_col = df.columns[0]
    for c in ["Total Usage by Org", "Total Balance by Org", "Turns", "Average Balance"]:
        df[c] = num(df[c]).values
    g = df.groupby(name_col, as_index=False).agg(
        TotalUsage=("Total Usage by Org", "sum"), BalanceSum=("Total Balance by Org", "sum"),
        AverageBalance=("Average Balance", "max"), Turns=("Turns", "max"))
    g = g.rename(columns={name_col: "Name"})
    g["Name"] = g["Name"].astype(str).str.strip()
    g["Level"] = level
    g["Org"] = g["Name"].map(lambda s: ORG.get(s, "CO" if level == "Company" else s))
    g["PeriodEnd"] = month
    return "turns", g[["PeriodEnd", "Level", "Name", "Org", "TotalUsage", "BalanceSum", "AverageBalance", "Turns"]], ["PeriodEnd", "Level"]


def parse_open_pos(df):
    df = drop_footer(df, df.columns[0])
    df = df[df["PO Number"].astype(str).str.strip() != ""].copy()
    prom = pd.to_datetime(df["Promised Date"], errors="coerce")
    need = pd.to_datetime(df["Need By Date"], errors="coerce")
    due = prom.fillna(need)
    days = num(df["Days Past Due"]).values
    asof = (due.dt.normalize() + pd.to_timedelta(days, unit="D")).mode()
    asof = asof.iloc[0] if len(asof) else pd.Timestamp.today().normalize()
    qty = num(df["Quantity"]).values
    qty_out = num(df["Quantity Outstanding"]).values
    price = num(df["Unit Price"]).fillna(0).values
    amount = num(df["Amount"]).fillna(0).values
    # Quantity lines: outstanding qty x price. Amount-based (service) lines have no quantity; use the line amount.
    open_amt = [a if (q == 0 or pd.isna(q)) and a > 0 else (qo or 0) * p for q, qo, p, a in zip(qty, qty_out, price, amount)]
    out = pd.DataFrame({
        "AsOfDate": asof.strftime("%Y-%m-%d"),
        "PONumber": df["PO Number"].astype(str).str.strip(),
        "LineNum": df["PO Line Number"].astype(str).str.strip(),
        "Supplier": df["Supplier Name"].str.strip(),
        "Buyer": df["Buyer Name"].str.strip(),
        "ShipTo": df["Ship To"].str.strip(),
        "Org": df["Ship To"].map(org_from_text),
        "ItemID": df["Item ID"].astype(str).str.strip(),
        "Description": df["Item Description"].astype(str).str.slice(0, 80),
        "PromisedDate": prom.dt.strftime("%Y-%m-%d"),
        "NeedByDate": need.dt.strftime("%Y-%m-%d"),
        "DaysPastDue": days,
        "QtyOutstanding": qty_out,
        "UnitPrice": price,
        "OpenAmount": pd.Series(open_amt).round(2).values,
    })
    out["AgeBucket"] = pd.cut(out["DaysPastDue"], [-1, 30, 60, 90, 10**6], labels=["0-30", "31-60", "61-90", "90+"]).astype(str)
    return "open_pos", out, ["AsOfDate"]


def parse_reqs(df):
    for c in ["Operating Unit", "Dest Organization", "Destination Sub"]:
        df[c] = df[c].replace("", pd.NA).ffill()
    df = drop_footer(df, "Operating Unit")
    df = df[pd.to_datetime(df["Creation Date"], errors="coerce").notna()].copy()
    created = pd.to_datetime(df["Creation Date"])
    qty = num(df["Total QTY"]).values
    price = num(df["Unit_Price"]).values
    out = pd.DataFrame({
        "MonthStart": created.dt.to_period("M").dt.start_time.dt.strftime("%Y-%m-%d"),
        "CreationDate": created.dt.strftime("%Y-%m-%d"),
        "Org": df["Dest Organization"].map(lambda s: ORG.get(str(s).strip(), "Other")),
        "DestSub": df["Destination Sub"].astype(str).str.replace(r"\.0$", "", regex=True),
        "ReqNumber": df["Requisition Number"].astype(str).str.replace(r"\.0$", "", regex=True),
        "LineNum": df["Line Number"].astype(str).str.replace(r"\.0$", "", regex=True),
        "ItemCategory": df["Item Category"].replace("", pd.NA).ffill().str.replace("PO.", "", regex=False),
        "ItemID": df["Item ID"].astype(str),
        "Qty": qty,
        "QtyDelivered": num(df["QTY Delivered"]).values,
        "UnitPrice": price,
        "Amount": (pd.Series(qty) * pd.Series(price)).round(2).values,
        "Status": df["Authorization Status"],
        "Cancelled": df["Cancel Flag"],
        "Buyer": df["Buyer"].replace("", pd.NA),
        "Supplier": df["Supplier"].replace("", pd.NA),
    })
    out["SubType"] = [sub_type(o, s) for o, s in zip(out.Org, out.DestSub)]
    return "requisitions", out, ["MonthStart"]


def parse_file(path, month):
    raw = read_any(path)
    flat = raw.where(raw.notna(), "").astype(str)
    text = " ".join(flat.head(8).values.ravel())
    if "Subinventory" in text and re.search(r"\b[A-Z]{3}-\d{2}\b", text):
        return parse_balance(raw)
    for must, fn in [
        (["PO Number", "Days Past Due", "Quantity Outstanding"], lambda d: parse_open_pos(d)),
        (["Requisition Number", "Creation Date", "Unit_Price"], lambda d: parse_reqs(d)),
        (["IN or OUT", "TRX Type", "Value"], lambda d: parse_inout(d, month)),
        (["Organization Name", "Turns", "Total Usage by Org"], lambda d: parse_turns(d, month, "Organization")),
        (["Division Name", "Turns", "Total Usage by Org"], None),
    ]:
        df = promote_header(flat, must)
        if df is not None:
            if fn is None:
                names = set(df["Division Name"].str.strip()) - {""}
                level = "Company" if names <= {"SWGAS"} else "Division"
                return parse_turns(df, month, level)
            return fn(df)
    return None


def merge(data_dir, name, new, keys):
    path = data_dir / f"{name}.csv"
    if path.exists():
        old = pd.read_csv(path, dtype=str)
        new_keys = set(map(tuple, new[keys].astype(str).values))
        old = old[[tuple(r) not in new_keys for r in old[keys].astype(str).values]]
        new = pd.concat([old, new.astype(object).where(new.notna(), "").astype(str)], ignore_index=True)
    sort = [c for c in ["MonthStart", "PeriodEnd", "AsOfDate", "Org", "Level", "Subinventory", "PONumber"] if c in new.columns]
    new = new.sort_values(sort, kind="stable")
    new.to_csv(path, index=False)
    return len(new)


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("--month", help="YYYY-MM for exports with no date column (default: last month)")
    ap.add_argument("--data", default=str(Path(__file__).resolve().parent.parent / "data"))
    a = ap.parse_args(argv)
    if a.month:
        month = f"{a.month}-01"
    else:
        t = dt.date.today().replace(day=1) - dt.timedelta(days=1)
        month = t.replace(day=1).isoformat()
    data_dir = Path(a.data)
    data_dir.mkdir(parents=True, exist_ok=True)
    for f in a.files:
        res = parse_file(f, month)
        if res is None:
            print(f"skipped (not recognised): {Path(f).name}")
            continue
        name, df, keys = res
        total = merge(data_dir, name, df, keys)
        print(f"{Path(f).name} -> {name}.csv: {len(df)} rows loaded, {total} rows total")


if __name__ == "__main__":
    sys.exit(main())
