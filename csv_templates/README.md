# CSV templates for the Inventory KPIs dashboard and workbook

Fill one row per month per location. Keep the header row exactly as it is.

- Dates: YYYY-MM-DD. MonthStart is the first day of the month (2026-10-01).
- Subinventory: 2111 (Las Vegas), 2151 (Henderson). Surplus uses 2111S and 2151S.
- Money: plain numbers, no $ or commas (1283483.92).
- History: as far back as you have, ideally Jan 2018 to now (YoY needs at least 13 months).

| File | Feeds | Example row |
|---|---|---|
| 1_receipts_monthly.csv | Receipts MoM, QoQ, YoY | 2026-10-01,2111,412,1180,925000.00 |
| 2_overdue_pos.csv | Overdue POs (snapshot on upload day) | 2026-11-05,PO123456,1,Acme Supply,J Smith,2111,2026-09-15,2026-09-30,40,5200.00 |
| 3_inventory_monthly.csv | Average balance, usage, turns, in/out | 2026-10-01,2111,4100000,4050000,610000,590000,640000 |
| 4_surplus_monthly.csv | Surplus | 2026-10-31,2111S,1283483.92,2026-10-31 |

Notes
- Overdue POs: send every open PO line, not only overdue ones. I flag a line as late when its promised date (or need-by date if blank) is before AsOfDate.
- Turns: rolling 12-month usage divided by the 12-month average of AverageBalance.
- In/out: ValueIn is the dollar value received; ValueOut is the dollar value issued. Net change = ValueIn - ValueOut.
- Optional columns (ReceiptLines, ReceiptValue, Buyer, EndingBalance) can be left blank.
