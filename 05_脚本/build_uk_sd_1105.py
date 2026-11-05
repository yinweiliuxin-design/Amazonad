import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
H = ["Product","Entity","Operation","Campaign ID","Portfolio ID","Ad Group ID","Ad ID","Targeting ID","Campaign Name","Ad Group Name",
     "Start Date","End Date","State","Tactic","Budget Type","Budget","SKU","Ad Group Default Bid","Bid","Bid Optimization","Cost Type","Targeting Expression"]
SD = "Sponsored Display"
def row(**k):
    r = {h: None for h in H}; r["Product"] = SD
    m = {"entity":"Entity","op":"Operation","cid":"Campaign ID","agid":"Ad Group ID","cname":"Campaign Name","agname":"Ad Group Name",
         "start":"Start Date","state":"State","tactic":"Tactic","btype":"Budget Type","budget":"Budget","sku":"SKU","agbid":"Ad Group Default Bid",
         "bid":"Bid","bidopt":"Bid Optimization","cost":"Cost Type","expr":"Targeting Expression"}
    for a, v in k.items(): r[m[a]] = v
    return r
rows = []
C, G = "SD-UK-浏览再营销", "看过没买-30天"
rows.append(row(entity="Campaign", op="Create", cid=C, cname=C, start="20261105", state="enabled", tactic="T00030", btype="daily", budget=3, cost="cpc"))
rows.append(row(entity="Ad Group", op="Create", cid=C, agid=G, agname=G, state="enabled", agbid=0.40, bidopt="conversions"))
for sku in ["HEI-CSQ-4/4", "V1-DS9X-ZY88", "02-H1HJ-R28I"]:   # B0H2ZQTWL3, B0H4WFM9BQ, B0H686FV6Q
    rows.append(row(entity="Product Ad", op="Create", cid=C, agid=G, state="enabled", sku=sku))
rows.append(row(entity="Audience Targeting", op="Create", cid=C, agid=G, state="enabled", bid=0.40, expr="views=(exact-product lookback=30)"))
rows.append(row(entity="Audience Targeting", op="Create", cid=C, agid=G, state="enabled", bid=0.35, expr="views=(similar-product lookback=30)"))
TEXT = {"Campaign ID","Ad Group ID","Start Date","SKU"}
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sponsored Display Campaigns"; ws.append(H)
for r in rows: ws.append([r[h] for h in H])
for c in range(1, len(H)+1):
    hdr = H[c-1]
    for cell in ws[get_column_letter(c)]:
        cell.font = Font(name="Arial", size=10, bold=(cell.row == 1))
        if cell.row > 1 and hdr in TEXT and cell.value is not None: cell.value = str(cell.value); cell.number_format = "@"
        if cell.row > 1 and hdr in ("Bid","Ad Group Default Bid","Budget") and cell.value is not None: cell.number_format = "0.00"
    ws.column_dimensions[get_column_letter(c)].width = max(12, min(36, len(hdr)+4))
ws.freeze_panes = "A2"
for cell in ws[1]: cell.fill = PatternFill("solid", fgColor="DDDDDD")
out = "VELLEN_UK_SD_展示再营销_1105.xlsx"
wb.save(out); print(len(rows), "rows ->", out)
