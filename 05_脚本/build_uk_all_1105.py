import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
B = "./"
out_wb = openpyxl.Workbook(); out_wb.remove(out_wb.active)

def add_sheet(title, header, rows, text_cols, num_cols):
    ws = out_wb.create_sheet(title); ws.append(header)
    for r in rows: ws.append(r)
    for c in range(1, len(header)+1):
        hdr = header[c-1]
        for cell in ws[get_column_letter(c)]:
            cell.font = Font(name="Arial", size=10, bold=(cell.row == 1))
            if cell.row > 1 and hdr in text_cols and cell.value is not None: cell.value = str(cell.value); cell.number_format = "@"
            if cell.row > 1 and hdr in num_cols and cell.value is not None: cell.number_format = "0.00"
        ws.column_dimensions[get_column_letter(c)].width = max(12, min(40, len(hdr)+4))
    ws.freeze_panes = "A2"
    for cell in ws[1]: cell.fill = PatternFill("solid", fgColor="DDDDDD")

# 1) 商品推广：直接取已复核的 1105_v2
for fname, title, tcols, ncols in [("VELLEN_UK_SP_批量上传_1105_v2.xlsx", "Sponsored Products Campaigns",
                                    {"Campaign ID","Ad Group ID","Ad ID","Keyword ID","Product Targeting ID","Start Date","End Date","SKU"},
                                    {"Bid","Ad Group Default Bid","Daily Budget"}),
                                   ("VELLEN_UK_SD_展示再营销_1105.xlsx", "Sponsored Display Campaigns",
                                    {"Campaign ID","Ad Group ID","Start Date","SKU"}, {"Bid","Ad Group Default Bid","Budget"})]:
    ws = openpyxl.load_workbook(B + fname).active
    vals = list(ws.iter_rows(values_only=True))
    add_sheet(title, list(vals[0]), [list(v) for v in vals[1:]], tcols, ncols)

# 3) 品牌推广·商品集
H = ["Product","Entity","Operation","Campaign ID","Draft Campaign ID","Portfolio ID","Ad Group ID","Keyword ID","Product Targeting ID",
     "Campaign Name","Start Date","End Date","State","Budget Type","Budget","Bid Optimization","Bid Multiplier","Bid","Keyword Text","Match Type",
     "Product Targeting Expression","Ad Format","Landing Page URL","Landing Page ASINs","Brand Entity ID","Brand Name","Brand Logo Asset ID",
     "Custom Image Asset ID","Creative Headline","Creative ASINs","Video Media IDs","Creative Type"]
def sb(**k):
    r = {h: None for h in H}; r["Product"] = "Sponsored Brands"; r.update(k); return [r[h] for h in H]
C = "SB-UK-商品集"; ASINS = "B0H2ZQTWL3,B0H4WFM9BQ,B0H686FV6Q"
rows = [sb(**{"Entity":"Campaign","Operation":"Create","Campaign ID":C,"Campaign Name":C,"Start Date":"20261105","State":"enabled",
              "Budget Type":"daily","Budget":5,"Bid Optimization":"Manual","Ad Format":"productCollection","Landing Page ASINs":ASINS,
              "Brand Entity ID":"ENTITY3TETS7J6K1PDG","Brand Name":"VELLEN",
              "Brand Logo Asset ID":"amzn1.assetlibrary.asset1.ae1a72b1cbf15162a68ff68174d6b107:version_v1",
              "Custom Image Asset ID":"amzn1.assetlibrary.asset1.4af177341278f689252fe011303e0d86:version_v1",
              "Creative Headline":"Full Size Violin Kit for Adult Beginners","Creative ASINs":ASINS})]
for kw in ["full size violin", "violin starter kit", "beginners violin", "violin outfit"]:
    rows.append(sb(**{"Entity":"Keyword","Operation":"Create","Campaign ID":C,"State":"enabled","Bid":0.80,"Keyword Text":kw,"Match Type":"exact"}))
NEG_PHRASE = ["kids","children","child","toddler","small violin","electric","silent violin","banjolele","ukulele","stentor","cecilio","aileen",
              "rosin","shoulder rest","violin strings","violin bow","violin case"]
for kw in NEG_PHRASE:
    rows.append(sb(**{"Entity":"Negative Keyword","Operation":"Create","Campaign ID":C,"State":"enabled","Keyword Text":kw,"Match Type":"negativePhrase"}))
rows.append(sb(**{"Entity":"Negative Keyword","Operation":"Create","Campaign ID":C,"State":"enabled","Keyword Text":"instruments","Match Type":"negativeExact"}))
add_sheet("Sponsored Brands Campaigns", H, rows, {"Campaign ID","Start Date","Brand Entity ID"}, {"Bid","Budget"})
out = B + "VELLEN_UK_一键上传_1105_全部广告.xlsx"
out_wb.save(out); print(out, [ (ws.title, ws.max_row-1) for ws in out_wb.worksheets])
