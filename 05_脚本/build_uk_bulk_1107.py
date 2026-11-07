import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

H = ["Product","Entity","Operation","Campaign ID","Ad Group ID","Portfolio ID","Ad ID","Keyword ID","Product Targeting ID",
     "Campaign Name","Ad Group Name","Start Date","End Date","Targeting Type","State","Daily Budget","SKU","ASIN",
     "Ad Group Default Bid","Bid","Keyword Text","Match Type","Bidding Strategy","Placement","Percentage","Product Targeting Expression"]
SP = "Sponsored Products"; STRAT = "Dynamic bids - down only"

def row(**k):
    r = {h: None for h in H}; r["Product"] = SP
    m = {"entity":"Entity","op":"Operation","cid":"Campaign ID","agid":"Ad Group ID","adid":"Ad ID","kwid":"Keyword ID","ptid":"Product Targeting ID",
         "cname":"Campaign Name","agname":"Ad Group Name","start":"Start Date","end":"End Date","ttype":"Targeting Type","state":"State",
         "budget":"Daily Budget","sku":"SKU","agbid":"Ad Group Default Bid","bid":"Bid","kw":"Keyword Text","mt":"Match Type",
         "strat":"Bidding Strategy","place":"Placement","pct":"Percentage","pte":"Product Targeting Expression"}
    for a,v in k.items(): r[m[a]] = v
    return r

MAIN_SKUS = ["17-EDY4-EPEG","HEI-CSQ-4/4","V1-DS9X-ZY88"]   # B0H5D8BDGB, B0H2ZQTWL3, B0H4WFM9BQ
NEG_PHRASE = ["kids","children","child","toddler","small violin","electric","silent violin","banjolele",
              "ukulele","stentor","cecilio","aileen","rosin","shoulder rest","violin strings","violin bow","violin case"]
NEG_EXACT_CAMP = ["instruments"]
HARVESTED = ["full size violin","violin starter kit","beginners violin","beginner violin","violin","violin full size",
             "violin outfit","violin for beginners"]
MOVE_OUT = ["violin","violin full size","violin outfit"]

def camp_negs(cid):
    out = [row(entity="Campaign Negative Keyword", op="Create", cid=cid, state="enabled", kw=t, mt="negativePhrase") for t in NEG_PHRASE]
    out += [row(entity="Campaign Negative Keyword", op="Create", cid=cid, state="enabled", kw=t, mt="negativeExact") for t in NEG_EXACT_CAMP]
    return out



uk = []
def pause_ad(cid, agid, adid, sku):
    uk.append(row(entity="Product Ad", op="Update", cid=cid, agid=agid, adid=adid, state="paused", sku=sku))
M, MAG = "229047174948599", "183914124760252"
A, AAG = "20508816371867", "406880089330685"
# 1) B0H58HS28V（蓝色）两周 19 次点击 0 单 → 全部暂停
pause_ad(M, MAG, "558039812942441", "M1-5KHH-ECEF")
pause_ad(A, AAG, "542532478069127", "M1-5KHH-ECEF")
# 2) B0H686FV6Q（£82.50 实木款）两周 56 次点击 0 单 → 只留在高价场景（品类高价段、高价竞品、自家防御）和展示广告
pause_ad(A, AAG, "558823372257993", "02-H1HJ-R28I")
pause_ad(M, "180291292430716", "424930363444436", "02-H1HJ-R28I")
pause_ad("269772264727484", "176533990601582", "513360249067404", "02-H1HJ-R28I")
pause_ad("97938043967621", "145129340622005", "531509740998752", "02-H1HJ-R28I")
# 3) B0H4WFM9BQ（白色）在自动广告 18 次点击 0 单
pause_ad(A, AAG, "399861935363316", "V1-DS9X-ZY88")
# 4) 自动广告：订单都在搜索首页顶部（38 次点击 3 单），商品页 40 次点击 0 单
uk.append(row(entity="Bidding Adjustment", op="Update", cid=A, strat=STRAT, place="Placement Top", pct=80))
uk.append(row(entity="Product Targeting", op="Update", cid=A, agid=AAG, ptid="65635971890494", state="enabled", bid=0.34, pte="substitutes"))
TEXT_COLS = {"Campaign ID","Ad Group ID","Ad ID","Keyword ID","Product Targeting ID","Start Date","End Date","SKU"}
def save(rows, path):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sponsored Products Campaigns"
    ws.append(H)
    for r in rows:
        ws.append([r[h] for h in H])
    for c in range(1, len(H)+1):
        hdr = H[c-1]
        for cell in ws[get_column_letter(c)]:
            cell.font = Font(name="Arial", size=10, bold=(cell.row == 1))
            if cell.row > 1 and hdr in TEXT_COLS and cell.value is not None:
                cell.value = str(cell.value); cell.number_format = "@"
            if cell.row > 1 and hdr in ("Bid","Ad Group Default Bid","Daily Budget") and cell.value is not None:
                cell.number_format = "0.00"
        ws.column_dimensions[get_column_letter(c)].width = max(12, min(34, len(hdr) + 4))
    ws.freeze_panes = "A2"
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="DDDDDD")
    wb.save(path)
    return len(rows)



out = "./"
n = save(uk, out + "VELLEN_UK_SP_批量上传_1107.xlsx")
import collections
print("rows", n, collections.Counter((r["Entity"], r["Operation"]) for r in uk))
