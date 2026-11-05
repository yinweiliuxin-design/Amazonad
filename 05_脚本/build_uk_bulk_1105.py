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
# 1) SP-UK-主词violin：B0H2ZQTWL3 在 violin 上 48 次点击出 4 单，抢回首页顶部
V = "19159137932017"
uk.append(row(entity="Campaign", op="Update", cid=V, cname="SP-UK-主词violin", state="enabled", budget=10, strat=STRAT))
uk.append(row(entity="Bidding Adjustment", op="Update", cid=V, strat=STRAT, place="Placement Top", pct=100))

# 2) 英国广告自动
A, AAG = "20508816371867", "406880089330685"
for ptid, e, b in [("176313668105615","close-match",0.60),("65635971890494","substitutes",0.40),("8384754637153","loose-match",0.32)]:
    uk.append(row(entity="Product Targeting", op="Update", cid=A, agid=AAG, ptid=ptid, state="enabled", bid=b, pte=e))
uk.append(row(entity="Negative Product Targeting", op="Create", cid=A, agid=AAG, state="enabled", pte='asin="B015ELTM8A"'))
for t in ["left handed", "childrens", "forenza", "phoenix"]:
    uk.append(row(entity="Campaign Negative Keyword", op="Create", cid=A, state="enabled", kw=t, mt="negativePhrase"))

# 3) 英国广告手动：full size violin 两周 33 次点击 1 单，ACOS 47%
M, MAG = "229047174948599", "183914124760252"
uk.append(row(entity="Keyword", op="Update", cid=M, agid=MAG, kwid="156138128538792", state="enabled", bid=0.68, kw="full size violin", mt="exact"))

# 4) SP-UK-商品投放：搜索词 b07vry3d33 两周 20 次点击 0 单（达到 15 次点击 0 单的规则），暂停这个目标，出价保留 0.35
uk.append(row(entity="Product Targeting", op="Update", cid="250033464216246", agid="96921562714739", ptid="107248139457563", state="paused", bid=0.35, pte='asin="B07VRY3D33"'))
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
n = save(uk, out + "VELLEN_UK_SP_批量上传_1105_v2.xlsx")
import collections
print("rows", n, collections.Counter((r["Entity"], r["Operation"]) for r in uk))
