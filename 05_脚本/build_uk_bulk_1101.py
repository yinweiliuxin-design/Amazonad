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



EXTRA_NEG = ["electronic", "second hand", "viola", "plastic", "toy"]
uk = []
def pause_ad(cid, agid, adid, sku):
    uk.append(row(entity="Product Ad", op="Update", cid=cid, agid=agid, adid=adid, state="paused", sku=sku))

# 1) 新活动：主词 violin 单独成活动，预算封顶，只推转化最好的 B0H2ZQTWL3
V = "SP-UK-主词violin"
uk.append(row(entity="Campaign", op="Create", cid=V, cname=V, start="20261101", ttype="Manual", state="enabled", budget=8, strat=STRAT))
uk.append(row(entity="Bidding Adjustment", op="Create", cid=V, strat=STRAT, place="Placement Top", pct=50))
uk.append(row(entity="Ad Group", op="Create", cid=V, agid="violin", agname="violin", state="enabled", agbid=0.50))
uk.append(row(entity="Product Ad", op="Create", cid=V, agid="violin", state="enabled", sku="HEI-CSQ-4/4"))
uk.append(row(entity="Keyword", op="Create", cid=V, agid="violin", state="enabled", bid=0.50, kw="violin", mt="exact"))

# 2) 英国广告手动
M, MAG = "229047174948599", "183914124760252"
uk.append(row(entity="Ad Group", op="Update", cid=M, agid="40634066844513", state="paused"))   # 主词violin
uk.append(row(entity="Ad Group", op="Update", cid=M, agid="59743358080140", state="paused"))   # UK_单数控价
for kid, kw, bid, st in [("83813070560166","violin starter kit",0.95,"enabled"),
                         ("266106025908088","beginner violin",None,"paused"),
                         ("144234842447683","full size violin for beginners",0.80,"enabled"),
                         ("92369341821610","beginner violin set",0.80,"enabled")]:
    uk.append(row(entity="Keyword", op="Update", cid=M, agid=MAG, kwid=kid, state=st, bid=bid, kw=kw, mt="exact"))
pause_ad(M, MAG, "300024335676873", "17-EDY4-EPEG")              # B0H5D8BDGB 原广告组
pause_ad(M, MAG, "372316185953749", "02-H1HJ-R28I")              # B0H686FV6Q 原广告组 17 次点击 0 单
pause_ad(M, "180291292430716", "298000543800250", "17-EDY4-EPEG")  # B0H5D8BDGB 规格词

# 3) 英国广告自动
A, AAG = "20508816371867", "406880089330685"
uk.append(row(entity="Bidding Adjustment", op="Update", cid=A, strat=STRAT, place="Placement Top", pct=50))
for ptid, e, b, st in [("176313668105615","close-match",0.52,"enabled"),("65635971890494","substitutes",0.46,"enabled"),
                       ("8384754637153","loose-match",0.38,"enabled"),("246560076849167","complements",None,"paused")]:
    uk.append(row(entity="Product Targeting", op="Update", cid=A, agid=AAG, ptid=ptid, state=st, bid=b, pte=e))
for t in EXTRA_NEG:
    uk.append(row(entity="Campaign Negative Keyword", op="Create", cid=A, state="enabled", kw=t, mt="negativePhrase"))

# 4) SP-UK-商品投放
P = "250033464216246"
G_TOP, G_HIGH, G_CAT, G_DEF = "96921562714739", "87149978391144", "221594475359256", "277770336144606"
uk.append(row(entity="Product Targeting", op="Update", cid=P, agid=G_TOP, ptid="107248139457563", state="enabled", bid=0.35, pte='asin="B07VRY3D33"'))
uk.append(row(entity="Product Targeting", op="Update", cid=P, agid=G_TOP, ptid="245050517102629", state="enabled", bid=0.52, pte='asin="B0DRFWNV26"'))
uk.append(row(entity="Product Targeting", op="Create", cid=P, agid=G_DEF, state="enabled", bid=0.25, pte='asin="B0H5D8BDGB"'))
pause_ad(P, G_TOP, "430837185816712", "02-H1HJ-R28I")   # B0H686FV6Q 头部竞品 15 次点击 0 单
for agid, adid in [(G_TOP,"429895487557202"),(G_HIGH,"301226565784607"),(G_CAT,"378053289400869"),(G_DEF,"531846008851244")]:
    pause_ad(P, agid, adid, "17-EDY4-EPEG")              # B0H5D8BDGB

# 5) SP-UK-词组拓展
C, CAG = "269772264727484", "176533990601582"
pause_ad(C, CAG, "433589306641091", "17-EDY4-EPEG")
for t in EXTRA_NEG:
    uk.append(row(entity="Campaign Negative Keyword", op="Create", cid=C, state="enabled", kw=t, mt="negativePhrase"))

# 6) SP-UK-礼品Q4（11-02 自动开始，不动活动本身）
GF, GAG = "97938043967621", "145129340622005"
pause_ad(GF, GAG, "296796127457022", "17-EDY4-EPEG")
for t in EXTRA_NEG:
    uk.append(row(entity="Campaign Negative Keyword", op="Create", cid=GF, state="enabled", kw=t, mt="negativePhrase"))
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
n = save(uk, out + "VELLEN_UK_SP_批量上传_1101.xlsx")
import collections
print("rows", n, collections.Counter((r["Entity"], r["Operation"]) for r in uk))
