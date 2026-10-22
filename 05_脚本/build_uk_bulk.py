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


MAIN4 = ["17-EDY4-EPEG","HEI-CSQ-4/4","V1-DS9X-ZY88","02-H1HJ-R28I"]   # B0H5D8BDGB, B0H2ZQTWL3, B0H4WFM9BQ, B0H686FV6Q
uk = []
M, MAG = "229047174948599", "183914124760252"
uk.append(row(entity="Campaign", op="Update", cid=M, state="enabled", budget=20, strat=STRAT))
uk.append(row(entity="Bidding Adjustment", op="Update", cid=M, strat=STRAT, place="Placement Top", pct=100))
for kid, kw, bid, st in [("156138128538792","full size violin",0.80,"enabled"),("83813070560166","violin starter kit",0.85,"enabled"),
                         ("264460227097585","beginners violin",0.60,"enabled"),("62677920243743","violin for beginners",0.40,"enabled"),
                         ("269898694728437","violin",None,"paused"),("10646902757399","violin full size",None,"paused"),
                         ("257656545250958","mendini violin",None,"paused"),
                         ("250226390082941","violin for kids",None,"paused"),("213549787818623","violin for children",None,"paused"),
                         ("274635874964501","stentor violin",None,"paused"),("53646658957072","cecilio violin",None,"paused")]:
    uk.append(row(entity="Keyword", op="Update", cid=M, agid=MAG, kwid=kid, state=st, bid=bid, kw=kw, mt="exact"))
for adid, sku in [("414010789274759","OW-UVGA-Q0VE"),("325796317434266","AH-0GF9-XQYJ"),("474195856033054","UY-L9B6-LTZD"),
                  ("342779097216438","9O-2CAR-WQP7"),("503364881363560","I9-6TP3-PJSZ")]:
    uk.append(row(entity="Product Ad", op="Update", cid=M, agid=MAG, adid=adid, state="paused", sku=sku))
for t in MOVE_OUT:
    uk.append(row(entity="Negative Keyword", op="Create", cid=M, agid=MAG, state="enabled", kw=t, mt="negativeExact"))
uk += camp_negs(M)
for ag, dbid, skus, kws in [("主词violin", 0.60, MAIN4[:2], [("violin",0.60)]),
                            ("规格词", 0.40, MAIN4, [("violin full size",0.45),("violin outfit",0.45)])]:
    uk.append(row(entity="Ad Group", op="Create", cid=M, agid=ag, agname=ag, state="enabled", agbid=dbid))
    for s in skus: uk.append(row(entity="Product Ad", op="Create", cid=M, agid=ag, state="enabled", sku=s))
    for kw, b in kws: uk.append(row(entity="Keyword", op="Create", cid=M, agid=ag, state="enabled", bid=b, kw=kw, mt="exact"))

A, AAG = "20508816371867", "406880089330685"
uk.append(row(entity="Campaign", op="Update", cid=A, state="enabled", budget=20, strat=STRAT))
uk.append(row(entity="Bidding Adjustment", op="Update", cid=A, strat=STRAT, place="Placement Top", pct=30))
for ptid, e, b in [("176313668105615","close-match",0.45),("8384754637153","loose-match",0.45),("65635971890494","substitutes",0.40)]:
    uk.append(row(entity="Product Targeting", op="Update", cid=A, agid=AAG, ptid=ptid, state="enabled", bid=b, pte=e))
for adid, sku in [("530780434389162","OW-UVGA-Q0VE"),("349888461829841","AH-0GF9-XQYJ"),("529417267531948","UY-L9B6-LTZD"),
                  ("343014827164215","9O-2CAR-WQP7"),("296824582442766","I9-6TP3-PJSZ")]:
    uk.append(row(entity="Product Ad", op="Update", cid=A, agid=AAG, adid=adid, state="paused", sku=sku))
for t in HARVESTED:
    uk.append(row(entity="Negative Keyword", op="Create", cid=A, agid=AAG, state="enabled", kw=t, mt="negativeExact"))
uk += camp_negs(A)

C = "SP-UK-词组拓展"
uk.append(row(entity="Campaign", op="Create", cid=C, cname=C, start="20261022", ttype="Manual", state="enabled", budget=5, strat=STRAT))
uk.append(row(entity="Bidding Adjustment", op="Create", cid=C, strat=STRAT, place="Placement Top", pct=50))
uk.append(row(entity="Ad Group", op="Create", cid=C, agid="词组", agname="词组", state="enabled", agbid=0.40))
for s in MAIN4: uk.append(row(entity="Product Ad", op="Create", cid=C, agid="词组", state="enabled", sku=s))
for kw in ["beginner violin","violin set","violin kit","adult violin","full size violin","violin outfit"]:
    uk.append(row(entity="Keyword", op="Create", cid=C, agid="词组", state="enabled", bid=0.40, kw=kw, mt="phrase"))
for t in HARVESTED:
    uk.append(row(entity="Negative Keyword", op="Create", cid=C, agid="词组", state="enabled", kw=t, mt="negativeExact"))
uk += camp_negs(C)

P = "SP-UK-商品投放"
uk.append(row(entity="Campaign", op="Create", cid=P, cname=P, start="20261022", ttype="Manual", state="enabled", budget=5, strat=STRAT))
groups = [("高价竞品", 0.40, ['asin="B08GSBLP67"','asin="B0D2VZ5H6S"','asin="B0D2VZ2V2M"','asin="B08199RYL3"','asin="B0BLP7DCXY"'], True),
          ("头部竞品", 0.45, ['asin="B0DRFWNV26"','asin="B07VRY3D33"','asin="B0F9XW3LYH"'], True),
          ("品类高价段", 0.35, ['category="407859031" price>60'], True),
          ("自家防御", 0.25, ['asin="B0H5DD2GJ2"','asin="B0H4RTDCLY"','asin="B0H67YD6GS"','asin="B0H67M2KYG"','asin="B0H58G8R6G"'], False)]
for ag, b, exprs, neg in groups:
    uk.append(row(entity="Ad Group", op="Create", cid=P, agid=ag, agname=ag, state="enabled", agbid=b))
    for s in MAIN4: uk.append(row(entity="Product Ad", op="Create", cid=P, agid=ag, state="enabled", sku=s))
    for e in exprs: uk.append(row(entity="Product Targeting", op="Create", cid=P, agid=ag, state="enabled", bid=b, pte=e))
    if neg: uk.append(row(entity="Negative Product Targeting", op="Create", cid=P, agid=ag, state="enabled", pte='asin="B015ELTM8A"'))

G = "SP-UK-礼品Q4"
uk.append(row(entity="Campaign", op="Create", cid=G, cname=G, start="20261102", end="20261218", ttype="Manual", state="enabled", budget=5, strat=STRAT))
uk.append(row(entity="Bidding Adjustment", op="Create", cid=G, strat=STRAT, place="Placement Top", pct=30))
uk.append(row(entity="Ad Group", op="Create", cid=G, agid="礼品词", agname="礼品词", state="enabled", agbid=0.45))
for s in MAIN4: uk.append(row(entity="Product Ad", op="Create", cid=G, agid="礼品词", state="enabled", sku=s))
for kw in ["violin gift","violin gifts","gift for violin beginner","violin for adults","violin for teenager","violin starter kit adult"]:
    uk.append(row(entity="Keyword", op="Create", cid=G, agid="礼品词", state="enabled", bid=0.45, kw=kw, mt="phrase"))
uk += camp_negs(G)
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
n = save(uk, out + "VELLEN_UK_SP_批量上传_1022_v2.xlsx")
import collections
print("rows", n, collections.Counter((r["Entity"], r["Operation"]) for r in uk))
