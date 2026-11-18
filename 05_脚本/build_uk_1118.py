import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
B = "/tmp/claude-0/-home-claude/c54da192-4a1a-535e-b769-ae849e7895f0/scratchpad/"
SP_H = [c.value for c in openpyxl.load_workbook(B+"tpl/sp.xlsx").active[1]]
SD_H = [c.value for c in openpyxl.load_workbook(B+"tpl/sd.xlsx").active[1]]
SB_H = [c.value for c in openpyxl.load_workbook(B+"tpl/sb.xlsx").active[1]]
STRAT = "Dynamic bids - down only"
TEXT = {"Campaign ID","Ad Group ID","Ad ID","Keyword ID","Product Targeting ID","Targeting ID","Start Date","End Date","SKU"}
NUM = {"Bid","Ad Group Default Bid","Daily Budget","Budget"}
# ---- IDs (11-18 下载文件) ----
AUTO, AUTO_AG = "20508816371867", "406880089330685"
MAN, AI, SPEC = "229047174948599", "183914124760252", "180291292430716"
VIO, VIO_AG, VIO_KW = "19159137932017", "100129758954501", "202600401370051"
PT, PT_TOP = "250033464216246", "96921562714739"
GIFT, GIFT_AG, GIFT_ADULT = "97938043967621", "145129340622005", "67157434285473"
SB_C, SB_AG = "10326251134091", "365427574718820"
SB_KW = {"full size violin": "28592408317053", "violin starter kit": "145451046089998", "beginners violin": "20183012039583", "violin outfit": "262295944693013"}
NEW_KW = ["adult violin", "violin adult", "adult beginner violin", "violin for adult beginners", "violin full set", "violin set"]

def mk(H, prod, k):
    r = {h: None for h in H}; r["Product"] = prod
    for key in k:
        assert key in H, key
    r.update(k); return r
sp = lambda **k: mk(SP_H, "Sponsored Products", k)
sb = lambda **k: mk(SB_H, "Sponsored Brands", k)
def kw(c, ag, kid, text, bid, mt="exact"):
    return sp(**{"Entity":"Keyword","Operation":"Update","Campaign ID":c,"Ad Group ID":ag,"Keyword ID":kid,"State":"enabled","Bid":bid,"Keyword Text":text,"Match Type":mt})
def pt(c, ag, pid, expr, bid):
    return sp(**{"Entity":"Product Targeting","Operation":"Update","Campaign ID":c,"Ad Group ID":ag,"Product Targeting ID":pid,"State":"enabled","Bid":bid,"Product Targeting Expression":expr})
def tos(c, pct):
    return sp(**{"Entity":"Bidding Adjustment","Operation":"Update","Campaign ID":c,"Bidding Strategy":STRAT,"Placement":"Placement Top","Percentage":pct})
def agbid(c, ag, name, bid, state="enabled"):
    return sp(**{"Entity":"Ad Group","Operation":"Update","Campaign ID":c,"Ad Group ID":ag,"Ad Group Name":name,"State":state,"Ad Group Default Bid":bid})
def ad(c, ag, adid, sku, state):
    return sp(**{"Entity":"Product Ad","Operation":"Update","Campaign ID":c,"Ad Group ID":ag,"Ad ID":adid,"State":state,"SKU":sku})
def sbkw(text, bid):
    return sb(**{"Entity":"Keyword","Operation":"Update","Campaign ID":SB_C,"Ad Group ID":SB_AG,"Keyword ID":SB_KW[text],"State":"enabled","Bid":bid,"Keyword Text":text,"Match Type":"exact"})

def levels(L, gift=True):
    s = [kw(VIO, VIO_AG, VIO_KW, "violin", L["violin"]), tos(VIO, L["vtos"]), tos(AUTO, L["atos"]), tos(MAN, L["mtos"]),
         pt(AUTO, AUTO_AG, "176313668105615", "close-match", L["close"]),
         pt(AUTO, AUTO_AG, "8384754637153", "loose-match", L["loose"]),
         pt(AUTO, AUTO_AG, "65635971890494", "substitutes", L["subs"]),
         kw(MAN, AI, "156138128538792", "full size violin", L["fsv"]),
         kw(MAN, AI, "264460227097585", "beginners violin", L["beg"]),
         kw(MAN, AI, "62677920243743", "violin for beginners", L["vfb"]),
         kw(MAN, AI, "83813070560166", "violin starter kit", L["vsk"]),
         kw(MAN, SPEC, "35531464760178", "violin full size", L["vfs"]),
         agbid(MAN, SPEC, "规格词", L["spec"])]
    if gift: s.append(kw(GIFT, GIFT_AG, GIFT_ADULT, "violin for adults", L["gift"], "phrase"))
    return s, [sbkw(t, L["sb"]) for t in SB_KW], []

base = dict(violin=0.45, vtos=70, atos=120, mtos=60, close=0.70, loose=0.32, subs=0.34, fsv=0.58, beg=0.55, vfb=0.40, vsk=0.95, vfs=0.45, spec=0.55, gift=0.55, sb=0.80)
bf   = dict(base, vtos=100, close=0.80, loose=0.36, subs=0.38, fsv=0.65, spec=0.63, gift=0.63, sb=0.92)
cut  = {k: (v if k in ("vtos","atos","mtos") else round(v*0.7 + 1e-9, 2)) for k, v in base.items()}

s1, b1, _ = levels(base)
s1 += [
    ad(MAN, AI, "521784471659665", "V1-DS9X-ZY88", "paused"),
    ad(MAN, SPEC, "447505358938581", "V1-DS9X-ZY88", "paused"),
    ad(MAN, SPEC, "298000543800250", "17-EDY4-EPEG", "enabled"),
    agbid(PT, PT_TOP, "头部竞品", 0.45, "paused"),
] + [sp(**{"Entity":"Keyword","Operation":"Create","Campaign ID":MAN,"Ad Group ID":SPEC,"State":"enabled","Keyword Text":t,"Match Type":"exact"}) for t in NEW_KW]
files = {
  "1_1118_出单优先_现在上传": (s1, [], []),
  "2_黑五开始当天_优惠券生效后传": levels(bf),
  "3_1202_黑五后恢复": levels(base),
  "4_圣诞送不到那天_降价": levels(cut, gift=False),
  "5_1226_节礼日恢复": levels(base, gift=False),
}
def write(name, sp_rows, sb_rows, sd_rows):
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    for title, H, rows in [("Sponsored Products Campaigns", SP_H, sp_rows), ("Sponsored Brands Campaigns", SB_H, sb_rows), ("Sponsored Display Campaigns", SD_H, sd_rows)]:
        if not rows: continue
        ws = wb.create_sheet(title); ws.append(H)
        for r in rows: ws.append([r[h] for h in H])
        for c in range(1, len(H)+1):
            hdr = H[c-1]
            for cell in ws[get_column_letter(c)]:
                cell.font = Font(name="Arial", size=10, bold=(cell.row == 1))
                if cell.row > 1 and hdr in TEXT and cell.value is not None: cell.value = str(cell.value); cell.number_format = "@"
                if cell.row > 1 and hdr in NUM and cell.value is not None: cell.number_format = "0.00"
            ws.column_dimensions[get_column_letter(c)].width = max(12, min(36, len(hdr)+4))
        ws.freeze_panes = "A2"
        for cell in ws[1]: cell.fill = PatternFill("solid", fgColor="DDDDDD")
    out = B + "out/VELLEN_UK_" + name + ".xlsx"; wb.save(out)
    return name, [(ws.title, ws.max_row-1) for ws in wb.worksheets]
if __name__ == "__main__":
    for n, (a, b, c) in files.items(): print(write(n, a, b, c))
    print("cut:", cut)
