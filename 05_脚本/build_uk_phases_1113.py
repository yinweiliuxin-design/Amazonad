import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
B = "/tmp/claude-0/-home-claude/c54da192-4a1a-535e-b769-ae849e7895f0/scratchpad/bulk/"
SP_H = [c.value for c in openpyxl.load_workbook(B + "VELLEN_UK_SP_批量上传_1022_v2.xlsx").active[1]]
SD_H = [c.value for c in openpyxl.load_workbook(B + "VELLEN_UK_SD_展示再营销_1105.xlsx").active[1]]
SB_H = [c.value for c in openpyxl.load_workbook(B + "VELLEN_UK_SB_品牌商品集_1105_修正.xlsx").active[1]]
STRAT = "Dynamic bids - down only"
TEXT = {"Campaign ID","Ad Group ID","Ad ID","Keyword ID","Product Targeting ID","Targeting ID","Start Date","End Date","SKU","Brand Entity ID"}
NUM = {"Bid","Ad Group Default Bid","Daily Budget","Budget","Percentage"}

# ---------- 编号（全部来自 11-12 下载文件） ----------
AUTO, AUTO_AG = "20508816371867", "406880089330685"
MAN, MAN_AG = "229047174948599", "183914124760252"
VIO, VIO_AG, VIO_KW = "19159137932017", "100129758954501", "202600401370051"
PT, PT_CAT, PT_TOP, PT_HIGH = "250033464216246", "221594475359256", "96921562714739", "87149978391144"
PHR = "269772264727484"
SB_C, SB_AG = "10326251134091", "365427574718820"
SB_KW = {"full size violin": "28592408317053", "violin starter kit": "145451046089998", "beginners violin": "20183012039583", "violin outfit": "262295944693013"}
SD_C, SD_AG, SD_EXACT, SD_SIM = "92054206571906", "376658876390970", "74535790350551", "172997081492927"

def sp(**k):
    r = {h: None for h in SP_H}; r["Product"] = "Sponsored Products"; r.update(k); return r
def sb(**k):
    r = {h: None for h in SB_H}; r["Product"] = "Sponsored Brands"; r.update(k); return r
def sd(**k):
    r = {h: None for h in SD_H}; r["Product"] = "Sponsored Display"; r.update(k); return r

def sp_kw(cid, ag, kid, text, bid):
    return sp(**{"Entity":"Keyword","Operation":"Update","Campaign ID":cid,"Ad Group ID":ag,"Keyword ID":kid,"State":"enabled","Bid":bid,"Keyword Text":text,"Match Type":"exact"})
def sp_pt(cid, ag, ptid, expr, bid, state="enabled"):
    return sp(**{"Entity":"Product Targeting","Operation":"Update","Campaign ID":cid,"Ad Group ID":ag,"Product Targeting ID":ptid,"State":state,"Bid":bid,"Product Targeting Expression":expr})
def sp_tos(cid, pct):
    return sp(**{"Entity":"Bidding Adjustment","Operation":"Update","Campaign ID":cid,"Bidding Strategy":STRAT,"Placement":"Placement Top","Percentage":pct})
def sb_kw(text, bid):
    return sb(**{"Entity":"Keyword","Operation":"Update","Campaign ID":SB_C,"Ad Group ID":SB_AG,"Keyword ID":SB_KW[text],"State":"enabled","Bid":bid,"Keyword Text":text,"Match Type":"exact"})
def sd_aud(tid, expr, bid, state="enabled"):
    return sd(**{"Entity":"Audience Targeting","Operation":"Update","Campaign ID":SD_C,"Ad Group ID":SD_AG,"Targeting ID":tid,"State":state,"Bid":bid,"Targeting Expression":expr})
EXACT, SIM = "views=(exact-product lookback=30)", "views=(similar-product lookback=30)"

def levels(violin, vtos, atos, close, loose, subs, fsv, beg, vfs, vfb, vsk, drf, sbbid):
    """一整套出价水平 → 对应的更新行（每个文件都是完整状态，不依赖上一个文件是否上传）"""
    s = [sp_kw(VIO, VIO_AG, VIO_KW, "violin", violin), sp_tos(VIO, vtos), sp_tos(AUTO, atos),
         sp_pt(AUTO, AUTO_AG, "176313668105615", "close-match", close),
         sp_pt(AUTO, AUTO_AG, "8384754637153", "loose-match", loose),
         sp_pt(AUTO, AUTO_AG, "65635971890494", "substitutes", subs),
         sp_kw(MAN, MAN_AG, "156138128538792", "full size violin", fsv),
         sp_kw(MAN, MAN_AG, "264460227097585", "beginners violin", beg),
         sp_kw(MAN, "180291292430716", "35531464760178", "violin full size", vfs),
         sp_kw(MAN, MAN_AG, "62677920243743", "violin for beginners", vfb),
         sp_kw(MAN, MAN_AG, "83813070560166", "violin starter kit", vsk),
         sp_pt(PT, PT_TOP, "245050517102629", 'asin="B0DRFWNV26"', drf)]
    b = [sb_kw(t, sbbid) for t in SB_KW]
    return s, b, []

files = {}
# A 11-13 现在：收紧
A_sp = [sp_kw(VIO, VIO_AG, VIO_KW, "violin", 0.45), sp_tos(VIO, 50), sp_tos(AUTO, 120),
        sp_pt(AUTO, AUTO_AG, "65635971890494", "substitutes", 0.29),
        sp_kw(MAN, MAN_AG, "156138128538792", "full size violin", 0.58),
        sp(**{"Entity":"Ad Group","Operation":"Update","Campaign ID":PT,"Ad Group ID":PT_CAT,"Ad Group Name":"品类高价段","State":"paused","Ad Group Default Bid":0.35}),
        sp(**{"Entity":"Ad Group","Operation":"Update","Campaign ID":PT,"Ad Group ID":PT_HIGH,"Ad Group Name":"高价竞品","State":"paused","Ad Group Default Bid":0.40}),
        sp_pt(PT, PT_HIGH, "33643310510477", 'asin="B0D2VZ2V2M"', 0.40, "paused"),
        sp(**{"Entity":"Product Ad","Operation":"Update","Campaign ID":PT,"Ad Group ID":PT_TOP,"Ad ID":"331365363440353","State":"paused","SKU":"V1-DS9X-ZY88"}),
        sp(**{"Entity":"Campaign","Operation":"Update","Campaign ID":PHR,"Campaign Name":"SP-UK-词组拓展","State":"paused","Daily Budget":5,"Bidding Strategy":STRAT}),
        sp(**{"Entity":"Campaign Negative Keyword","Operation":"Create","Campaign ID":AUTO,"State":"enabled","Keyword Text":"half size","Match Type":"negativePhrase"})]
A_sd = [sd_aud(SD_SIM, SIM, 0.35, "paused")]
files["1_1113_现在上传"] = (A_sp, [], A_sd)
# 基准（A 之后）与各阶段
base = dict(violin=0.45, vtos=50, atos=120, close=0.60, loose=0.32, subs=0.29, fsv=0.58, beg=0.60, vfs=0.45, vfb=0.40, vsk=0.95, drf=0.52, sbbid=0.80)
bf   = dict(base, vtos=80, close=0.69, fsv=0.67, drf=0.60, sbbid=0.92)
cut  = {k: (v if k in ("vtos", "atos") else round(v*0.7 + 1e-9, 2)) for k, v in base.items()}
files["2_1120_黑五周_有优惠券才传"] = levels(**bf)
files["3_1202_黑五后恢复"] = levels(**base)
files["4_1221_圣诞截单后降价"] = levels(**cut)
files["5_1226_节礼日恢复"] = levels(**base)

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
    out = B + "phases/VELLEN_UK_" + name + ".xlsx"; wb.save(out)
    return out, [(ws.title, ws.max_row-1) for ws in wb.worksheets]
for n, (a, b, c) in files.items():
    print(write(n, a, b, c))
print("cut levels:", cut)
