import sys, glob, openpyxl; sys.path.insert(0, "work")
from load import load
D = load("d1118/data.xlsx"); SP = D["商品推广活动"]; SB = D["品牌推广活动"]
ok = True
def bad(*a):
    global ok; ok = False; print("  !!", *a)
def one(rows, pred):
    m = [r for r in rows if pred(r)]
    if len(m) != 1: bad("match count", len(m)); return {}
    return m[0]
for f in sorted(glob.glob("out/*.xlsx")):
    print(f.split("/")[-1]); wb = openpyxl.load_workbook(f)
    for ws in wb.worksheets:
        H = [c.value for c in ws[1]]
        for row in ws.iter_rows(min_row=2):
            d = {H[i]: c.value for i, c in enumerate(row)}; e, op = d["Entity"], d["Operation"]
            for c in row:
                h = H[c.column-1]
                if h in ("Campaign ID","Ad Group ID","Ad ID","Keyword ID","Product Targeting ID","Targeting ID","SKU") and c.value is not None and (not isinstance(c.value, str) or c.number_format != "@"): bad("id not text", h)
            if d.get("Keyword Text") and "/" in d["Keyword Text"]: bad("slash")
            if op == "Update" and e in ("Keyword","Product Targeting") and (d["Bid"] is None or d["State"] is None): bad("blank bid/state", d)
            if ws.title.startswith("Sponsored Brands"):
                s = one(SB, lambda r: str(r.get("关键词编号")) == d["Keyword ID"])
                if (str(s["广告活动编号"]), str(s["广告组编号"]), s["关键词文本"]) != (d["Campaign ID"], d["Ad Group ID"], d["Keyword Text"]): bad(d)
                desc = f'品牌 {d["Keyword Text"]}: {s["竞价"]}→{d["Bid"]}'
            elif e == "Keyword" and op == "Update":
                s = one(SP, lambda r: str(r.get("关键词编号")) == d["Keyword ID"] and r["实体层级"] == "关键词")
                mt = {"精准":"exact","词组":"phrase"}[s["匹配类型"]]
                if (str(s["广告活动编号"]), str(s["广告组编号"]), s["关键词文本"], mt) != (d["Campaign ID"], d["Ad Group ID"], d["Keyword Text"], d["Match Type"]): bad(d)
                desc = f'{s["广告活动名称（仅供参考）"]} 关键词 {d["Keyword Text"]}: {s["竞价"]}→{d["Bid"]}'
            elif e == "Keyword" and op == "Create":
                ex = [r for r in SP if r["实体层级"] == "关键词" and str(r["广告组编号"]) == d["Ad Group ID"] and r["关键词文本"] == d["Keyword Text"]]
                if ex: bad("already exists", d["Keyword Text"])
                ag = one(SP, lambda r: r["实体层级"] == "广告组" and str(r["广告组编号"]) == d["Ad Group ID"])
                if str(ag["广告活动编号"]) != d["Campaign ID"]: bad("ag/campaign")
                desc = f'新建 {ag["广告组名称"]} 精准词 {d["Keyword Text"]}（用广告组默认竞价）'
            elif e == "Product Targeting":
                s = one(SP, lambda r: str(r.get("商品投放 ID")) == d["Product Targeting ID"] and r["实体层级"] == "商品定向")
                if (str(s["广告活动编号"]), str(s["广告组编号"]), s["拓展商品投放编号"]) != (d["Campaign ID"], d["Ad Group ID"], d["Product Targeting Expression"]): bad(d)
                desc = f'自动 {d["Product Targeting Expression"]}: {s["竞价"]}→{d["Bid"]}'
            elif e == "Bidding Adjustment":
                s = one(SP, lambda r: r["实体层级"] == "竞价调整" and str(r["广告活动编号"]) == d["Campaign ID"] and r["广告位"] == "广告位：搜索结果首页首位")
                if not d["Bidding Strategy"] or d["Placement"] != "Placement Top": bad(d)
                desc = f'首页首位 {s["广告活动名称（仅供参考）"]}: {s["百分比"]}→{d["Percentage"]}'
            elif e == "Ad Group":
                s = one(SP, lambda r: r["实体层级"] == "广告组" and str(r["广告组编号"]) == d["Ad Group ID"])
                if (str(s["广告活动编号"]), s["广告组名称"]) != (d["Campaign ID"], d["Ad Group Name"]) or d["Ad Group Default Bid"] is None: bad(d)
                desc = f'广告组 {d["Ad Group Name"]}: {s["状态"]}/{s["广告组默认竞价"]} → {d["State"]}/{d["Ad Group Default Bid"]}'
            elif e == "Product Ad":
                s = one(SP, lambda r: str(r.get("广告编号")) == d["Ad ID"] and r["实体层级"] == "商品广告")
                if (str(s["广告活动编号"]), str(s["广告组编号"]), s["SKU"]) != (d["Campaign ID"], d["Ad Group ID"], d["SKU"]): bad(d)
                desc = f'广告 {s["广告组名称（仅供参考）"]}/{d["SKU"]}: {s["状态"]}→{d["State"]}'
            else:
                bad("unknown", e); desc = ""
            print("   ", desc)
print("ALL OK" if ok else "PROBLEMS")
