import openpyxl, glob, sys
data = openpyxl.load_workbook("b1112/data.xlsx")
def rows_of(name):
    rs=list(data[name].iter_rows(values_only=True)); H=rs[0]; return [dict(zip(H,r)) for r in rs[1:]]
SP=rows_of("商品推广活动"); SB=rows_of("品牌推广活动"); SD=rows_of("展示型推广活动")
def find(rows, col, v): 
    m=[r for r in rows if str(r.get(col))==v]; assert len(m)>=1, (col,v); return m[0]
ok=True
def bad(*a):
    global ok; ok=False; print("  !!", *a)
for f in sorted(glob.glob("bulk/phases/*.xlsx")):
    name=f.split("/")[-1]; print(name); wb=openpyxl.load_workbook(f)
    for ws in wb.worksheets:
        H=[c.value for c in ws[1]]
        for row in ws.iter_rows(min_row=2):
            d={H[i]:c.value for i,c in enumerate(row)}; e=d["Entity"]
            for c in row:
                if H[c.column-1] in ("Campaign ID","Ad Group ID","Ad ID","Keyword ID","Product Targeting ID","Targeting ID","SKU") and c.value is not None and (not isinstance(c.value,str) or c.number_format!="@"): bad("id not text", H[c.column-1])
            if d.get("Keyword Text") and "/" in d["Keyword Text"]: bad("slash")
            if d["Operation"]=="Update" and e in ("Keyword","Product Targeting","Audience Targeting") and (d["Bid"] is None or d["State"] is None): bad("blank bid/state", d)
            sheet=ws.title.split()[1]
            if sheet=="Products":
                if e=="Keyword":
                    s=find(SP,"关键词编号",d["Keyword ID"]); assert (str(s["广告活动编号"]),str(s["广告组编号"]),s["关键词文本"],s["匹配类型"])==(d["Campaign ID"],d["Ad Group ID"],d["Keyword Text"],"精准")
                    desc=f'关键词 {d["Keyword Text"]}: {s["竞价"]}→{d["Bid"]}'
                elif e=="Product Targeting":
                    s=find(SP,"商品投放 ID",d["Product Targeting ID"]); assert (str(s["广告活动编号"]),str(s["广告组编号"]),s["拓展商品投放编号"])==(d["Campaign ID"],d["Ad Group ID"],d["Product Targeting Expression"])
                    desc=f'投放 {d["Product Targeting Expression"]}: {s["竞价"]}→{d["Bid"]} {d["State"]}'
                elif e=="Bidding Adjustment":
                    s=[r for r in SP if r["实体层级"]=="竞价调整" and str(r["广告活动编号"])==d["Campaign ID"] and r["广告位"]=="广告位：搜索结果首页首位"][0]
                    assert d["Bidding Strategy"] and d["Placement"]=="Placement Top"
                    desc=f'首页首位 {s["广告活动名称（仅供参考）"]}: {s["百分比"]}→{d["Percentage"]}'
                elif e=="Ad Group":
                    s=[r for r in SP if r["实体层级"]=="广告组" and str(r["广告组编号"])==d["Ad Group ID"]][0]
                    assert (str(s["广告活动编号"]),s["广告组名称"],s["广告组默认竞价"])==(d["Campaign ID"],d["Ad Group Name"],d["Ad Group Default Bid"])
                    desc=f'广告组 {d["Ad Group Name"]}→{d["State"]}'
                elif e=="Product Ad":
                    s=find(SP,"广告编号",d["Ad ID"]); assert (str(s["广告活动编号"]),str(s["广告组编号"]),s["SKU"])==(d["Campaign ID"],d["Ad Group ID"],d["SKU"])
                    desc=f'广告 {s["广告组名称（仅供参考）"]}/{d["SKU"]}→{d["State"]}'
                elif e=="Campaign":
                    s=[r for r in SP if r["实体层级"]=="广告活动" and str(r["广告活动编号"])==d["Campaign ID"]][0]
                    assert (s["广告活动名称"],s["每日预算"])==(d["Campaign Name"],d["Daily Budget"]) and not s["结束日期"] and d["Bidding Strategy"]
                    desc=f'活动 {d["Campaign Name"]}→{d["State"]}'
                else:
                    assert e=="Campaign Negative Keyword" and d["Match Type"]=="negativePhrase"
                    existing=[r["关键词文本"] for r in SP if r["实体层级"]=="广告活动否定关键词" and str(r["广告活动编号"])==d["Campaign ID"]]
                    if d["Keyword Text"] in existing: bad("neg already exists")
                    desc=f'否定词组 {d["Keyword Text"]}'
            elif sheet=="Brands":
                s=find(SB,"关键词编号",d["Keyword ID"]); assert (str(s["广告活动编号"]),str(s["广告组编号"]),s["关键词文本"])==(d["Campaign ID"],d["Ad Group ID"],d["Keyword Text"])
                desc=f'品牌 {d["Keyword Text"]}: {s["竞价"]}→{d["Bid"]}'
            else:
                s=find(SD,"投放 ID",d["Targeting ID"]); assert (str(s["广告活动编号"]),str(s["广告组编号"]),s["投放表达式"])==(d["Campaign ID"],d["Ad Group ID"],d["Targeting Expression"])
                desc=f'展示 {d["Targeting Expression"]}: {s["竞价"]}→{d["Bid"]} {d["State"]}'
            print("   ", desc)
print("ALL OK" if ok else "PROBLEMS")
