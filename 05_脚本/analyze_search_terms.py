import csv, collections, sys, re
f = sys.argv[1]
r = [x for x in csv.DictReader(open(f, encoding='utf-8-sig')) if x['预算货币']=='GBP']
def num(v):
    try: return float(v)
    except: return 0.0
def rng(x):
    m = re.findall(r'(\d+)月(\d+)日', x['日期范围']); return (int(m[0][1]), int(m[1][1]))
def agg(keyf, rows):
    d = collections.defaultdict(lambda: [0,0,0.0,0,0.0])
    for x in rows:
        a = d[keyf(x)]
        a[0]+=num(x['展示量']); a[1]+=num(x['点击量']); a[2]+=num(x['总成本']); a[3]+=num(x['购买量']); a[4]+=num(x['销售额'])
    return d
fmt = lambda v: 'impr %d clk %d £%.2f ord %d £%.2f' % tuple(v)
print('--- violin exact keyword: date ranges ---')
v = [x for x in r if x['投放方案']=='violin' and x['广告活动名称']=='英国广告手动']
print(collections.Counter(x['日期范围'] for x in v))
print('--- violin kw by search term ---')
for k,a in sorted(agg(lambda x:x['搜索词'], v).items(), key=lambda z:-z[1][2]): print(k, fmt(a))
print('--- violin kw by placement ---')
for k,a in sorted(agg(lambda x:x['广告位分类'], v).items(), key=lambda z:-z[1][2]): print(k, fmt(a))
print('--- violin kw by ASIN ---')
for k,a in sorted(agg(lambda x:x['推广的商品编号'], v).items(), key=lambda z:-z[1][2]): print(k, fmt(a))
print('\n--- bucket by date range (all UK) ---')
b = agg(lambda x: ('pre<=17' if rng(x)[1]<=17 else ('post>=19' if rng(x)[0]>=19 else 'span')), r)
for k,a in b.items(): print(k, fmt(a))
print('\n--- placement all UK ---')
for k,a in sorted(agg(lambda x:(x['广告活动名称'],x['广告位分类']), r).items(), key=lambda z:-z[1][2]): print(k, fmt(a))
print('\n--- search terms all UK (spend desc) ---')
for k,a in sorted(agg(lambda x:x['搜索词'], r).items(), key=lambda z:-z[1][2])[:45]: print(repr(k), fmt(a))
print('\n--- ASIN x campaign ad IDs ---')
for k,a in sorted(agg(lambda x:(x['广告活动名称'],x['推广的商品编号'],x['推广的商品 SKU-Advertised product SKU'],x['广告 ID']), r).items()): print(k, fmt(a))
print('\n--- auto targets ---')
for k,a in agg(lambda x:(x['投放方案'],x['投放方案编号'],x['投放状态'],x['目标竞价']), [x for x in r if x['广告活动名称']=='英国广告自动']).items(): print(k, fmt(a))
