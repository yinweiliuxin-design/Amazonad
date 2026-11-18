import openpyxl, warnings
warnings.filterwarnings("ignore")
def load(path):
    wb=openpyxl.load_workbook(path)
    out={}
    for ws in wb.worksheets:
        rows=list(ws.iter_rows(values_only=True))
        if not rows: continue
        H=rows[0]; out[ws.title]=[dict(zip(H,r)) for r in rows[1:] if any(v not in (None,"") for v in r)]
    return out
