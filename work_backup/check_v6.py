import os, sys, io, openpyxl, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

BASE = os.getcwd()
fmap = {f: os.path.join(BASE, f) for f in os.listdir(BASE)}
def fp(kw):
    for k,v in fmap.items():
        if kw in k: return v
    return None

FILE_V5 = fp("v5.xlsx")
FILE_V6 = fp("v6.xlsx")
NEW_TREND = 1.63
COLORS = {
    "Corn":1.775,"Pastel Lavender-Confetti Pat":4.0,
    "Chalk Rose":2.222,"Green":2.222,"Black Oak":2.222,
    "Teal":2.222,"Fossil":2.222,"Midnight Purple":2.222,
}
def n(val,d=0):
    if isinstance(val,(int,float)): return float(val)
    try: return float(str(val).replace(",",""))
    except: return float(d)

wb5=openpyxl.load_workbook(FILE_V5)
wb6=openpyxl.load_workbook(FILE_V6)

print("カラーシート 需要/計画 比較 (6-12月合計):")
print(f"{'カラー':<35} {'v5需要':>7} {'v6需要':>7} {'v5計画':>7} {'v6計画':>7} {'B2トレンド'}")
for cn in COLORS:
    if cn not in wb5.sheetnames or cn not in wb6.sheetnames: continue
    ws5=wb5[cn]; ws6=wb6[cn]
    d5=sum(n(ws5.cell(m+5,6).value) for m in range(6,13))
    d6=sum(n(ws6.cell(m+5,6).value) for m in range(6,13))
    p5=sum(n(ws5.cell(m+5,8).value) for m in range(6,13))
    p6=sum(n(ws6.cell(m+5,8).value) for m in range(6,13))
    b2 = ws6["B2"].value or ""
    trend_match = re.search(r"トレンド比:\s*([\d.]+)", b2)
    trend_str = trend_match.group(1) if trend_match else "?"
    changed = "✅" if abs(d5-d6)>0.5 else "⚠️"
    print(f"{changed} {cn:<33} {int(d5):>7} {int(d6):>7} {int(p5):>7} {int(p6):>7}   {trend_str}")
