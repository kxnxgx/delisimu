import os, sys, io, re, subprocess, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

BASE = os.getcwd()
fmap = {f: os.path.join(BASE, f) for f in os.listdir(BASE)}
def fp(kw):
    for k,v in fmap.items():
        if kw in k: return v
    return None

FILE_V6 = fp("v6.xlsx")
FILE_BACKEND = fp("backend_calc.py")

def n(val,d=0):
    if isinstance(val,(int,float)): return float(val)
    try: return float(str(val).replace(",",""))
    except: return float(d)

wb = openpyxl.load_workbook(FILE_V6)
ws_s = wb["全体サマリー"]
ws_plc = wb["Pastel Lavender-Confetti Pat"]

# PLCのv6最新値を取得
supply_plc = 0
b2_text = ws_plc["B2"].value or ""
m = re.search(r"総供給:\s*(\d+)", b2_text)
if m: supply_plc = int(m.group(1))

demand_6_12 = int(round(sum(n(ws_plc.cell(mo+5, 6).value) for mo in range(6,13))))
plan_6_12   = int(round(sum(n(ws_plc.cell(mo+5, 8).value) for mo in range(6,13))))
plan_1_5    = int(round(sum(n(ws_plc.cell(mo+5, 8).value) for mo in range(1, 6))))
annual_plan = plan_1_5 + plan_6_12

print(f"PLC (Pastel Lavender-Confetti Pat) の最新値:")
print(f"  供給: {supply_plc}")
print(f"  需要(6-12月): {demand_6_12}")
print(f"  計画(6-12月): {plan_6_12}")
print(f"  年間計画: {annual_plan}")
print(f"  過不足: {supply_plc - demand_6_12}")

# 全体サマリーのPLC行 (Row30) を更新
print(f"\n  全体サマリー Row30 更新...")
print(f"  更新前: 需要={ws_s.cell(30,9).value} 方法={ws_s.cell(30,10).value} 計画={ws_s.cell(30,11).value} 過不足={ws_s.cell(30,12).value}")

ws_s.cell(30, 9).value  = demand_6_12
ws_s.cell(30, 10).value = "需要追随（制約なし）"
ws_s.cell(30, 11).value = annual_plan
ws_s.cell(30, 12).value = supply_plc - demand_6_12

print(f"  更新後: 需要={demand_6_12} 方法=需要追随（制約なし） 計画={annual_plan} 過不足={supply_plc - demand_6_12}")

# 合計行 (Row32) 再計算
for col in [9, 11, 12]:
    ws_s.cell(32, col).value = int(round(sum(n(ws_s.cell(r, col).value) for r in range(5, 32))))
print(f"\n  合計行更新: 需要合計={ws_s.cell(32,9).value} 計画合計={ws_s.cell(32,11).value}")

wb.save(FILE_V6)
print(f"\n✅ 全体サマリー修正・v6保存完了")

# ファイルサイズ確認
print(f"\nv6ファイルサイズ: {os.path.getsize(FILE_V6):,} bytes")

# backend_calc.py 更新
print("\n=== ダッシュボード更新 ===")
with open(FILE_BACKEND, "r", encoding="utf-8") as f: bc = f.read()

if "v5.xlsx" in bc:
    bc = bc.replace("v5.xlsx", "v6.xlsx")
    with open(FILE_BACKEND, "w", encoding="utf-8") as f: f.write(bc)
    print("  ✅ backend_calc.py: v5 → v6 に変更")
else:
    print("  ℹ️ すでに v6 参照")

# 現在の参照ファイルを確認
m = re.search(r'FILE_KANKEN\s*=\s*r"([^"]+)"', bc)
if m: print(f"  参照ファイル: {os.path.basename(m.group(1))}")

result = subprocess.run(
    [sys.executable, FILE_BACKEND],
    cwd=BASE, capture_output=True, text=True, encoding="utf-8", errors="replace"
)
for line in result.stdout.strip().split("\n"):
    if line.strip(): print(f"  > {line}")
if result.returncode != 0:
    print(f"  ⚠️ エラー:\n{result.stderr[:800]}")
else:
    print("  ✅ ダッシュボード更新完了")

print("\n✅ 全処理完了")
