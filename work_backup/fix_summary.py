import os, sys, io, shutil, csv, re, subprocess, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

BASE = os.getcwd()
fmap = {f: os.path.join(BASE, f) for f in os.listdir(BASE)}
def fp(kw):
    for k,v in fmap.items():
        if kw in k: return v
    return None

FILE_V6      = fp("v6.xlsx")
FILE_BACKEND = fp("backend_calc.py")

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

if FILE_V6 is None:
    print("v6が見つかりません。先にv6を作成してください。")
    sys.exit(1)

print("v6:", FILE_V6, "存在:", os.path.exists(FILE_V6))

wb=openpyxl.load_workbook(FILE_V6)

# 各カラーの計画デルタを再取得
# (v5からもとの値を読む必要があるのでv5から再計算)
FILE_V5 = fp("v5.xlsx")
wb5 = openpyxl.load_workbook(FILE_V5)

dd={m:0.0 for m in range(6,13)}
dp={m:0.0 for m in range(6,13)}

for cn in COLORS:
    if cn not in wb.sheetnames or cn not in wb5.sheetnames: continue
    ws6=wb[cn]; ws5=wb5[cn]
    for m in range(6,13):
        ri=m+5
        old_d=n(ws5.cell(ri,6).value); new_d=n(ws6.cell(ri,6).value)
        old_p=n(ws5.cell(ri,8).value); new_p=n(ws6.cell(ri,8).value)
        dd[m]+=new_d-old_d; dp[m]+=new_p-old_p

print("\n  月別 計画デルタ:")
for m in range(6,13): print(f"    {m}月: {int(round(dp[m])):+d}")

# 全体サマリー
print("\n  全体サマリー確認...")
ws_s=wb["全体サマリー"]
for cn in COLORS:
    tgt=None
    for ri in range(5,35):
        if ws_s.cell(ri,3).value and str(ws_s.cell(ri,3).value).strip()==cn: tgt=ri; break
    if tgt:
        print(f"    {cn}: 需要={ws_s.cell(tgt,9).value}, 計画={ws_s.cell(tgt,11).value}, 過不足={ws_s.cell(tgt,12).value}")

# 2026年着地予測 更新
print("\n  2026年着地予測 更新...")
ws_f=wb["2026年着地予測"]

# 現在のMEI残から6-12月を確認（列4=需要, 列6=計画）
print("  更新前 着地予測 6-12月:")
for m in range(6,13):
    ri=m+4
    print(f"    {m}月 需要={int(n(ws_f.cell(ri,4).value))} 計画={int(n(ws_f.cell(ri,6).value))}")

# デルタ適用（既にv6のカラーシートが更新済なのでサマリーも更新)
# 着地予測の6-12月は既にv6カラーシートの結果を反映していないため、デルタで修正
cumul=n(ws_f.cell(9,8).value)  # MAY累計
for m in range(6,13):
    ri=m+4
    new_d=round(n(ws_f.cell(ri,4).value)+dd[m])
    new_p=round(n(ws_f.cell(ri,6).value)+dp[m])
    ws_f.cell(ri,4).value=new_d; ws_f.cell(ri,6).value=new_p
    cumul+=new_p; ws_f.cell(ri,8).value=round(cumul)

td_=round(sum(n(ws_f.cell(r,4).value) for r in range(5,17)))
tp_=round(sum(n(ws_f.cell(r,6).value) for r in range(5,17)))
ws_f.cell(17,4).value=td_; ws_f.cell(17,6).value=tp_; ws_f.cell(17,7).value=100; ws_f.cell(17,8).value=tp_
for ri in range(5,17): ws_f.cell(ri,7).value=round(n(ws_f.cell(ri,6).value)/tp_*100,2) if tp_ else 0
print(f"  ✅ 着地予測更新: 年間需要={td_:,} / 年間計画={tp_:,}")

# 供給需要シミュレーション 更新
print("\n  供給需要シミュレーション 更新...")
ws_sim=wb["供給需要シミュレーション"]
cur=n(ws_sim.cell(5,3).value)
print(f"    6月月初倉庫: {int(cur)}")
for m in range(6,13):
    ri=m-1; ws_sim.cell(ri,3).value=round(cur)
    arr=n(ws_sim.cell(ri,4).value)
    nv=max(0,round(n(ws_sim.cell(ri,5).value)+dp[m])); ws_sim.cell(ri,5).value=nv
    end=max(0,cur+arr-nv); ws_sim.cell(ri,6).value=round(end); cur=end
    print(f"    {m}月: 入荷={int(arr)} 出荷={nv} 月末={int(end)}")
ws_sim.cell(12,5).value=round(sum(n(ws_sim.cell(r,5).value) for r in range(5,12)))
ws_sim.cell(12,6).value=round(cur)
print("  ✅ シミュレーション更新")

wb.save(FILE_V6)
print(f"\n  ✅ v6 保存完了: {os.path.basename(FILE_V6)}")

# Phase 3: backend
print("\n"+"="*50+"\nPhase 3: ダッシュボード更新\n"+"="*50)
with open(FILE_BACKEND,"r",encoding="utf-8") as f: bc=f.read()
if "v5.xlsx" in bc:
    bc=bc.replace("v5.xlsx","v6.xlsx")
    with open(FILE_BACKEND,"w",encoding="utf-8") as f: f.write(bc)
    print("  ✅ backend_calc.py: v5→v6")
else: print("  ℹ️ すでにv6参照")
result=subprocess.run([sys.executable,FILE_BACKEND],cwd=BASE,capture_output=True,text=True,encoding="utf-8",errors="replace")
for line in result.stdout.strip().split("\n"):
    if line.strip(): print(f"  > {line}")
if result.returncode!=0: print(f"  ⚠️ エラー:\n{result.stderr[:500]}")
else: print("  ✅ ダッシュボード更新完了")

print("\n"+"="*60+"\n  ✅ 全処理完了\n"+"="*60)
