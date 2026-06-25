import os, sys, io, shutil, csv, re, subprocess, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

BASE = os.getcwd()
fmap = {f: os.path.join(BASE, f) for f in os.listdir(BASE)}

def fp(kw):
    for k,v in fmap.items():
        if kw in k: return v
    return None

FILE_V5      = fp("v5.xlsx")
FILE_V6      = FILE_V5.replace("v5.xlsx","v6.xlsx")
FILE_JISSEKI = fp("6.24")
FILE_CSV     = fp("2026実績.csv")
FILE_BACKEND = fp("backend_calc.py")

NEW_TREND = 1.63
COLORS = {
    "Corn":1.775,"Pastel Lavender-Confetti Pat":4.0,
    "Chalk Rose":2.222,"Green":2.222,"Black Oak":2.222,
    "Teal":2.222,"Fossil":2.222,"Midnight Purple":2.222,
}
SMAP = {
    "FJALLRAVEN by 3NITY TOKYO":"TOKYO",
    "FJALLRAVEN by 3NITY ルクア大阪":"ルクア大阪",
    "FJALLRAVEN POPUP NARITA":"FLAGS",
    "FJALLRAVEN by 3NITY 京王新宿":"京王新宿",
    "FJALLRAVEN by 3NITY 大丸心斎橋":"大丸心斎橋",
    "FJALLRAVEN by 3NITY玉川高島屋S・C":"玉川高島屋",
    "FJALLRAVEN STORE 名古屋ファッションワン":"名古屋",
    "FJALLRAVEN by 3NITY SAPPORO HUTTE":"HUTTE",
    "FJALLRAVEN by 3NITY POPUP":"OIOI",
}

def n(val,d=0):
    if isinstance(val,(int,float)): return float(val)
    try: return float(str(val).replace(",",""))
    except: return float(d)

def parse_b2(text):
    r={"wh0":0,"eb":0,"ab":0,"supply":0}
    if not text: return r
    for key,pat in [("wh0",r"6/1倉庫\(BULK\+NWA\):\s*(\d+)"),("eb",r"EB:(\d+)"),("ab",r"AB:(\d+)"),("supply",r"総供給:\s*(\d+)")]:
        m=re.search(pat,text)
        if m: r[key]=int(m.group(1))
    return r

def gpp(dd,supply):
    ms=sorted(dd); td=sum(dd.values())
    if td<=0: return {m:0 for m in ms}
    if supply>=td: return {m:round(v) for m,v in dd.items()}
    plan,alloc={},0
    for i,m in enumerate(ms):
        if i==len(ms)-1: plan[m]=max(0,supply-alloc)
        else: p=round(dd[m]/td*supply); plan[m]=p; alloc+=p
    return plan

def whmov(wh0,eb,ab,pd):
    fw={7:eb,9:ab}; rows,cur=[],wh0
    for m in range(6,13):
        arr=fw.get(m,0); p=pd.get(m,0); end=max(0,cur+arr-p)
        rows.append((m,cur,arr,p,end)); cur=end
    return rows

# ===== Phase 1: 6月実績 =====
print("="*50+"\nPhase 1: 6月実績 → 2026実績.csv\n"+"="*50)
wb_j=openpyxl.load_workbook(FILE_JISSEKI)
ws_j=wb_j["Sheet2"]
new_rows,skipped=[],[]
for row in ws_j.iter_rows(min_row=2,max_row=ws_j.max_row,values_only=True):
    sr=str(row[0]).strip() if row[0] is not None else ""
    sku=str(row[1]).strip() if row[1] is not None else ""
    hinmei=str(row[3]).strip() if row[3] is not None else "Kanken"
    cr=row[4]; qty=n(row[6])
    if qty<=0: continue
    if sr in SMAP: tenpo=SMAP[sr]
    else:
        try: tenpo=str(int(float(sr)))
        except: skipped.append(sr); continue
    cc=str(int(cr)) if isinstance(cr,float) else (str(cr).strip() if cr else "")
    new_rows.append({"YEAR":"2026","MTH":"JUN","拠点":tenpo,"商品コード":sku,"数量":str(int(qty)),"販売合計":"","BRAND":"FRV","品番":"23510","品名":hinmei,"カラー":cc,"サイズ":"X"})
if skipped: print(f"  スキップ: {list(set(skipped))}")
print(f"  追加: {len(new_rows)}件 / {int(sum(n(r['数量']) for r in new_rows))}個")
existing,fieldnames=[],None
for enc in ["utf-8","utf-8-sig","cp932"]:
    try:
        with open(FILE_CSV,encoding=enc,newline="") as f:
            reader=csv.DictReader(f); fieldnames=reader.fieldnames; existing=list(reader)
        break
    except: continue
# BOM除去して正規化
fn_clean=[f.lstrip("\ufeff") for f in fieldnames]
existing_clean=[{k.lstrip("\ufeff"):v for k,v in r.items()} for r in existing]
all_rows=existing_clean+new_rows
with open(FILE_CSV,"w",encoding="utf-8-sig",newline="") as f:
    w=csv.DictWriter(f,fieldnames=fn_clean); w.writeheader(); w.writerows(all_rows)
print(f"  ✅ {len(existing)}+{len(new_rows)}={len(all_rows)}件")

# ===== Phase 2: v6 作成 =====
print("\n"+"="*50+"\nPhase 2: v6 Excel 作成\n"+"="*50)
shutil.copy(FILE_V5,FILE_V6); print("  コピー: v5→v6")
wb=openpyxl.load_workbook(FILE_V6)
results={}

def upd(ws,cn,old_t):
    scale=NEW_TREND/old_t
    b2o=ws["B2"].value or ""
    b2n=re.sub(r"(トレンド比:\s*)[\d.]+",rf"\g<1>{NEW_TREND:.3f}",b2o)
    b2n=re.sub(r"（[^）]*）","（全体163%ベース統一）",b2n,count=1)
    ws["B2"].value=b2n
    info=parse_b2(b2o); supply=info["supply"]
    od,op={},{}
    for ri in range(6,18): m=ri-5; od[m]=n(ws.cell(ri,6).value); op[m]=n(ws.cell(ri,8).value)
    nd={m:od[m]*scale for m in range(6,13)}
    tnd=sum(nd.values()); isc=supply<tnd
    np_=gpp(nd,supply) if isc else {m:round(nd[m]) for m in range(6,13)}
    meth="区間比例配分" if isc else "需要追随（制約なし）"
    p15=sum(op[m] for m in range(1,6)); tnp=sum(np_.values()); ann=p15+tnp
    d15=sum(od[m] for m in range(1,6))
    for m in range(6,13):
        ri=m+5; ws.cell(ri,6).value=round(nd[m]); ws.cell(ri,8).value=np_[m]
        ws.cell(ri,9).value=round(np_[m]/ann*100,2) if ann>0 else 0
    ws.cell(18,6).value=round(d15+tnd); ws.cell(18,8).value=round(ann); ws.cell(18,9).value=100
    for i,(mo,st,ar,pl,en) in enumerate(whmov(info["wh0"],info["eb"],info["ab"],np_)):
        ri=23+i; ws.cell(ri,3).value=st; ws.cell(ri,4).value=ar; ws.cell(ri,5).value=pl; ws.cell(ri,6).value=en
    b19=ws["B19"].value
    if b19 is not None:
        df=supply-round(tnd)
        ws["B19"].value=(f"📦 供給 {supply:,}個　｜　需要(6〜12月) {round(tnd):,}個　｜　"+
                         (f"不足(区間比例配分でカバー): {abs(df):,}個" if isc else f"超過(需要追随): {df:,}個"))
    return {"supply":supply,"method":meth,
            "odd":{m:od[m] for m in range(6,13)},"ndd":{m:round(nd[m]) for m in range(6,13)},
            "odp":{m:op[m] for m in range(6,13)},"ndp":np_,
            "ann":round(ann),"tnd":round(tnd),"isc":isc}

for cn,ot in COLORS.items():
    if cn not in wb.sheetnames: print(f"  ⚠️ {cn} シート未発見"); continue
    res=upd(wb[cn],cn,ot); results[cn]=res
    dc=res["tnd"]-round(sum(res["odd"].values()))
    pc=sum(res["ndp"].values())-round(sum(res["odp"].values()))
    print(f"  ✅ {cn}: 需要{dc:+d}/計画{pc:+d} ({'制約' if res['isc'] else '追随'})")

dd={m:sum(res["ndd"][m]-res["odd"][m] for res in results.values()) for m in range(6,13)}
dp={m:sum(res["ndp"][m]-res["odp"][m] for res in results.values()) for m in range(6,13)}
print("\n  月別 計画デルタ:")
for m in range(6,13): print(f"    {m}月: {dp[m]:+d}")

# 全体サマリー
ws_s=wb["全体サマリー"]
for cn,res in results.items():
    tgt=None
    for ri in range(5,35):
        if ws_s.cell(ri,3).value and str(ws_s.cell(ri,3).value).strip()==cn: tgt=ri; break
    if tgt is None: print(f"  ⚠️ {cn} サマリー未発見"); continue
    ws_s.cell(tgt,9).value=res["tnd"]; ws_s.cell(tgt,10).value=res["method"]
    ws_s.cell(tgt,11).value=res["ann"]; ws_s.cell(tgt,12).value=res["supply"]-res["tnd"]
for col in [9,11,12]: ws_s.cell(32,col).value=round(sum(n(ws_s.cell(r,col).value) for r in range(5,32)))
print("  ✅ 全体サマリー更新")

# 2026年着地予測
ws_f=wb["2026年着地予測"]
cumul=n(ws_f.cell(9,8).value)
for m in range(6,13):
    ri=m+4
    ws_f.cell(ri,4).value=round(n(ws_f.cell(ri,4).value)+dd[m])
    nv=round(n(ws_f.cell(ri,6).value)+dp[m]); ws_f.cell(ri,6).value=nv
    cumul+=nv; ws_f.cell(ri,8).value=round(cumul)
td_=round(sum(n(ws_f.cell(r,4).value) for r in range(5,17)))
tp_=round(sum(n(ws_f.cell(r,6).value) for r in range(5,17)))
ws_f.cell(17,4).value=td_; ws_f.cell(17,6).value=tp_; ws_f.cell(17,7).value=100; ws_f.cell(17,8).value=tp_
for ri in range(5,17): ws_f.cell(ri,7).value=round(n(ws_f.cell(ri,6).value)/tp_*100,2) if tp_ else 0
print("  ✅ 2026年着地予測更新")

# 供給需要シミュレーション
ws_sim=wb["供給需要シミュレーション"]
cur=n(ws_sim.cell(5,3).value)
for m in range(6,13):
    ri=m-1; ws_sim.cell(ri,3).value=round(cur)
    arr=n(ws_sim.cell(ri,4).value)
    nv=max(0,round(n(ws_sim.cell(ri,5).value)+dp[m])); ws_sim.cell(ri,5).value=nv
    end=max(0,cur+arr-nv); ws_sim.cell(ri,6).value=round(end); cur=end
ws_sim.cell(12,5).value=round(sum(n(ws_sim.cell(r,5).value) for r in range(5,12)))
ws_sim.cell(12,6).value=round(cur)
print("  ✅ 供給需要シミュレーション更新")

# 計算ロジック説明
ws_l=wb["計算ロジック説明"]
for ri in range(1,60):
    for ci in range(1,12):
        v=ws_l.cell(ri,ci).value
        if isinstance(v,str) and ("2.2222" in v or "2.222" in v):
            ws_l.cell(ri,ci).value=v.replace("2.2222","1.630").replace("2.222","1.630")
print("  ✅ 計算ロジック説明更新")

wb.save(FILE_V6)
print(f"\n  ✅ v6 保存完了")

# ===== Phase 3: ダッシュボード =====
print("\n"+"="*50+"\nPhase 3: ダッシュボード更新\n"+"="*50)
with open(FILE_BACKEND,"r",encoding="utf-8") as f: bc=f.read()
if "v5.xlsx" in bc:
    bc=bc.replace("v5.xlsx","v6.xlsx")
    with open(FILE_BACKEND,"w",encoding="utf-8") as f: f.write(bc)
    print("  ✅ backend_calc.py: v5→v6")
result=subprocess.run([sys.executable,FILE_BACKEND],cwd=BASE,capture_output=True,text=True,encoding="utf-8",errors="replace")
for line in result.stdout.strip().split("\n"):
    if line.strip(): print(f"  > {line}")
if result.returncode!=0: print(f"  ⚠️ エラー:\n{result.stderr[:500]}")
else: print("  ✅ ダッシュボード更新完了")

print("\n"+"="*60+"\n  ✅ 全処理完了\n"+"="*60)
