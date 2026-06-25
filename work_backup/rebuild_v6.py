import os, sys, io, shutil, re, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

BASE = os.getcwd()
fmap = {f: os.path.join(BASE, f) for f in os.listdir(BASE)}
def fp(kw):
    for k,v in fmap.items():
        if kw in k: return v
    return None

FILE_V5 = fp("v5.xlsx")
FILE_V6 = FILE_V5.replace("v5.xlsx","v6.xlsx")

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

print("v5コピー→v6...")
shutil.copy(FILE_V5, FILE_V6)
wb=openpyxl.load_workbook(FILE_V6)
results={}

for cn,old_t in COLORS.items():
    if cn not in wb.sheetnames:
        print(f"⚠️ {cn} シート未発見"); continue
    ws=wb[cn]
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
        ri=m+5
        ws.cell(ri,6).value=round(nd[m])
        ws.cell(ri,8).value=np_[m]
        ws.cell(ri,9).value=round(np_[m]/ann*100,2) if ann>0 else 0
    ws.cell(18,6).value=round(d15+tnd); ws.cell(18,8).value=round(ann); ws.cell(18,9).value=100

    for i,(mo,st,ar,pl,en) in enumerate(whmov(info["wh0"],info["eb"],info["ab"],np_)):
        ri=23+i
        ws.cell(ri,3).value=st; ws.cell(ri,4).value=ar
        ws.cell(ri,5).value=pl; ws.cell(ri,6).value=en

    b19=ws["B19"].value
    if b19 is not None:
        df=supply-round(tnd)
        ws["B19"].value=(f"📦 供給 {supply:,}個　｜　需要(6〜12月) {round(tnd):,}個　｜　"+
                         (f"不足(区間比例配分でカバー): {abs(df):,}個" if isc else f"超過(需要追随): {df:,}個"))

    dc=round(tnd)-round(sum(od.values())); pc=tnp-round(sum(op[m] for m in range(6,13)))
    print(f"✅ {cn}: 需要{dc:+d}/計画{pc:+d} ({'制約' if isc else '追随'}) {old_t:.3f}→{NEW_TREND:.3f}")

    results[cn]={"supply":supply,"method":meth,
                 "odd":{m:od[m] for m in range(6,13)},"ndd":{m:round(nd[m]) for m in range(6,13)},
                 "odp":{m:op[m] for m in range(6,13)},"ndp":np_,
                 "ann":round(ann),"tnd":round(tnd),"isc":isc}

dd={m:int(round(sum(res["ndd"][m]-res["odd"][m] for res in results.values()))) for m in range(6,13)}
dp={m:int(round(sum(res["ndp"][m]-res["odp"][m] for res in results.values()))) for m in range(6,13)}
print("\n月別 計画デルタ:")
for m in range(6,13): print(f"  {m}月: {dp[m]:+d}")

ws_s=wb["全体サマリー"]
for cn,res in results.items():
    tgt=None
    for ri in range(5,35):
        if ws_s.cell(ri,3).value and str(ws_s.cell(ri,3).value).strip()==cn: tgt=ri; break
    if tgt is None: print(f"⚠️ {cn} サマリー未発見"); continue
    ws_s.cell(tgt,9).value=res["tnd"]; ws_s.cell(tgt,10).value=res["method"]
    ws_s.cell(tgt,11).value=res["ann"]; ws_s.cell(tgt,12).value=res["supply"]-res["tnd"]
for col in [9,11,12]: ws_s.cell(32,col).value=int(round(sum(n(ws_s.cell(r,col).value) for r in range(5,32))))
print("✅ 全体サマリー更新")

ws_f=wb["2026年着地予測"]
cumul=n(ws_f.cell(9,8).value)
for m in range(6,13):
    ri=m+4
    ws_f.cell(ri,4).value=int(round(n(ws_f.cell(ri,4).value)+dd[m]))
    nv=int(round(n(ws_f.cell(ri,6).value)+dp[m])); ws_f.cell(ri,6).value=nv
    cumul+=nv; ws_f.cell(ri,8).value=int(round(cumul))
td_=int(round(sum(n(ws_f.cell(r,4).value) for r in range(5,17))))
tp_=int(round(sum(n(ws_f.cell(r,6).value) for r in range(5,17))))
ws_f.cell(17,4).value=td_; ws_f.cell(17,6).value=tp_; ws_f.cell(17,7).value=100; ws_f.cell(17,8).value=tp_
for ri in range(5,17): ws_f.cell(ri,7).value=round(n(ws_f.cell(ri,6).value)/tp_*100,2) if tp_ else 0
print(f"✅ 着地予測更新: 需要={td_:,} 計画={tp_:,}")

ws_sim=wb["供給需要シミュレーション"]
cur=n(ws_sim.cell(5,3).value)
for m in range(6,13):
    ri=m-1; ws_sim.cell(ri,3).value=int(round(cur))
    arr=n(ws_sim.cell(ri,4).value)
    nv=max(0,int(round(n(ws_sim.cell(ri,5).value)+dp[m]))); ws_sim.cell(ri,5).value=nv
    end=max(0,cur+arr-nv); ws_sim.cell(ri,6).value=int(round(end)); cur=end
ws_sim.cell(12,5).value=int(round(sum(n(ws_sim.cell(r,5).value) for r in range(5,12))))
ws_sim.cell(12,6).value=int(round(cur))
print("✅ シミュレーション更新")

ws_l=wb["計算ロジック説明"]
for ri in range(1,60):
    for ci in range(1,12):
        v=ws_l.cell(ri,ci).value
        if isinstance(v,str) and ("2.2222" in v or "2.222" in v):
            ws_l.cell(ri,ci).value=v.replace("2.2222","1.630").replace("2.222","1.630")
print("✅ 計算ロジック説明更新")

wb.save(FILE_V6)
print(f"\n✅ v6保存完了: {os.path.basename(FILE_V6)}")

print("\n=== 最終検証 ===")
wb6=openpyxl.load_workbook(FILE_V6)
for cn in COLORS:
    if cn not in wb6.sheetnames: continue
    ws=wb6[cn]
    b2=ws["B2"].value or ""
    tm=re.search(r"トレンド比:\s*([\d.]+)",b2)
    tr=tm.group(1) if tm else "?"
    d6=int(round(sum(n(ws.cell(m+5,6).value) for m in range(6,13))))
    p6=int(round(sum(n(ws.cell(m+5,8).value) for m in range(6,13))))
    ok="✅" if tr==str(NEW_TREND) else "⚠️"
    print(f"  {ok} {cn}: トレンド={tr} 需要={d6} 計画={p6}")
