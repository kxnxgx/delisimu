import sys, os, glob
import openpyxl
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")

EXCEL_FILE  = "Kanken23510_2026\u5e74\u7248_v6.xlsx"
STOCK_CSV   = "\u5728\u5eab\u4e00\u89a7.csv"
BUNPAI_XLSX = "\u5206\u914d\u5e97\u8217\u5411\u3051.xlsx"

WAREHOUSE_LOCATIONS = {"\u30d0\u30eb\u30af", "New Way-A"}
EXCLUDE_LOCATIONS   = {"New Way-B", "New Way-C", "New Way-G"}

CUMULATIVE_PROGRESS = {
    1: 0.0437, 2: 0.1050, 3: 0.1970, 4: 0.3090, 5: 0.3968, 6: 0.4921,
    7: 0.5968, 8: 0.7067, 9: 0.7903, 10: 0.8677, 11: 0.9360, 12: 1.0000
}

COL_COLOR=3; COL_WH=4; COL_STORE=5; COL_FW=6
COL_SUPPLY=8; COL_DEMAND=9; COL_DIFF=12
DATA_START=5; TOTAL_LABEL="\u5408\u8a08"


def read_csv_auto(path):
    for enc in ("cp932","utf-8-sig","utf-8","shift_jis"):
        try:
            return pd.read_csv(path, encoding=enc)
        except (UnicodeDecodeError, LookupError):
            continue
    raise ValueError(f"\u30a8\u30f3\u30b3\u30fc\u30c9\u5931\u6557: {path}")


def find_latest_sales_csv():
    base = os.path.dirname(os.path.abspath(__file__))
    files = glob.glob(os.path.join(base, "*\u5b9f\u7e3e.csv"))
    if not files:
        raise FileNotFoundError("*\u5b9f\u7e3e.csv \u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093")
    return max(files, key=os.path.getmtime)


def load_master(path):
    df = pd.read_excel(path, sheet_name="master")
    df["\u54c1\u756a"]    = df["\u54c1\u756a"].astype(str).str.strip()
    df["\u30ab\u30e9\u30fc\u540d"] = df["\u30ab\u30e9\u30fc\u540d"].astype(str).str.strip()
    return dict(zip(df["\u54c1\u756a"], df["\u30ab\u30e9\u30fc\u540d"]))


def load_stock(path, master):
    df = read_csv_auto(path)
    df["\u5546\u54c1\u30b3\u30fc\u30c9"] = df["\u5546\u54c1\u30b3\u30fc\u30c9"].astype(str).str.strip()
    df["\u73fe\u5728\u6570\u91cf"] = pd.to_numeric(df["\u73fe\u5728\u6570\u91cf"], errors="coerce").fillna(0).clip(lower=0)
    df["\u30ab\u30e9\u30fc\u540d"] = df["\u5546\u54c1\u30b3\u30fc\u30c9"].map(master)
    unmapped = df[df["\u30ab\u30e9\u30fc\u540d"].isna()]["\u5546\u54c1\u30b3\u30fc\u30c9"].unique()
    if len(unmapped) > 0:
        print(f"  [\u60c5\u5831] \u5728\u5eab\u672a\u63d2\u5165\uff1a {unmapped.tolist()[:10]}")
    df = df.dropna(subset=["\u30ab\u30e9\u30fc\u540d"])
    df = df[~df["\u62e0\u70b9\u540d"].isin(EXCLUDE_LOCATIONS)]
    df["\u533a\u5206"] = df["\u62e0\u70b9\u540d"].apply(lambda x: "\u5009\u5eab" if x in WAREHOUSE_LOCATIONS else "\u5e97\u8217")
    wh    = df[df["\u533a\u5206"]=="\u5009\u5eab"].groupby("\u30ab\u30e9\u30fc\u540d")["\u73fe\u5728\u6570\u91cf"].sum().astype(int).to_dict()
    store = df[df["\u533a\u5206"]=="\u5e97\u8217"].groupby("\u30ab\u30e9\u30fc\u540d")["\u73fe\u5728\u6570\u91cf"].sum().astype(int).to_dict()
    return wh, store


def load_sales_actual(sales_csv, bunpai_xlsx, master):
    try:
        ds = pd.read_excel(bunpai_xlsx, sheet_name="2026\u58f2\u4e0a")
        ds["\u30ab\u30e9\u30fc\u540d"] = ds["\u30ab\u30e9\u30fc\u540d"].astype(str).str.strip()
        ds["\u8ca9\u58f2\u6570"] = pd.to_numeric(ds["\u8ca9\u58f2\u6570"], errors="coerce").fillna(0).clip(lower=0)
        sbc = ds.groupby("\u30ab\u30e9\u30fc\u540d")["\u8ca9\u58f2\u6570"].sum().astype(int).to_dict()
        print(f"      [\u30bd\u30fc\u30b91] 2026\u58f2\u4e0a: {len(sbc)} \u30ab\u30e9\u30fc, \u5408\u8a08 {sum(sbc.values()):,}\u500b")
    except Exception as e:
        print(f"      [\u8b66\u544a] 2026\u58f2\u4e0a\u8aad\u8fbc\u307f\u5931\u6557: {e}")
        sbc = {}

    dr = read_csv_auto(sales_csv)
    if "\u6570\u91cf" not in dr.columns and "\u8ca9\u58f2\u6570" in dr.columns:
        dr = dr.rename(columns={"\u8ca9\u58f2\u6570": "\u6570\u91cf"})
    qty = next((c for c in ["\u6570\u91cf","\u8ca9\u58f2\u6570","Qty","qty"] if c in dr.columns), None)
    if qty is None:
        return sbc, 1
    dr[qty] = pd.to_numeric(dr[qty], errors="coerce").fillna(0).clip(lower=0)

    item_col = next((c for c in ["3rd Item No.","\u54c1\u756a","\u5546\u54c1\u30b3\u30fc\u30c9","\u884c\u30e9\u30d9\u30eb"] if c in dr.columns), None)
    if item_col:
        dr["key"] = dr[item_col].astype(str).str.strip()
        dr["cn"]  = dr["key"].map(master)
        dk = dr.dropna(subset=["cn"])
    else:
        dk = dr

    if not sbc and item_col:
        sbc = dk.groupby("cn")[qty].sum().astype(int).to_dict()
        print(f"      [\u30bd\u30fc\u30b92(FB)]: {len(sbc)} \u30ab\u30e9\u30fc, {sum(sbc.values()):,}\u500b")

    if "\u55b6\u696d\u65e5\u4ed8" in dk.columns:
        tmp = dk.copy()
        tmp["dt"] = pd.to_datetime(tmp["\u55b6\u696d\u65e5\u4ed8"], errors="coerce")
        tmp["m"]  = tmp["dt"].dt.month
        mq = tmp.groupby("m")[qty].sum()
        vm = [m for m, v in mq.items() if v > 0 and pd.notna(m)]
        mc = len(vm)
        print(f"      \u5b9f\u7e3e\u6708\uff1a {sorted([int(m) for m in vm])}")
    else:
        mc = 1

    return sbc, max(1, min(12, mc))


def calc_demand(sbc, months_count, ws):
    pr = CUMULATIVE_PROGRESS.get(months_count, 1.0)
    old_total = sum(
        float(row[COL_DEMAND-1].value or 0)
        for row in ws.iter_rows(min_row=DATA_START, max_row=ws.max_row)
        if str(row[COL_COLOR-1].value or "").strip() not in ("", TOTAL_LABEL)
        and isinstance(row[COL_DEMAND-1].value, (int, float))
    )
    ta = sum(sbc.values())
    if pr > 0 and ta > 0 and old_total > 0:
        h2 = 1.0 - CUMULATIVE_PROGRESS.get(months_count, 0)
        scale = (ta / pr * h2) / old_total
    else:
        scale = None

    out = {}
    for row in ws.iter_rows(min_row=DATA_START, max_row=ws.max_row):
        cv = str(row[COL_COLOR-1].value or "").strip()
        if not cv or cv == TOTAL_LABEL:
            continue
        od = float(row[COL_DEMAND-1].value or 0)
        if cv in sbc and pr > 0:
            ca = sbc[cv]
            out[cv] = round(max(0.0, ca / pr - ca))
        elif scale is not None:
            out[cv] = round(od * scale)
        else:
            out[cv] = round(od)
    return out


def main():
    print("=" * 65)
    print("  build_summary.py -- \u5168\u4f53\u30b5\u30de\u30ea\u30fc\u518d\u69cb\u7bc9\u30b9\u30af\u30ea\u30d7\u30c8")
    print("=" * 65)

    print("\n[1/5] \u30de\u30b9\u30bf\u30c7\u30fc\u30bf\u8aad\u307f\u8fbc\u307f...")
    master = load_master(BUNPAI_XLSX)
    print(f"      {len(master)} \u54c1\u756a")

    print(f"\n[2/5] \u5728\u5eab\u4e00\u89a7 ({STOCK_CSV}) \u8aad\u307f\u8fbc\u307f...")
    wh_stock, store_stock = load_stock(STOCK_CSV, master)
    print(f"      \u5009\u5eabWH : {len(wh_stock)} \u30ab\u30e9\u30fc / \u5408\u8a08 {sum(wh_stock.values()):,}\u500b")
    print(f"      \u5e97\u8217\u5728\u5eab: {len(store_stock)} \u30ab\u30e9\u30fc / \u5408\u8a08 {sum(store_stock.values()):,}\u500b")

    print(f"\n[3/5] \u5b9f\u7e3e\u30c7\u30fc\u30bf\u8aad\u307f\u8fbc\u307f...")
    csv = find_latest_sales_csv()
    print(f"      \u5b9f\u7e3e\u30d5\u30a1\u30a4\u30eb: {os.path.basename(csv)}")
    sbc, months = load_sales_actual(csv, BUNPAI_XLSX, master)
    pr = CUMULATIVE_PROGRESS.get(months, 1.0)
    print(f"      \u5b9f\u7e3e\u6708\u6570: {months} \u30f6\u6708 (\u9032\u6357\u7387: {pr*100:.2f}%)")
    print(f"      Kanken\u7d2f\u8a08\u5b9f\u7e3e: {sum(sbc.values()):,}\u500b")

    print(f"\n[4/5] v6\u30d5\u30a1\u30a4\u30eb\u8aad\u8fbc\u307f...")
    if not os.path.exists(EXCEL_FILE):
        raise FileNotFoundError(f"'{EXCEL_FILE}' \u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093")
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["\u5168\u4f53\u30b5\u30de\u30ea\u30fc"]
    demand = calc_demand(sbc, months, ws)
    print(f"      \u9700\u8981\u4e88\u6e2c\u5408\u8a08(\u65b0): {sum(demand.values()):,}\u500b")

    print(f"\n[5/5] \u66f8\u304d\u8fbc\u307f\u4e2d...")
    updated = td = te = th = ti = tl = 0

    for row in ws.iter_rows(min_row=DATA_START, max_row=ws.max_row):
        cv = str(row[COL_COLOR-1].value or "").strip()
        if not cv or cv == TOTAL_LABEL:
            continue
        wc = row[COL_WH-1];    sc = row[COL_STORE-1]
        fc = row[COL_FW-1];    hc = row[COL_SUPPLY-1]
        ic = row[COL_DEMAND-1]; lc = row[COL_DIFF-1]

        ow = int(wc.value) if isinstance(wc.value,(int,float)) else 0
        os_ = int(sc.value) if isinstance(sc.value,(int,float)) else 0
        fv = int(fc.value) if isinstance(fc.value,(int,float)) else 0
        oh = int(hc.value) if isinstance(hc.value,(int,float)) else 0
        oi = int(ic.value) if isinstance(ic.value,(int,float)) else 0

        nw = wh_stock.get(cv, ow)
        ns = store_stock.get(cv, os_)
        nh = nw + ns + fv     # H = D + E + F (\u30d0\u30b0\u4fee\u6b63\u306e\u6838\u5fc3)
        ni = demand.get(cv, oi)
        nl = nh - ni

        wc.value = nw; sc.value = ns; hc.value = nh; ic.value = ni; lc.value = nl

        td += nw; te += ns; th += nh; ti += ni; tl += nl; updated += 1
        print(f"  {cv:<32} wh:{ow}->{nw} st:{os_}->{ns} fw:{fv} sup:{oh}->{nh} dem:{oi}->{ni} diff:{nl:+}")

    # \u5408\u8a08\u884c
    for row in ws.iter_rows(min_row=DATA_START, max_row=ws.max_row):
        if str(row[COL_COLOR-1].value or "").strip() == TOTAL_LABEL:
            row[COL_WH-1].value = td; row[COL_STORE-1].value = te
            row[COL_SUPPLY-1].value = th; row[COL_DEMAND-1].value = ti
            row[COL_DIFF-1].value = tl
            print(f"  \u5408\u8a08\u884c: D={td} E={te} H={th} I={ti} L={tl:+}")
            break

    wb.save(EXCEL_FILE)
    print(f"\n\u2705 {updated} \u30ab\u30e9\u30fc\u3092\u66f4\u65b0, \u4fdd\u5b58\u5b8c\u4e86")
    print(f"  \u7dcf\u4f9b\u7d66(H): {th:,}  \u9700\u8981\u4e88\u6e2c(I): {ti:,}  \u904e\u4e0d\u8db3(L): {tl:+,}")
    print("=" * 65)


if __name__ == "__main__":
    main()
