import pandas as pd
import json
import os
import re
import glob

# ファイルパス（固定のもの）
FILE_KANKEN = r"C:\Users\kesuzuki\Desktop\シミュ２\Kanken23510_2026年版_v6.xlsx"
FILE_BUNPAI = r"C:\Users\kesuzuki\Desktop\シミュ２\分配店舗向け.xlsx"
OUTPUT_CSV  = r"C:\Users\kesuzuki\Desktop\シミュ２\検証用_統合データ.csv"
OUTPUT_JSON = r"C:\Users\kesuzuki\Desktop\シミュ２\dashboard_data.json"

def find_latest_sales_csv() -> str:
    """フォルダ内の *実績.csv を探し、更新日時が最新のファイルパスを返す。"""
    base_dir   = os.path.dirname(os.path.abspath(FILE_KANKEN))
    pattern    = os.path.join(base_dir, '*実績.csv')
    candidates = glob.glob(pattern)
    if not candidates:
        raise FileNotFoundError(
            "'*実績.csv' が見つかりません。フォルダに実績CSVを配置してください。"
        )
    return max(candidates, key=os.path.getmtime)

# 2024年の実績に基づく月別累積進捗率（係数）
# ※2026年も2024年と同様の季節パターン（売れ方のカーブ）になると仮定して年間予測を推計しています。
CUMULATIVE_PROGRESS = {
    1: 0.0437, 2: 0.1050, 3: 0.1970, 4: 0.3090, 5: 0.3968, 6: 0.4921,
    7: 0.5968, 8: 0.7067, 9: 0.7903, 10: 0.8677, 11: 0.9360, 12: 1.0000
}

def main():
    print("バックエンド処理（データ統合と計算）を開始します...")

    # --- 1. Kankenデータの読み込み ---
    try:
        if not os.path.exists(FILE_KANKEN):
            raise FileNotFoundError(f"最新のKankenデータファイルが見つかりません。\nパス: {FILE_KANKEN}")
        df_kanken = pd.read_excel(FILE_KANKEN, sheet_name="全体サマリー", header=3)
    except FileNotFoundError as e:
        print(f"\n【エラー】{e}")
        print("ファイル名や保存場所が正しいか確認してください。")
        return
    except ValueError as e:
        print(f"\n【エラー】Kankenデータファイル内に「全体サマリー」シートが見つかりません。")
        print(f"詳細エラー: {e}")
        return
    except Exception as e:
        print(f"\n【エラー】Kankenデータファイルの読み込み中に予期せぬエラーが発生しました。")
        print(f"詳細エラー: {e}")
        return
    
    # カラム名の改行文字を削除してからマッチングを行う
    df_kanken.columns = df_kanken.columns.astype(str).str.replace('\n', '', regex=False).str.strip()
    
    cols = df_kanken.columns.tolist()
    stock_col = next((c for c in cols if "倉庫WH" in c or "倉庫" in c), "6/1倉庫WH")
    
    stock_date = "6月1日" # デフォルト
    m = re.search(r"(\d{1,2}/\d{1,2})", stock_col)
    if m:
        parts = m.group(1).split('/')
        stock_date = f"{parts[0]}月{parts[1]}日"

    kanken_cols = {
        "カラー": "カラー名",
        stock_col: "倉庫在庫",
        "店舗在庫": "店舗在庫",
        "FW入荷": "FW入荷予定",
        "需要予測(6〜12月)": "残り需要予測"
    }
    
    available_cols = [c for c in kanken_cols.keys() if c in df_kanken.columns]
    df_kanken = df_kanken[available_cols].rename(columns=kanken_cols)
    df_kanken = df_kanken.dropna(subset=["カラー名"])
    df_kanken["カラー名"] = df_kanken["カラー名"].astype(str).str.strip()

    # --- 2. 実績CSVの自動検索・読み込みと実績月の判定 ---
    try:
        file_sales = find_latest_sales_csv()
        print(f"  使用する実績ファイル: {os.path.basename(file_sales)}")
    except FileNotFoundError as e:
        print(f"\n【エラー】{e}")
        return

    try:
        df_stores = pd.read_csv(file_sales, encoding="cp932")
    except UnicodeDecodeError:
        try:
            df_stores = pd.read_csv(file_sales, encoding="utf-8")
        except Exception as e:
            print(f"\n【エラー】実績データファイル（CSV）の読み込み中に文字コードエラーが発生しました。")
            print(f"詳細エラー: {e}")
            return
    except Exception as e:
        print(f"\n【エラー】実績データファイルの読み込み中に予期せぬエラーが発生しました。")
        print(f"詳細エラー: {e}")
        return

    # 列名のゆらぎ対応（数量 / 販売数）
    if '数量' not in df_stores.columns and '販売数' in df_stores.columns:
        df_stores = df_stores.rename(columns={'販売数': '数量'})
    # 列名のゆらぎ対応（商品コード / 行ラベル → 品番として使用）
    if '品番' not in df_stores.columns and '行ラベル' in df_stores.columns:
        df_stores = df_stores.rename(columns={'行ラベル': '品番'})
        
    # 月数の判定（欠落や異常値に備えた厳格化）
    if "MTH" in df_stores.columns:
        VALID_MONTHS = {'JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'}
        # 各月の売上合計を計算し、数量が0より大きい月のみ実績月数として数える（大文字に統一）
        df_stores['MTH_UPPER'] = df_stores['MTH'].dropna().astype(str).str.upper()
        monthly_sales = df_stores.groupby('MTH_UPPER')['数量'].sum()
        valid_active_months = [m for m, val in monthly_sales.items() if val > 0 and m in VALID_MONTHS]
        actual_months_count = len(valid_active_months)
    else:
        actual_months_count = 1
    
    # マジックナンバーを排し、1〜12の範囲に収める
    actual_months_count = max(1, min(12, actual_months_count))
        
    progress_rate = CUMULATIVE_PROGRESS.get(actual_months_count, 1.0)
    remain_months = max(0, 12 - actual_months_count)

    # --- 3. マスタと売上（分配店舗向け）の読み込み ---
    try:
        if not os.path.exists(FILE_BUNPAI):
            raise FileNotFoundError(f"分配店舗向けデータファイルが見つかりません。\nパス: {FILE_BUNPAI}")
        
        df_master = pd.read_excel(FILE_BUNPAI, sheet_name="master")
        df_master["カラー名"] = df_master["カラー名"].astype(str).str.strip()
        
        df_sales = pd.read_excel(FILE_BUNPAI, sheet_name="2026売上")
        df_sales["カラー名"] = df_sales["カラー名"].astype(str).str.strip()
        df_sales = df_sales.rename(columns={"販売数": "累計実績(販売数)"})
        df_sales = df_sales.groupby("カラー名", as_index=False)["累計実績(販売数)"].sum()
    except FileNotFoundError as e:
        print(f"\n【エラー】{e}")
        print("ファイル名や保存場所が正しいか確認してください。")
        return
    except ValueError as e:
        print(f"\n【エラー】分配店舗向けデータファイル内に必要なシート（「master」または「2026売上」）が見つかりません。")
        print(f"詳細エラー: {e}")
        return
    except Exception as e:
        print(f"\n【エラー】分配店舗向けデータファイルの読み込み中に予期せぬエラーが発生しました。")
        print(f"詳細エラー: {e}")
        return
    
    # --- 4. データの結合 ---
    df_merged = pd.merge(df_master, df_sales[["カラー名", "累計実績(販売数)"]], on="カラー名", how="left")
    df_merged = pd.merge(df_merged, df_kanken, on="カラー名", how="left")

    fill_cols = ["倉庫在庫", "店舗在庫", "FW入荷予定", "累計実績(販売数)", "残り需要予測"]
    for c in fill_cols:
        if c in df_merged.columns:
            df_merged[c] = pd.to_numeric(df_merged[c], errors='coerce').fillna(0).clip(lower=0).round().astype(int)

    # --- 5. 計算ロジック（ローリングフォーキャスト） ---
    df_merged["年間売上予測(計算)"] = (df_merged["累計実績(販売数)"] / progress_rate).round().astype(int)
    # ※Kankenファイル由来の需要予測は参考値として読み込んでいますが、
    # 最終的な計算には以下のローリングキャストで算出した「残り需要予測(現実ベース)」を使用します。
    df_merged["残り需要予測(現実ベース)"] = (df_merged["年間売上予測(計算)"] - df_merged["累計実績(販売数)"]).clip(lower=0).astype(int)
    
    df_merged["期首在庫"] = df_merged["倉庫在庫"] + df_merged["店舗在庫"]
    df_merged["年末着地見込み"] = df_merged["期首在庫"] + df_merged["FW入荷予定"] - df_merged["残り需要予測(現実ベース)"]
    df_merged["仕込ギャップ"] = df_merged["FW入荷予定"] - df_merged["残り需要予測(現実ベース)"]
    
    output_cols = [
        "品番", "カラー名", "累計実績(販売数)", "年間売上予測(計算)",
        "期首在庫", "残り需要予測(現実ベース)", "FW入荷予定", 
        "仕込ギャップ", "年末着地見込み"
    ]
    df_merged = df_merged[[c for c in output_cols if c in df_merged.columns]]

    # 実績も入荷も在庫もないゼロ行を除外 (画面表示をスッキリさせるため)
    df_merged = df_merged[(df_merged["累計実績(販売数)"] > 0) | (df_merged["FW入荷予定"] > 0) | (df_merged["期首在庫"] > 0)]

    # --- 6. 店舗別構成比（ベースシェア）の算出 ---
    df_stores["拠点"] = df_stores["拠点"].replace({
        "越一": "TOKYO NODE",
        "心斎橋パルコ": "大丸心斎橋",
        "FLAGS": "NARITA(FLAGS)"
    })
    
    df_stores["数量"] = pd.to_numeric(df_stores["数量"], errors='coerce').fillna(0).clip(lower=0)
    store_totals = df_stores.groupby("拠点")["数量"].sum()
    total_sales = store_totals.sum()
    
    base_shares = []
    if total_sales > 0:
        for store, qty in store_totals.items():
            base_shares.append({
                "store": store,
                "share": float(qty) / float(total_sales)
            })
    else:
        num_stores = len(store_totals)
        if num_stores > 0:
            for store in store_totals.index:
                base_shares.append({
                    "store": store,
                    "share": 1.0 / num_stores
                })
    base_shares = sorted(base_shares, key=lambda x: x["share"], reverse=True)

    # --- 7. CSVとJSの出力 ---
    df_merged.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    records = df_merged.to_dict(orient="records")
    
    output_data = {
        "summary": records,
        "base_shares": base_shares,
        "meta": {
            "stock_date": stock_date,
            "actual_months": int(actual_months_count),
            "remain_months": int(remain_months)
        }
    }
    
    OUTPUT_JS = r"C:\Users\kesuzuki\Desktop\シミュ２\dashboard_data.js"
    with open(OUTPUT_JS, "w", encoding="utf-8") as f:
        js_content = "const dashboardData = " + json.dumps(output_data, ensure_ascii=False, indent=2) + ";"
        f.write(js_content)

    print(f"処理完了: CSV -> {OUTPUT_CSV}")
    print(f"処理完了: JS  -> {OUTPUT_JS}")
    print(f"実績月数: {actual_months_count}ヶ月 (進捗率: {progress_rate*100:.2f}%)")

if __name__ == "__main__":
    main()
