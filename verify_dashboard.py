import pandas as pd
import json
import re
import sys

# Windows環境での文字化け対策
sys.stdout.reconfigure(encoding='utf-8')

def verify_dashboard_data():
    # 1. HTMLからINJECTED_DATAを抽出
    html_file = 'Kanken_Dashboard_最新版.html'
    try:
        with open(html_file, 'r', encoding='utf-8') as f:
            content = f.read()
    except Exception as e:
        print(f"HTMLファイルの読み込みに失敗しました: {e}")
        return

    match = re.search(r'const INJECTED_DATA = (\{.*?\});', content)
    if not match:
        print("HTML内に INJECTED_DATA が見つかりませんでした。")
        return
        
    injected = json.loads(match.group(1))
    
    # 2. Excelからデータを抽出
    excel_file = 'Kanken23510_2026年版_v6.xlsx'
    try:
        df = pd.read_excel(excel_file, sheet_name='全体サマリー', header=3)
    except Exception as e:
        print(f"Excelファイルの読み込みに失敗しました: {e}")
        return
        
    df.columns = df.columns.astype(str).str.replace('\n', '', regex=False).str.strip()
    df = df.dropna(subset=['カラー'])
    df = df[df['カラー'] != '合計'] # 合計行を除外
    
    # 3. KPIの計算と検証
    total_supply = int(df['総供給(6月〜)'].sum())
    total_demand = int(df['需要予測(6〜12月)'].sum())
    total_diff = int(df['過不足'].sum())
    
    print("=== 1. KPIの検証 ===")
    print(f"[HTML出力] 総供給={injected['kpi']['total_supply']}, 需要予測={injected['kpi']['total_demand']}, 過不足={injected['kpi']['total_diff']}")
    print(f"[元データ] 総供給={total_supply}, 需要予測={total_demand}, 過不足={total_diff}")
    
    kpi_match = (injected['kpi']['total_supply'] == total_supply and 
                 injected['kpi']['total_demand'] == total_demand and 
                 injected['kpi']['total_diff'] == total_diff)
    print("-> 判定: " + ("✅ 完全一致" if kpi_match else "❌ 不一致"))
        
    # 4. カテゴリの検証
    priority_excel = set(df[df['過不足'] >= 0]['カラー'].tolist())
    control_large_excel = set(df[df['過不足'] <= -500]['カラー'].tolist())
    control_small_excel = set(df[(df['過不足'] < 0) & (df['過不足'] > -500)]['カラー'].tolist())
    
    priority_html = set([item['name'] for item in injected['priority']])
    control_large_html = set([item['name'] for item in injected['control_large']])
    control_small_html = set([item['name'] for item in injected['control_small']])
    
    print("\n=== 2. カラーカテゴリ振り分けの検証 ===")
    p_diff = priority_excel ^ priority_html
    cl_diff = control_large_excel ^ control_large_html
    cs_diff = control_small_excel ^ control_small_html
    
    if not p_diff and not cl_diff and not cs_diff:
        print("-> 判定: ✅ 全カテゴリ（優先提案、コントロール大、コントロール小）の分類が完全に一致しています。")
    else:
        print("-> 判定: ❌ 不一致があります！")
        if p_diff: print(f"  - Priority差分: {p_diff}")
        if cl_diff: print(f"  - Control Large差分: {cl_diff}")
        if cs_diff: print(f"  - Control Small差分: {cs_diff}")

if __name__ == '__main__':
    verify_dashboard_data()
