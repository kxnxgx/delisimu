import openpyxl
import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')

# ==============================================================
# 設定エリア: C3/C4の変更前後の値
# ==============================================================
C3_OLD = 4535   # 変更前の2025年1-5月実績
C4_OLD = 7692   # 変更前の2025年6-12月実績
C3_NEW = 5490   # 変更後の2025年1-5月実績（現在の値）
C4_NEW = 7411   # 変更後の2025年6-12月実績（現在の値）
G3 = 7389       # 2026年1-5月実績（不変）
# ==============================================================

# G4の計算（分配店舗向け 昨対比シートの計算式: G3 / D3 × D4）
G4_OLD = G3 / (C3_OLD / (C3_OLD + C4_OLD)) * (C4_OLD / (C3_OLD + C4_OLD))
G4_NEW = G3 / (C3_NEW / (C3_NEW + C4_NEW)) * (C4_NEW / (C3_NEW + C4_NEW))
SCALE = G4_NEW / G4_OLD

print("=" * 55)
print("【変更内容の確認】")
print("=" * 55)
print(f"  C3 (2025年1-5月実績): {C3_OLD:,} → {C3_NEW:,}")
print(f"  C4 (2025年6-12月実績): {C4_OLD:,} → {C4_NEW:,}")
print()
print(f"  変更前 1-5月構成比: {C3_OLD/(C3_OLD+C4_OLD):.4f}")
print(f"  変更後 1-5月構成比: {C3_NEW/(C3_NEW+C4_NEW):.4f}")
print()
print(f"  変更前 G4 (全体6-12月需要見込み): {G4_OLD:,.0f}")
print(f"  変更後 G4 (全体6-12月需要見込み): {G4_NEW:,.0f}")
print(f"  スケール係数: {SCALE:.4f} ({(SCALE-1)*100:+.1f}%)")
print()

# V6の全体サマリーシートを読み込んで需要予測を更新
# 行番号マッピングのため pandas で先に読み込み
df = pd.read_excel('Kanken23510_2026年版_v6.xlsx', sheet_name='全体サマリー', header=3)
df.columns = df.columns.astype(str).str.replace('\n', '', regex=False).str.strip()
df = df.dropna(subset=['カラー'])
df = df[df['カラー'] != '合計']

# openpyxl でV6を開く（書式を壊さないよう data_only=False で開く）
wb = openpyxl.load_workbook('Kanken23510_2026年版_v6.xlsx')
ws = wb['全体サマリー']

print("=" * 55)
print("【V6 全体サマリー 需要予測更新プレビュー】")
print("=" * 55)
print(f"{'カラー':30} {'旧需要':>6} {'新需要':>6} {'旧過不足':>8} {'新過不足':>8}")
print("-" * 65)

# ヘッダーは4行目（row=4）なので、データは5行目から
# 列マッピング (openpyxlは1始まり): 
#   B=2(#), C=3(カラー), D=4(倉庫WH), E=5(店舗在庫), F=6(FW入荷)
#   G=7(2025入荷), H=8(総供給), I=9(需要予測), J=10(配分方式), K=11(年間出荷計画), L=12(過不足)

updated_colors = []
for row in ws.iter_rows(min_row=5, max_row=ws.max_row):
    color_cell = row[2]   # C列 = カラー
    supply_cell = row[7]  # H列 = 総供給(6月〜)
    demand_cell = row[8]  # I列 = 需要予測(6〜12月)
    diff_cell   = row[11] # L列 = 過不足

    if color_cell.value is None or str(color_cell.value).strip() in ('', '合計'):
        continue
    if demand_cell.value is None or not isinstance(demand_cell.value, (int, float)):
        continue

    old_demand = float(demand_cell.value)
    new_demand = round(old_demand * SCALE)
    supply = float(supply_cell.value) if supply_cell.value else 0
    old_diff = float(diff_cell.value) if diff_cell.value else 0
    new_diff = supply - new_demand

    print(f"{str(color_cell.value):30} {old_demand:>6.0f} {new_demand:>6.0f} {old_diff:>8.0f} {new_diff:>8.0f}")

    # 書き込み
    demand_cell.value = new_demand
    diff_cell.value = new_diff
    updated_colors.append(color_cell.value)

total_new_demand = sum(
    row[8].value for row in ws.iter_rows(min_row=5, max_row=ws.max_row)
    if row[2].value and str(row[2].value).strip() not in ('', '合計')
    and isinstance(row[8].value, (int, float))
)

print("-" * 65)
print(f"{'合計':30} {15362:>6.0f} {total_new_demand:>6.0f}")
print()
print(f"✅ {len(updated_colors)} 色を更新しました")

wb.save('Kanken23510_2026年版_v6.xlsx')
print("✅ Kanken23510_2026年版_v6.xlsx を上書き保存しました")
