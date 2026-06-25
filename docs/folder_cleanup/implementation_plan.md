# フォルダ整理の実装計画（Folder Cleanup Plan）

現在のフォルダ内にあるファイルを整理し、本番の運用に必要なファイルだけを残してすっきりとした状態にします。

## ユーザー確認・承認が必要な事項

> [!IMPORTANT]
> 散らかっている一時ファイル（Pythonスクリプトやテキストログなど）の扱いについて、以下のいずれの方法が良いかご指示ください。
> 
> - **方法A（推奨）**: フォルダ内に「`作業用バックアップ`（`work_backup`）」という新しいフォルダを作成し、一時的なファイルをすべてその中に移動して退避させる（後から確認が必要になった場合も安心です）。
> - **方法B**: 一時的なファイルは今後不要なため、すべて完全に削除する。

---

## 整理対象の分類案

### 1. ルート（元の場所）に残す本番ファイル
シミュレーションの運用やダッシュボードの表示に**直接必要なファイル**です。

- **シミュレーション・実績データ**:
  - `Kanken23510_2026年版_v6.xlsx`（最新のシミュレーションファイル）
  - `Kanken23510_2026年版_v5.xlsx`（前バージョンのファイル、履歴として保持）
  - `2026実績.csv`（最新の実績データ）
  - `6.24までの実績.xlsx`（元データとしていただいた実績ファイル）
  - `分配店舗向け.xlsx` / `分配店舗向けclaude用.xlsx`（シミュレーションの元になったファイル）
- **ダッシュボード・システム関連**:
  - `dashboard.html`（可視化ダッシュボード）
  - `dashboard_data.js`（ダッシュボード用データ）
  - `dashboard_data.json`（ダッシュボード用バックアップデータ）
  - `backend_calc.py`（データを計算・更新するプログラム）
  - `データ更新を実行する.bat`（ダブルクリックでデータを更新するバッチファイル）
- **システム設定**:
  - `AGENT.md` / `GEMINI.md` / `.gitignore` / `.git`（管理用ファイル）

---

### 2. 整理（移動または削除）する一時ファイル
今回のシミュレーション作成やデータ検証の過程で作成した、**作業用の一時的なファイル**です。

- **検証用スクリプト（Pythonファイル）**:
  - `check_csv.py`, `check_v6.py`, `finalize_v6.py`, `find_paths.py`, `fix_summary.py`, `list_files.py`, `rebuild_v6.py`, `run_v6.py`, `scan_summary.py`, `test_paths.py`, `update_v6.py`
- **調査ログ・テキストファイル**:
  - `claude_all_sheets.txt`, `claude_file_check.txt`, `claude_sheet1.txt`
  - `existing_csv_check.txt`, `jisseki_check.txt`, `jisseki_summary.txt`
  - `run_v6_log.txt`, `update_v6_log.txt`, `store_compare.txt`
  - `v5_all_colors_info.txt`, `v5_check.txt`, `v5_colors.txt`, `v5_fogpink.txt`, `v5_logic.txt`, `v5_sheets.txt`, `v5_sim.txt`, `v5_summary.txt`
- **その他の一時ファイル**:
  - `検証用_統合データ.csv`（検証用に作成した中間データ。backend_calc.py からいつでも再生成可能です）
  - `.~lock.Kanken23510_2026年版_v6.xlsx#`（Excelの一時ロックファイル。Excelを閉じれば自動で消えるか、手動削除可能です）

---

## 実施手順

1. **ユーザー確認**: ユーザーから「方法A（バックアップフォルダへ移動）」か「方法B（完全削除）」の指示をいただきます。
2. **フォルダ作成**: 方法Aの場合、`work_backup` フォルダを新規作成します。
3. **ファイル移動/削除**: 上記の「整理する一時ファイル」を指定に従って移動または削除します。
4. **動作確認**: 整理後、`データ更新を実行する.bat` を実行し、ダッシュボードの表示（`dashboard.html`）が正常に更新・表示されるか確認します。
