@echo off
cd /d "%~dp0"
chcp 65001 > nul
set PYTHONIOENCODING=utf-8

echo ==================================================
echo  Kanken 出荷計画 - データ一括更新
echo ==================================================
echo.
echo [Step 1/3] 全体サマリー再構築中 (build_summary.py)...
echo --------------------------------------------------
python build_summary.py
if %errorlevel% neq 0 (
    echo.
    echo [エラー] build_summary.py が失敗しました。処理を中断します。
    pause
    exit /b 1
)

echo.
echo [Step 2/3] ダッシュボード計算データを生成中...
echo --------------------------------------------------
python backend_calc.py
if %errorlevel% neq 0 (
    echo.
    echo [エラー] backend_calc.py が失敗しました。処理を中断します。
    pause
    exit /b 1
)

echo.
echo [Step 3/3] ダッシュボードHTMLを生成中...
echo --------------------------------------------------
python update_dashboard.py
if %errorlevel% neq 0 (
    echo.
    echo [エラー] update_dashboard.py が失敗しました。
    pause
    exit /b 1
)

echo.
echo ==================================================
echo  すべての処理が完了しました！
echo ==================================================
echo.
echo ダッシュボードを自動で開きます...
start "" "Kanken_Dashboard_最新版.html"
echo.
pause
