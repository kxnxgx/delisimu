@echo off
chcp 65001 > nul
set PYTHONIOENCODING=utf-8

echo --------------------------------------------------
echo Kanken 在庫分配シミュレーション - データ更新処理
echo --------------------------------------------------
echo.
echo 最新のエクセルとCSVファイルを読み込んで、
echo ダッシュボードのデータを最新版に更新します...
echo.

python backend_calc.py

echo.
echo 処理が完了しました！
echo dashboard.html を開いて（またはリロードして）最新のデータを確認してください。
echo --------------------------------------------------
pause
