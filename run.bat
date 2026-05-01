@echo off
chcp 65001 > nul
cd /d %~dp0

echo.
echo  ================================================
echo   WBS - Outlook 同期ツール
echo  ================================================
echo.

python --version > nul 2>&1
if errorlevel 1 (
    echo  [エラー] Python が見つかりません。
    echo  Python 3.x をインストールしてから再実行してください。
    pause
    exit /b 1
)

echo  依存パッケージを確認中...
python -m pip install flask openpyxl pywin32 -q
if errorlevel 1 (
    echo  [エラー] パッケージのインストールに失敗しました。
    pause
    exit /b 1
)
echo  パッケージ: OK
echo.
echo  サーバー起動中...  http://localhost:5000
echo  終了するには Ctrl+C を押してください。
echo.

timeout /t 1 /nobreak > nul
start "" "http://localhost:5000"

python -X utf8 app.py
pause
