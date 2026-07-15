@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo ====================================================
echo  SCOSCHE Packing App
echo ====================================================
echo.

:: 安裝相依套件（第一次執行）
pip install -r requirements.txt -q

:: 初次使用：先掃描歷史資料
if not exist "data\items.json" (
    echo [首次啟動] 掃描歷史 Packing List...
    python scanner.py --full
)

:: 啟動伺服器
echo.
echo 啟動中... 請在瀏覽器開啟 http://localhost:5001
echo 按 Ctrl+C 可關閉伺服器
echo.
python app.py
pause
