@echo off
chcp 65001 >nul
REM ========================================
REM  GitHub自動同期バッチ
REM  ダブルクリックまたはタスクスケジューラで実行
REM ========================================

cd /d "C:\Users\hitos\OneDrive\AI関連\DeepResearchをつかった情報調査"

echo [%date% %time%] Git同期を開始...

REM 変更があるかチェック
git status --porcelain > nul 2>&1
if errorlevel 1 (
    echo Gitリポジトリではありません。
    pause
    exit /b 1
)

REM 変更をステージング
git add .

REM 変更があるか確認
git diff --cached --quiet
if errorlevel 1 (
    REM 変更あり → コミット＆プッシュ
    for /f "tokens=1-3 delims=/ " %%a in ('date /t') do set DATESTR=%%a%%b%%c
    for /f "tokens=1-2 delims=: " %%a in ('time /t') do set TIMESTR=%%a%%b
    git commit -m "Auto sync %DATESTR% %TIMESTR%"
    git push origin main
    echo [%date% %time%] プッシュ完了！
) else (
    echo [%date% %time%] 変更なし。スキップ。
)

echo.
echo 完了しました。
timeout /t 5
