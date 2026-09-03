@echo off
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
cd /d "%~dp0"
call config.cmd
call :main > setup_log.txt 2>&1
type setup_log.txt
echo.
echo (log saved to setup_log.txt)
pause
exit /b

:main
echo ===== SETUP %date% %time% =====
where git >nul 2>nul
if errorlevel 1 (
    echo ERROR: git is not installed. Install from https://git-scm.com then run again.
    exit /b 1
)
if "%GH_USER%"=="YOUR-GITHUB-ID" (
    echo ERROR: open config.cmd and set GH_USER / GH_REPO first.
    exit /b 1
)

python build_data.py
if errorlevel 1 (
    echo ERROR: build_data.py failed
    exit /b 1
)

if not exist ".git" git init
git config user.name "%GH_USER%"
git config user.email "%GH_MAIL%"
git branch -M main
git remote remove origin >nul 2>nul
git remote add origin https://github.com/%GH_USER%/%GH_REPO%.git

git rm -r --cached __pycache__ >nul 2>nul
git add -A
git commit -m "first upload"
git push -u origin main
if errorlevel 1 (
    echo ERROR: push failed. If a GitHub login window appeared, sign in and run this again.
    exit /b 1
)
echo.
echo DONE - uploaded to https://github.com/%GH_USER%/%GH_REPO%
echo NEXT - Settings ^> Pages ^> Branch: main / (root) ^> Save
echo SITE  - https://%GH_USER%.github.io/%GH_REPO%/
exit /b 0
