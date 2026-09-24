@echo off
REM ============================================================
REM  push.bat — commit & push the mis-portal repo to GitHub.
REM  Vercel auto-deploys on push. Run from anywhere:
REM     C:\github\mis-portal\push.bat
REM     C:\github\mis-portal\push.bat "your commit message"
REM  With no message it stamps the date/time.
REM ============================================================
setlocal
cd /d "%~dp0"

REM optional commit message as first arg (quoted); default = timestamp
set "MSG=%~1"
if "%MSG%"=="" set "MSG=manual deploy %date% %time%"

echo Staging changes in %cd% ...
git add -A

REM nothing to commit? say so and still try to push in case of unpushed commits
git diff --cached --quiet
if %errorlevel%==0 (
  echo No file changes to commit.
) else (
  git commit -m "%MSG%"
)

echo Syncing with remote ...
git pull --rebase --autostash
if errorlevel 1 (
  echo.
  echo ERROR: git pull --rebase failed. Resolve the conflict, then run push.bat again.
  git rebase --abort >nul 2>&1
  pause
  exit /b 1
)

echo Pushing ...
git push
if errorlevel 1 (
  echo.
  echo ERROR: git push failed. Check your GitHub credentials / network.
  pause
  exit /b 1
)

echo.
echo Done. Vercel will redeploy in ~1 minute.
endlocal
