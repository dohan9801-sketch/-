@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo [설치] 필요한 프로그램을 설치합니다...
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo 설치에 실패했습니다. 파이썬이 설치되어 있는지 확인하세요.
  echo https://www.python.org/downloads/ 에서 설치할 때 "Add python.exe to PATH"를 꼭 체크하세요.
  pause
  exit /b 1
)
echo.
echo 설치가 끝났습니다.
pause
