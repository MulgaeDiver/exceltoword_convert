@echo off
echo Excel to Word 변환기 설치를 시작합니다...
echo.

REM Python이 설치되어 있는지 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo 오류: Python이 설치되어 있지 않습니다.
    echo Python 3.8 이상을 설치해주세요: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Python이 설치되어 있습니다.
echo.

REM 필요한 라이브러리 설치
echo 필요한 라이브러리를 설치합니다...
pip install -r requirements.txt

if errorlevel 1 (
    echo 오류: 라이브러리 설치에 실패했습니다.
    pause
    exit /b 1
)

echo.
echo 설치가 완료되었습니다!
echo.
echo 사용 방법:
echo 1. run_app.py 파일을 더블클릭하거나
echo 2. 명령 프롬프트에서 "python run_app.py" 실행
echo.
pause


