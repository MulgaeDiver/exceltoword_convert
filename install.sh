#!/bin/bash

echo "Excel to Word 변환기 설치를 시작합니다..."
echo

# Python이 설치되어 있는지 확인
if ! command -v python3 &> /dev/null; then
    echo "오류: Python3이 설치되어 있지 않습니다."
    echo "Python 3.8 이상을 설치해주세요."
    exit 1
fi

echo "Python이 설치되어 있습니다."
echo

# 필요한 라이브러리 설치
echo "필요한 라이브러리를 설치합니다..."
pip3 install -r requirements.txt

if [ $? -ne 0 ]; then
    echo "오류: 라이브러리 설치에 실패했습니다."
    exit 1
fi

echo
echo "설치가 완료되었습니다!"
echo
echo "사용 방법:"
echo "1. 터미널에서 './run_app.py' 실행 또는"
echo "2. 'python3 run_app.py' 실행"
echo


