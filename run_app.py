#!/usr/bin/env python3
"""
Excel to Word 변환기 실행 스크립트
이 스크립트를 실행하면 웹 브라우저에서 애플리케이션이 열립니다.
"""

import subprocess
import sys
import webbrowser
import time
import os

def main():
    print("📄 Excel to Word 변환기를 시작합니다...")
    print("=" * 50)
    
    # 현재 디렉토리 확인
    current_dir = os.path.dirname(os.path.abspath(__file__))
    app_file = os.path.join(current_dir, "excel_to_word_converter.py")
    
    if not os.path.exists(app_file):
        print("❌ 오류: excel_to_word_converter.py 파일을 찾을 수 없습니다.")
        print(f"현재 디렉토리: {current_dir}")
        input("엔터를 눌러 종료하세요...")
        return
    
    try:
        print("🚀 Streamlit 애플리케이션을 시작합니다...")
        print("📱 웹 브라우저가 자동으로 열립니다.")
        print("🔗 수동으로 접속하려면: http://localhost:8501")
        print("=" * 50)
        print("⚠️  애플리케이션을 종료하려면 Ctrl+C를 누르세요.")
        print("=" * 50)
        
        # Streamlit 실행
        subprocess.run([sys.executable, "-m", "streamlit", "run", app_file, "--server.port=8501", "--server.headless=true"])
        
    except KeyboardInterrupt:
        print("\n\n👋 애플리케이션이 종료되었습니다.")
    except Exception as e:
        print(f"\n❌ 오류가 발생했습니다: {e}")
        print("\n해결 방법:")
        print("1. 필요한 라이브러리가 설치되어 있는지 확인하세요:")
        print("   pip install -r requirements.txt")
        print("2. Python 버전이 3.8 이상인지 확인하세요.")
        input("\n엔터를 눌러 종료하세요...")

if __name__ == "__main__":
    main()


