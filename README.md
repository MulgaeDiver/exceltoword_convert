# Excel to Word 변환기

Excel 파일을 업로드하여 원하는 양식의 Word 문서로 변환하는 자동화 프로그램입니다.

## 기능

1. **Excel 파일 업로드**: .xlsx, .xls 형식의 Excel 파일을 업로드
2. **시트 분석**: 업로드된 Excel 파일의 시트명과 헤더 구조 자동 분석
3. **사용자 선택 인터페이스**: 
   - 변환할 시트 선택
   - Title 헤더 선택 (번호가 매겨질 메인 헤더)
   - Sub 헤더 선택 (들여쓰기로 표시될 하위 헤더들)
4. **Word 문서 생성**: 선택한 설정에 따라 Word 문서 자동 생성

## 설치 및 실행

### 1. 필요한 라이브러리 설치
```bash
pip install -r requirements.txt
```

### 2. 프로그램 실행
```bash
streamlit run excel_to_word_converter.py
```

### 3. 웹 브라우저에서 사용
- 자동으로 열리는 브라우저에서 프로그램을 사용할 수 있습니다
- 기본 주소: http://localhost:8501

## 사용 방법

1. **Excel 파일 업로드**: 변환하고 싶은 Excel 파일을 선택하여 업로드
2. **시트 선택**: 여러 시트가 있는 경우 변환할 시트를 선택
3. **헤더 설정**: 
   - Title 헤더: 번호가 매겨질 메인 헤더 선택 (예: "모델명")
   - Sub 헤더: 들여쓰기로 표시될 하위 헤더들 선택 (예: "Ticket", "Title", "증상", "해결")
4. **Word 문서 생성**: 설정을 확인한 후 "Word 문서 생성" 버튼 클릭
5. **다운로드**: 생성된 Word 문서를 다운로드

## 출력 형식

생성되는 Word 문서는 다음과 같은 구조를 가집니다:

```
1. 모델명: GC85
   Ticket: T22B045
   Title: S-DAP Cal Error
   증상: S-DAP Cal Error
   해결: DAP 관련 문의

2. 모델명: GM60
   Ticket: T22B040
   Title: error loading OS and charging issue
   증상: error loading OS and charging issue
   해결: OS recovery를 진행하지 못하여 가이드 전달
```

## 필요한 라이브러리

- pandas: Excel 파일 읽기 및 데이터 처리
- openpyxl: Excel 파일 형식 지원
- python-docx: Word 문서 생성
- streamlit: 웹 인터페이스

## 주의사항

- Excel 파일의 첫 번째 행은 헤더로 인식됩니다
- 빈 셀이나 공백만 있는 행은 자동으로 제외됩니다
- Title 헤더로 선택한 열의 값이 같은 행들은 하나의 그룹으로 묶입니다


