<<<<<<< HEAD
# 출장보고서 자동화 시스템

## 📋 개요
Streamlit을 사용한 웹 기반 출장보고서 자동 생성 시스템입니다.

## 🚀 실행 방법

### 1. 필수 패키지 설치
```bash
pip install -r requirements.txt
```

### 2. 프로그램 실행
```bash
streamlit run main.py
```

### 3. 웹 브라우저에서 접속
- 자동으로 브라우저가 열리며 `http://localhost:8501`에서 접속 가능합니다.

## 📁 파일 구조
```
보고서 자동화/
├── main.py                 # 메인 Streamlit 앱
├── excel_generator.py      # 엑셀 파일 생성 모듈
├── data_manager.py         # 데이터 관리 모듈
├── requirements.txt        # 필수 패키지 목록
├── business_trip_data.json # 데이터 저장 파일 (자동 생성)
└── README.md              # 이 파일
```

## ✨ 주요 기능

### 📝 출장보고서 입력
- **기본 정보**: 성명, 소속, 직책, 담당자, 과제책임자, 대표이사
- **출장 정보**: 연구과제명, 출장지, 출장 기간, 출장 목적/결과
- **비용 정보**: 일비, 숙박료, 식비 (자동 합계 계산)
- **추가 수당**: 휴일출장, 특수출장, 위험출장 수당

### 🎯 드롭다운 선택 방식
- 미리 설정된 옵션에서 선택하여 입력 실수 방지
- 새로운 데이터 추가 및 관리 기능

### 📄 자동 엑셀 생성
- A4 용지 크기에 최적화된 레이아웃
- 25행씩 2페이지로 자동 분할
- 완전한 테두리 및 서식 적용
- 즉시 다운로드 가능

### 💾 데이터 관리
- JSON 파일로 데이터 영구 저장
- 웹 인터페이스에서 직접 데이터 추가/삭제
- 기본 데이터 자동 생성

## 🔧 사용법

1. **웹 페이지 접속**: `streamlit run main.py` 실행 후 브라우저에서 접속
2. **정보 입력**: 드롭다운과 입력 필드를 통해 출장 정보 입력
3. **미리보기 확인**: 우측 패널에서 입력 내용 확인
4. **보고서 생성**: "출장보고서 생성" 버튼 클릭
5. **파일 다운로드**: 생성된 엑셀 파일 다운로드

## 📊 데이터 커스터마이징
- 사이드바에서 새로운 이름, 부서, 출장지 등 추가 가능
- `business_trip_data.json` 파일을 직접 편집하여 대량 데이터 추가 가능

## 🛠️ 기술 스택
- **Frontend**: Streamlit
- **Backend**: Python
- **Excel 생성**: openpyxl
- **데이터 저장**: JSON

## 📝 생성되는 출장보고서 포맷
- 표준 엔이비 출장복명서 양식
- A4 용지 2페이지 구성
- 모든 필수 항목 자동 배치
- 인쇄 최적화된 레이아웃

## 🚨 주의사항
- 필수 정보(성명, 소속, 출장지, 연구과제명)는 반드시 입력
- 파일은 현재 폴더에 저장됨
- 동일한 파일명 존재시 덮어쓰기됨

## 📞 문의
- 개발자: NEB AI 데이터팀
- 버전: v1.0
- 최종 수정일: 2025.06.17
=======
# 🎈 Blank app template

A simple Streamlit app template for you to modify!

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://blank-app-template.streamlit.app/)

### How to run it on your own machine

1. Install the requirements

   ```
   $ pip install -r requirements.txt
   ```

2. Run the app

   ```
   $ streamlit run streamlit_app.py
   ```
>>>>>>> a50529b87c11e0b405ba292716adaa830994b8d8
