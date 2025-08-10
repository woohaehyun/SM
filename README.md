
# 신명약품 발주서 생성 시스템

Streamlit 기반 웹 앱으로 매출, 매입, 현재고 데이터를 기반으로 발주서를 자동 생성합니다.

## 기능
- 매입처 또는 제조사 기준 발주서 생성
- 자동/수동 기간 필터링
- 중복 제품 제거 및 과재고, 부족수량 계산
- 엑셀로 ZIP 파일 다운로드

## 실행 방법
```bash
pip install -r requirements.txt
streamlit run app.py
```

## 필요 파일
- 매출자료.xlsx
- 매입자료.xlsx
- 현재고.xlsx
