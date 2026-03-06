# 비교견적 자동 생성기

본견적 엑셀 파일 1개를 업로드하면 첫 번째 시트를 본견적으로 읽고, `거성` / `해광` 비교견적 시트를 포함한 새 엑셀 파일을 자동 생성하는 Flask 프로젝트입니다.

## 기능
- 본견적 엑셀 업로드
- 첫 번째 시트를 `자재피아기본견적` 시트로 복사
- 품목 수 자동 인식 (빈 행 전까지)
- 업체별 가산율/할인율 입력
- 비교견적 시트 자동 생성
- 결과 파일 즉시 다운로드
- 기존 비교견적 양식 최대한 유지

## 폴더 구조
```text
quote_compare_project/
├─ app.py
├─ requirements.txt
├─ resources/
│  └─ comparison_template.xlsx
├─ src/
│  ├─ excel_utils.py
│  └─ quote_generator.py
├─ templates/
│  └─ index.html
├─ static/
│  └─ style.css
└─ tests/
   └─ smoke_test.py
```

## 실행 방법
```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
python app.py
```

브라우저에서 아래 주소를 열면 됩니다.
```text
http://127.0.0.1:5000
```

## 입력 파일 조건
- 첫 번째 시트가 본견적이어야 합니다.
- 형식은 동일하고, 품목 수만 달라져야 합니다.
- 기본적으로 다음 열을 사용합니다.
  - A열: 품목명
  - B열: 수량
  - C열: 가격(단가)
  - D열: 금액

## 사용 방법
1. 본견적 엑셀을 업로드합니다.
2. 업체명과 가산율/할인율을 입력합니다.
   - 예: `15` 입력 = 15% 가산
   - 예: `-5` 입력 = 5% 할인
3. `비교견적 생성`을 누릅니다.
4. 결과 엑셀을 다운로드합니다.

## 주의
- 현재 버전은 엑셀 생성 기능에 집중한 1차 버전입니다.
- PDF 직접 출력은 포함하지 않았습니다. 생성된 엑셀을 열어 PDF로 저장하면 됩니다.
- 템플릿 양식이 바뀌면 `src/quote_generator.py`의 행/열 기준도 같이 수정해야 합니다.

## 개선 예정
- 업체 수 동적 추가
- 웹 화면에서 비교견적 미리보기
- PDF 직접 출력
- 품목별 예외 단가 수정
