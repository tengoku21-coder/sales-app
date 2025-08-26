
# 영업 보고서 관리 (Streamlit 웹버전)

## 로컬 실행
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud 배포
1. 이 저장소를 GitHub에 업로드합니다. (파일: `app.py`, `requirements.txt`, `README.md`)
2. https://share.streamlit.io → **New app** → Repository/Branch 선택 → App file에 `app.py` 지정 → Deploy
3. 클라우드 인스턴스는 파일시스템이 수시로 초기화될 수 있으므로, 우측의 **엑셀 내려받기** 버튼으로 주기적 백업을 권장합니다.
   - 저장 경로를 바꾸려면 환경변수 **`SR_EXCEL_PATH`** 를 설정하세요.

## 기능
- 기본 영업자: **김범준** (변경 가능), 날짜 기본값: **오늘**
- 연락처 자동 하이픈(02/010 규칙)
- 충전기 모델·부대공사 수량 입력 + 기타모델 추가/삭제
- 미리보기 자동 생성
- 엑셀 저장/수정/삭제, 목록 표시는 ID 숨김
