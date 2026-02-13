# GPT_Realestate_controller
엄마부동산장부용 프로그램 (데스크탑 버전)

## 핵심 방향
- **웹서버 없이** 데스크탑에서 바로 실행/입력
- 물건/고객 장부는 로컬 SQLite에 저장
- **숨김=업로드 제외, 보임=업로드 대상** 기준으로 Google Sheets 연동

## 1) 아키텍처
- UI: Python `tkinter` 데스크탑 앱
- DB: SQLite (`app/ledger.db`)
- 업로드:
  - 1순위: `GOOGLE_SHEETS_WEBHOOK_URL`(Apps Script Web App URL)로 JSON 전송
  - 대체: `GOOGLE_DRIVE_SYNC_DIR` 폴더에 CSV 생성(Drive 데스크탑 동기화)

## 2) 제공 기능
- 물건 탭 관리: 아파트단지1/아파트단지2/상가/단독주택
- 물건 등록, 숨김/보임 전환, 삭제
- 고객요구 등록, 숨김/보임 전환, 삭제
- **숨김 제외 데이터 업로드 버튼**

## 3) 실행 방법
```bash
python run.py
```

## 4) Google Sheets 업로드 설정
### A. 권장(버튼 즉시 시트 반영)
1. Google Apps Script 웹앱을 배포해 POST JSON을 시트에 기록하도록 구성
2. 환경변수 설정:
```bash
export GOOGLE_SHEETS_WEBHOOK_URL="https://script.google.com/macros/s/.../exec"
```
3. 앱에서 `숨김 제외 데이터 구글시트 업로드` 클릭

### B. API 없이 운영(드라이브 동기화)
1. 환경변수 설정(선택):
```bash
export GOOGLE_DRIVE_SYNC_DIR="$HOME/Google Drive/My Drive/부동산장부"
```
2. 업로드 버튼 클릭 시 CSV 생성
3. Drive 데스크탑이 자동 동기화 → Gemini/Sheets에서 열람

## 5) 왜 서버가 없어도 되나?
- 입력/조회/숨김관리 모두 로컬 앱에서 처리
- 외부 조회/추천은 Gemini의 Drive 연동으로 별도 처리 가능
- 이 앱은 "현장 입력 + 업로드 버튼" 역할에 집중
