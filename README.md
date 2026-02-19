# GPT_Realestate_controller
엄마 부동산 장부용 프로그램 (데스크탑 + API 서비스)

## 핵심 방향
- **Google API 없이도 100% 운영 가능** (SQLite + 파일 내보내기)
- 필요 시 웹훅(Apps Script) 연동으로 즉시 Google Sheets 반영
- 데스크탑 UI(`tkinter`)와 경량 HTTP API 서비스를 함께 제공

## 빠른 시작
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 1) 데스크탑 실행
```bash
python run.py
```

### 2) API 서비스 실행
```bash
python service_run.py
```

- 기본 주소: `http://127.0.0.1:8000`

## 주요 API
- `GET /health`
- `GET /properties`, `POST /properties`, `PATCH /properties/{id}`
- `GET /customers`, `POST /customers`
- `GET /matching/{customer_id}`
- `POST /viewings`
- `POST /tasks/reconcile`, `GET /tasks`
- `POST /sync/export`
- `POST /proposal/message/{customer_id}`

## 로컬 파일 내보내기
기본 동기화 폴더(또는 API 입력한 `sync_dir`)에 다음 파일 생성:
- `exports/visible_properties.csv`
- `exports/visible_customers.csv`
- `exports/open_tasks.csv`
- `exports/snapshot.json`
- `exports/viewings.ics`

## 프로젝트 구조
- `app/storage.py`: SQLite 저장소/마이그레이션
- `app/desktop_app.py`: 데스크탑 UI
- `app/api_service.py`: 경량 HTTP API 서비스
- `app/tasks_engine.py`: 자동 할 일(Next Action) 엔진
- `app/matching.py`: 고객-매물 매칭 규칙
- `app/sheet_sync.py`: CSV/JSON/ICS 내보내기 + 웹훅 업로드
- `app/proposal.py`: 카카오 문구/제안서 생성
