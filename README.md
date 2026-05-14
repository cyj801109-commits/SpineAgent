# SPINE Agent

SI 프로젝트 요구사항정의서 생성 에이전트 - React + FastAPI + Vertex AI

## 실행 방법

### 백엔드
```bash
cd backend
pip install -r requirements.txt
gcloud auth application-default login
uvicorn main:app --reload --port 8000
```

### 프론트엔드
```bash
cd frontend
npm install
npm run dev
```

## GCP 설정
- Project ID: manual-auto-generator
- Region: asia-northeast3 (서울)
- 필요 API: Vertex AI API 활성화
