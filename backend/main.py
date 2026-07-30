from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import vertexai
from vertexai.generative_models import GenerativeModel, GenerationConfig, Part
import base64, json, os, tempfile, asyncio, threading
from typing import Optional
from pathlib import Path

# Railway 환경에서 GCP 서비스 계정 인증 처리
_creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
if _creds_json and not os.environ.get("GOOGLE_APPLICATION_CREDENTIALS"):
    _tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False)
    _tmp.write(_creds_json)
    _tmp.close()
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = _tmp.name

try:
    vertexai.init(project="manual-auto-generator", location="asia-northeast3")
except Exception as e:
    print(f"[WARN] vertexai.init failed: {e}")

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/health")
async def health():
    return {"status": "ok"}

class PdfFile(BaseModel):
    name: str
    data: str  # base64 encoded

class AnalyzeRequest(BaseModel):
    model: str
    prompt: str
    systemInstruction: str
    schema: str
    pdf_files: Optional[list[PdfFile]] = None
    temperature: Optional[float] = 0.2

@app.post("/api/analyze")
async def analyze(req: AnalyzeRequest):
    full_prompt = (
        f"{req.prompt}\n\n"
        f"위 데이터를 분석하여 아래 JSON 스키마 구조로만 정확하게 출력하라. "
        f"마크다운 백틱(```) 없이 순수 JSON 객체만 반환하라.\n{req.schema}"
    )
    model = GenerativeModel(
        req.model,
        system_instruction=req.systemInstruction,
    )

    content_parts = []
    if req.pdf_files:
        for pdf in req.pdf_files:
            pdf_bytes = base64.b64decode(pdf.data)
            content_parts.append(
                Part.from_data(data=pdf_bytes, mime_type="application/pdf")
            )
    content_parts.append(full_prompt)

    gen_config = GenerationConfig(
        response_mime_type="application/json",
        temperature=req.temperature,
        top_p=0.6 if req.temperature > 0 else 1.0,
        top_k=40 if req.temperature > 0 else 1,
        max_output_tokens=65536,
    )

    async def stream_generator():
        q: asyncio.Queue = asyncio.Queue()
        loop = asyncio.get_event_loop()

        def produce():
            try:
                responses = model.generate_content(
                    content_parts,
                    generation_config=gen_config,
                    stream=True,
                )
                for chunk in responses:
                    if chunk.text:
                        loop.call_soon_threadsafe(q.put_nowait, chunk.text)
            except Exception as exc:
                loop.call_soon_threadsafe(q.put_nowait, exc)
            finally:
                loop.call_soon_threadsafe(q.put_nowait, None)  # sentinel

        thread = threading.Thread(target=produce, daemon=True)
        thread.start()

        try:
            while True:
                item = await q.get()
                if item is None:
                    break
                if isinstance(item, Exception):
                    yield f"data: {json.dumps({'error': str(item)}, ensure_ascii=False)}\n\n"
                    return
                yield f"data: {json.dumps({'token': item}, ensure_ascii=False)}\n\n"
            yield "data: [DONE]\n\n"
        except Exception as e:
            yield f"data: {json.dumps({'error': str(e)}, ensure_ascii=False)}\n\n"
        finally:
            thread.join(timeout=5)

    return StreamingResponse(
        stream_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )

# --- 프론트엔드 정적 파일 서빙 ---
static_dir = Path(__file__).resolve().parent / "static"
if static_dir.is_dir():
    app.mount("/assets", StaticFiles(directory=static_dir / "assets"), name="assets")

    @app.get("/{full_path:path}")
    async def serve_spa(full_path: str):
        file_path = static_dir / full_path
        if file_path.is_file():
            return FileResponse(file_path)
        return FileResponse(static_dir / "index.html")
