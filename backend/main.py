from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import vertexai
from vertexai.generative_models import GenerativeModel, GenerationConfig
import json, re

vertexai.init(project="manual-auto-generator", location="asia-northeast3")

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

class AnalyzeRequest(BaseModel):
    model: str
    prompt: str
    systemInstruction: str
    schema: str

@app.post("/api/analyze")
async def analyze(req: AnalyzeRequest):
    try:
        full_prompt = (
            f"{req.prompt}\n\n"
            f"위 데이터를 분석하여 아래 JSON 스키마 구조로만 정확하게 출력하라. "
            f"마크다운 백틱(```) 없이 순수 JSON 객체만 반환하라.\n{req.schema}"
        )
        model = GenerativeModel(
            req.model,
            system_instruction=req.systemInstruction,
        )
        response = model.generate_content(
            full_prompt,
            generation_config=GenerationConfig(
                response_mime_type="application/json",
                temperature=0.1,
                top_p=0.6,
                max_output_tokens=65536,
            ),
        )
        text = response.text
        match = re.search(r'\{.*\}', text, re.DOTALL)
        if not match:
            raise HTTPException(
                status_code=500,
                detail="AI 응답에서 JSON을 찾지 못했습니다.",
            )
        return {"text": match.group(0)}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
