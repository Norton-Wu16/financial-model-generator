"""FastAPI 后端入口。

提供:
- POST /api/generate  接收模型类型+参数，返回 Excel 文件
- POST /api/preview   接收模型类型+参数，返回计算结果 JSON（用于网页端预览）
- GET  /api/defaults/{model_type}  返回默认参数
- GET  /api/presets               列出所有行业预设
- GET  /api/presets/{preset_id}/{model_type}  返回某预设某模型的参数
- POST /api/upload                上传 CSV/Excel 财报，返回映射后的参数
- 静态文件托管 (前端 build 产物)
"""
import io
import os
from pathlib import Path

from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, FileResponse, JSONResponse
from openpyxl import Workbook
from pydantic import BaseModel, ValidationError

from backend.schemas import (
    ThreeStatementParams, DCFParams, LBOParams, GenerateRequest,
)
from backend.builders import (
    build_three_statement, build_dcf, build_lbo,
)
from backend import calculations
from backend import presets as preset_lib

app = FastAPI(title="Financial Model Generator", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

BUILDERS = {
    "three_statement": (ThreeStatementParams, build_three_statement),
    "dcf": (DCFParams, build_dcf),
    "lbo": (LBOParams, build_lbo),
}

FILENAMES = {
    "three_statement": "Three_Statement_Model.xlsx",
    "dcf": "DCF_Model.xlsx",
    "lbo": "LBO_Model.xlsx",
}


@app.get("/api/defaults/{model_type}")
def get_defaults(model_type: str):
    if model_type not in BUILDERS:
        raise HTTPException(404, f"Unknown model type: {model_type}")
    schema_cls, _ = BUILDERS[model_type]
    defaults = schema_cls().model_dump()
    return JSONResponse(defaults)


@app.post("/api/generate")
def generate_model(req: GenerateRequest):
    if req.model_type not in BUILDERS:
        raise HTTPException(400, f"Unknown model type: {req.model_type}")

    schema_cls, builder = BUILDERS[req.model_type]
    try:
        params = schema_cls(**req.params)
    except ValidationError as e:
        raise HTTPException(422, str(e))

    wb: Workbook = builder(params)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    filename = FILENAMES[req.model_type]
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/api/preview")
def preview_model(req: GenerateRequest):
    """接收模型类型+参数，返回计算结果 JSON（用于网页端预览）。

    复用 GenerateRequest schema，调用 calculations.calculate()。
    """
    if req.model_type not in BUILDERS:
        raise HTTPException(400, f"Unknown model type: {req.model_type}")

    schema_cls, _ = BUILDERS[req.model_type]
    try:
        params = schema_cls(**req.params)
    except ValidationError as e:
        raise HTTPException(422, str(e))

    try:
        result = calculations.calculate(req.model_type, req.params)
        return JSONResponse(result)
    except Exception as e:
        raise HTTPException(500, f"计算失败: {str(e)}")


@app.get("/api/presets")
def list_presets():
    """列出所有行业预设。"""
    return JSONResponse(preset_lib.list_presets())


@app.get("/api/presets/{preset_id}/{model_type}")
def get_preset(preset_id: str, model_type: str):
    """返回某预设某模型的参数。"""
    if model_type not in BUILDERS:
        raise HTTPException(400, f"Unknown model type: {model_type}")
    try:
        params = preset_lib.get_preset(preset_id, model_type)
        return JSONResponse(params)
    except KeyError:
        raise HTTPException(404, f"Preset not found: {preset_id}")
    except ValueError as e:
        raise HTTPException(400, str(e))


@app.post("/api/upload")
async def upload_report(model_type: str = Form(...), file: UploadFile = File(...)):
    """上传 CSV/Excel 财报，返回映射到 schema 的参数。

    返回: {matched: {...}, unmatched: [...], warnings: [...]}
    """
    if model_type not in BUILDERS:
        raise HTTPException(400, f"Unknown model type: {model_type}")

    # 限制文件大小 10MB
    content = await file.read()
    if len(content) > 10 * 1024 * 1024:
        raise HTTPException(413, "File too large (max 10MB)")

    # 限制文件类型
    filename = file.filename or ""
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if ext not in ("csv", "xlsx", "xls"):
        raise HTTPException(400, "Unsupported file format. Only .csv, .xlsx, .xls allowed")

    try:
        from backend import upload_parser
        result = upload_parser.parse_financial_report(content, filename, model_type)
        return JSONResponse(result)
    except Exception as e:
        raise HTTPException(500, f"解析失败: {str(e)}")


@app.get("/api/health")
def health():
    return {"status": "ok"}


# 静态文件托管 (前端单页应用)
FRONTEND_DIR = Path(__file__).parent.parent / "frontend"
if (FRONTEND_DIR / "index.html").exists():
    from fastapi.staticfiles import StaticFiles
    app.mount("/", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")
