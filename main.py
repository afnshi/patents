from __future__ import annotations

import os
import re
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from docxtpl import DocxTemplate

# ========= 路径配置 =========
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# 模板文件名（必须存在）
TPL_CLAIMS = TEMPLATE_DIR / "100001权利要求书.docx"
TPL_SPEC = TEMPLATE_DIR / "100002说明书.docx"
TPL_DRAWINGS = TEMPLATE_DIR / "100003说明书附图.docx"
TPL_ABSTRACT = TEMPLATE_DIR / "100004说明书摘要.docx"

# 如果你用 cloudflared/ngrok 暴露到公网，建议设置这个环境变量为公网前缀
# 例如：export PUBLIC_BASE_URL="https://xxxxx.trycloudflare.com"
PUBLIC_BASE_URL = os.environ.get("PUBLIC_BASE_URL", "").strip().rstrip("/")

app = FastAPI(title="Patent Local Generator (docxtpl)", version="1.0.0")


# ========= 输入结构（Coze -> 后端）=========
class PatentPayload(BaseModel):
    title: str = ""
    tech_field: str = ""
    background: str = ""
    invention_content: str = ""
    drawings_desc: str = ""
    embodiment: str = ""
    claims: List[str] = Field(default_factory=list)
    abstract: str = ""


# ========= 工具函数（轻量处理，不严格）=========
def safe_text(v: Any) -> str:
    """None/非字符串安全转字符串；并做简单 strip"""
    if v is None:
        return ""
    if isinstance(v, str):
        return v.strip()
    return str(v).strip()


def normalize_claims(claims: Any) -> List[str]:
    """
    本地自用：不做严格校验，只做“能用”的整理。
    - 支持 list[str] 或 单个字符串（按行拆）
    - 自动补编号
    - 自动补中文句号
    - 轻量去掉“权利要求1：”前缀
    """
    if claims is None:
        items: List[str] = []
    elif isinstance(claims, list):
        items = [safe_text(x) for x in claims if safe_text(x)]
    else:
        raw = safe_text(claims)
        items = [line.strip() for line in raw.splitlines() if line.strip()]

    out: List[str] = []
    for i, c in enumerate(items, start=1):
        # 去掉“权利要求1：”这种前缀（轻量）
        c = re.sub(r"^\s*权利要求\s*\d+\s*[:：]\s*", "", c).strip()

        # 统一编号为 “i. ”
        if not re.match(rf"^\s*{i}\.\s+", c):
            # 也可能是 1)、1、1. 这类，先去掉前缀再补
            c = re.sub(r"^\s*\d+\s*[\)\、\.]\s*", "", c).strip()
            c = f"{i}. {c}"

        # 补中文句号（仅末尾）
        c = c.rstrip()
        if not c.endswith("。"):
            c = c.rstrip(" .;；:：") + "。"

        out.append(c)
    return out


def ensure_templates_exist() -> None:
    missing = [p for p in (TPL_CLAIMS, TPL_SPEC, TPL_DRAWINGS, TPL_ABSTRACT) if not p.exists()]
    if missing:
        raise HTTPException(
            status_code=500,
            detail="模板不存在或文件名不匹配：\n" + "\n".join(str(p) for p in missing),
        )


def render_docx(template_path: Path, context: Dict[str, Any], out_path: Path) -> None:
    tpl = DocxTemplate(str(template_path))
    tpl.render(context)
    tpl.save(str(out_path))


def build_download_url(request: Request, rel_path: str) -> str:
    """
    返回可点击下载链接：
    - 如果设置了 PUBLIC_BASE_URL（推荐线上 Coze），返回 PUBLIC_BASE_URL + rel_path
    - 否则用 request.base_url（本地/反向代理）拼起来
    """
    if PUBLIC_BASE_URL:
        return f"{PUBLIC_BASE_URL}{rel_path}"
    # request.base_url 形如 "http://127.0.0.1:8000/"
    return str(request.base_url).rstrip("/") + rel_path


# ========= API =========
@app.get("/health")
def health():
    return {"ok": True}


@app.post("/generate-patent")
def generate_patent(payload: PatentPayload, request: Request):
    ensure_templates_exist()

    uid = uuid.uuid4().hex[:10]

    # 轻量清洗
    title = safe_text(payload.title)
    tech_field = safe_text(payload.tech_field)
    background = safe_text(payload.background)
    invention_content = safe_text(payload.invention_content)
    drawings_desc = safe_text(payload.drawings_desc)
    embodiment = safe_text(payload.embodiment)
    abstract = safe_text(payload.abstract)

    claims_list = normalize_claims(payload.claims)

    # docxtpl 变量映射（与你模板中的 {{变量}} 对应）
    context = {
        "TITLE": title,
        "TECH_FIELD": tech_field,
        "BACKGROUND": background,
        "INVENTION_CONTENT": invention_content,
        "DRAWINGS_DESC": drawings_desc,
        "EMBODIMENT": embodiment,
        "CLAIMS": "\n".join(claims_list),
        "ABSTRACT": abstract,
    }

    # 输出文件名
    spec_name = f"{uid}_说明书.docx"
    claims_name = f"{uid}_权利要求书.docx"
    drawings_name = f"{uid}_说明书附图.docx"
    abstract_name = f"{uid}_说明书摘要.docx"

    spec_path = OUTPUT_DIR / spec_name
    claims_path = OUTPUT_DIR / claims_name
    drawings_path = OUTPUT_DIR / drawings_name
    abstract_path = OUTPUT_DIR / abstract_name

    try:
        render_docx(TPL_SPEC, context, spec_path)
        render_docx(TPL_CLAIMS, context, claims_path)
        render_docx(TPL_DRAWINGS, context, drawings_path)
        render_docx(TPL_ABSTRACT, context, abstract_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"生成失败：{e}")

    # 返回完整 URL（更适合 Coze 直接展示可点击链接）
    rel = {
        "spec": f"/download/{spec_name}",
        "claims": f"/download/{claims_name}",
        "abstract": f"/download/{abstract_name}",
        "drawings": f"/download/{drawings_name}",
    }

    return {
        "status": "success",
        "id": uid,
        "files": {k: build_download_url(request, v) for k, v in rel.items()},
    }


@app.get("/download/{filename}")
def download(filename: str):
    # 简单防路径穿越
    if "/" in filename or "\\" in filename:
        raise HTTPException(status_code=400, detail="invalid filename")

    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="file not found")

    return FileResponse(
        path=str(file_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
