"""
Document Formatting Agent — works on Render, Railway, and locally.
Vercel adapter (Mangum) included but file size limits apply there.
"""

import json
import io
import zipfile
import os
import sys
import shutil
import tempfile
import traceback
from pathlib import Path

# Project root on path so ppm_format_agent imports correctly
ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles

from ppm_format_agent import (
    apply_all_fixes,
    has_h2,
    run_all_checks,
    SEVERITY_ORDER,
    SPEC,
)

# ---------------------------------------------------------------------------
app = FastAPI(title="Document Formatting Agent", version="1.2.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Mount static folder — resolve relative to this file so it works on any host
_static = ROOT / "static"
if _static.exists():
    app.mount("/static", StaticFiles(directory=str(_static)), name="static")

AGENTS = {
    "docx": {
        "label":       "Word Document",
        "accept":      ".docx",
        "icon":        "📄",
        "run_checks":  run_all_checks,
        "apply_fixes": apply_all_fixes,
    },
}


def _load_doc(agent_key: str, data: bytes):
    if agent_key == "docx":
        from docx import Document
        return Document(io.BytesIO(data))
    raise HTTPException(400, f"Unknown agent: {agent_key}")


def _build_report(filename, flags, doc):
    doc_has_h2 = has_h2(doc)
    return {
        "file": filename,
        "total_flags": len(flags),
        "h1_context": (
            f"H2 present → H1 expected at {SPEC['h1_pt_with_h2']}pt"
            if doc_has_h2
            else f"No H2 → H1 expected at {SPEC['h1_pt_no_h2']}pt"
        ),
        "severity_summary": {
            sev: sum(1 for f in flags if f.get("severity") == sev)
            for sev in ("critical", "high", "medium", "low")
        },
        "flags": sorted(
            flags,
            key=lambda f: SEVERITY_ORDER.get(f.get("severity", "low"), 9),
        ),
    }


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def root():
    candidates = [
        ROOT / "static" / "index.html",
        Path("static/index.html"),
    ]
    for p in candidates:
        if p.exists():
            return HTMLResponse(p.read_text(encoding="utf-8"))
    return HTMLResponse("<h1>UI not found</h1>", status_code=404)


@app.get("/health")
async def health():
    return {"status": "ok", "agents": list(AGENTS)}


@app.get("/agents")
async def list_agents():
    return {k: {x: v[x] for x in ("label", "accept", "icon")} for k, v in AGENTS.items()}


@app.post("/check/{agent_key}")
async def check_document(agent_key: str, file: UploadFile = File(...)):
    if agent_key not in AGENTS:
        raise HTTPException(404, f"Unknown agent '{agent_key}'")
    agent = AGENTS[agent_key]
    if not file.filename.lower().endswith(agent["accept"]):
        raise HTTPException(400, f"Expected {agent['accept']} file")
    data = await file.read()
    try:
        doc   = _load_doc(agent_key, data)
        flags = agent["run_checks"](doc)
        return JSONResponse(_build_report(file.filename, flags, doc))
    except HTTPException:
        raise
    except Exception:
        raise HTTPException(500, traceback.format_exc())


@app.post("/fix/{agent_key}")
async def fix_document(agent_key: str, file: UploadFile = File(...)):
    if agent_key not in AGENTS:
        raise HTTPException(404, f"Unknown agent '{agent_key}'")
    agent = AGENTS[agent_key]
    if not file.filename.lower().endswith(agent["accept"]):
        raise HTTPException(400, f"Expected {agent['accept']} file")

    data = await file.read()
    try:
        stem = Path(file.filename).stem
        ext  = agent["accept"]

        doc   = _load_doc(agent_key, data)
        flags = agent["run_checks"](doc)
        report_json = json.dumps(
            _build_report(file.filename, flags, doc), indent=2, ensure_ascii=False
        ).encode("utf-8")

        agent["apply_fixes"](doc)

        # Save fixed docx to memory
        fixed_buf = io.BytesIO()
        doc.save(fixed_buf)
        fixed_buf.seek(0)

        # Build zip entirely in memory — no temp files, works on any host
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(f"{stem}_fixed{ext}", fixed_buf.read())
            zf.writestr(f"{stem}_flags.json", report_json)
        zip_buf.seek(0)

        return Response(
            content=zip_buf.read(),
            media_type="application/zip",
            headers={"Content-Disposition": f"attachment; filename={stem}_result.zip"},
        )
    except HTTPException:
        raise
    except Exception:
        raise HTTPException(500, traceback.format_exc())

# ---------------------------------------------------------------------------
# Vercel adapter (only used when deployed to Vercel)
# ---------------------------------------------------------------------------
try:
    from mangum import Mangum
    handler = Mangum(app, lifespan="off")
except ImportError:
    handler = None
