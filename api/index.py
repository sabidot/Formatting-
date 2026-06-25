"""
Document Formatting Agent
"""

import json
import io
import sys
import traceback
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles

from ppm_format_agent import (
    apply_all_fixes,
    has_h2,
    run_all_checks,
    SEVERITY_ORDER,
    SPEC,
)

# ---------------------------------------------------------------------------
app = FastAPI(title="Document Formatting Agent", version="1.3.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

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


def _load_doc(data: bytes):
    from docx import Document
    return Document(io.BytesIO(data))


def _build_report(filename, flags, doc):
    return {
        "file": filename,
        "total_flags": len(flags),
        "h1_context": (
            f"H2 present → H1 expected at {SPEC['h1_pt_with_h2']}pt"
            if has_h2(doc)
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
    for p in [ROOT / "static" / "index.html", Path("static/index.html")]:
        if p.exists():
            return HTMLResponse(p.read_text(encoding="utf-8"))
    return HTMLResponse("<h1>UI not found</h1>", status_code=404)


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.get("/agents")
async def list_agents():
    return {k: {x: v[x] for x in ("label", "accept", "icon")} for k, v in AGENTS.items()}


@app.post("/check/{agent_key}")
async def check_document(agent_key: str, file: UploadFile = File(...)):
    if agent_key not in AGENTS:
        raise HTTPException(404, f"Unknown agent '{agent_key}'")
    if not file.filename.lower().endswith(AGENTS[agent_key]["accept"]):
        raise HTTPException(400, f"Expected {AGENTS[agent_key]['accept']} file")
    data = await file.read()
    try:
        doc   = _load_doc(data)
        flags = AGENTS[agent_key]["run_checks"](doc)
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
        ext  = agent["accept"]   # ".docx"

        # Load → fix → save to memory buffer
        doc = _load_doc(data)
        agent["apply_fixes"](doc)

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)

        # Return as plain .docx — browser downloads it directly
        return Response(
            content=buf.read(),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f'attachment; filename="{stem}_fixed{ext}"'
            },
        )
    except HTTPException:
        raise
    except Exception:
        raise HTTPException(500, traceback.format_exc())


# Vercel adapter
try:
    from mangum import Mangum
    handler = Mangum(app, lifespan="off")
except ImportError:
    handler = None
