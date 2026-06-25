"""
Vercel serverless entry point.
Vercel looks for handlers in the /api/ folder.
This file wraps the FastAPI app with Mangum so Vercel can invoke it.
"""

import json
import os
import sys
import shutil
import tempfile
import traceback
from pathlib import Path

# Make sure the project root is on the path so ppm_format_agent imports correctly
sys.path.insert(0, str(Path(__file__).parent.parent))

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from mangum import Mangum

from ppm_format_agent import (
    apply_all_fixes,
    has_h2,
    run_all_checks,
    SEVERITY_ORDER,
    SPEC,
)

# ---------------------------------------------------------------------------
app = FastAPI(title="PPM Formatting Agent", version="1.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

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
        import io
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

@app.get("/")
async def root():
    """Serve the frontend."""
    # Works on both Render and Vercel regardless of working directory
    base = Path(__file__).parent.parent
    candidates = [
        base / "static" / "index.html",
        Path("static/index.html"),
        Path("api/../static/index.html"),
    ]
    for p in candidates:
        if p.exists():
            return HTMLResponse(p.read_text(encoding="utf-8"))
    return HTMLResponse("<h1>UI not found</h1><p>static/index.html missing</p>", status_code=404)


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
    if not file.filename.endswith(agent["accept"].lstrip(".")):
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
    if not file.filename.endswith(agent["accept"].lstrip(".")):
        raise HTTPException(400, f"Expected {agent['accept']} file")

    data = await file.read()
    tmp  = tempfile.mkdtemp()
    try:
        stem        = Path(file.filename).stem
        ext         = agent["accept"]
        fixed_path  = Path(tmp) / f"{stem}_fixed{ext}"
        report_path = Path(tmp) / f"{stem}_flags.json"
        zip_base    = Path(tmp) / f"{stem}_result"

        doc   = _load_doc(agent_key, data)
        flags = agent["run_checks"](doc)
        report_path.write_text(
            json.dumps(_build_report(file.filename, flags, doc), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        agent["apply_fixes"](doc)
        doc.save(str(fixed_path))

        zip_path = shutil.make_archive(str(zip_base), "zip", tmp)
        return FileResponse(
            zip_path,
            media_type="application/zip",
            filename=f"{stem}_result.zip",
        )
    except HTTPException:
        raise
    except Exception:
        raise HTTPException(500, traceback.format_exc())
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ---------------------------------------------------------------------------
# Vercel handler — THIS is what Vercel calls
# ---------------------------------------------------------------------------
handler = Mangum(app, lifespan="off")
