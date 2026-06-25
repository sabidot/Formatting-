"""
PPM Formatting Agent — Cloud-ready server
==========================================
Designed to run on Vercel (via Mangum), Railway, Render, or locally.

Plugin architecture: drop a new agent module in /agents/ and register
it in AGENTS below — the UI picks it up automatically.

Current agents:  docx (PPM Impact Fund Word docs)
Planned slots:   xlsx, pptx
"""

import json
import os
import shutil
import tempfile
import traceback
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

# ── Agent registry ──────────────────────────────────────────────────────────
# To add Excel or PowerPoint support later:
#   1. Create agents/ppm_format_agent_xlsx.py  (same interface)
#   2. Add an entry here
#   3. Add a tile in static/index.html (the JS reads /agents endpoint)

from ppm_format_agent import (
    apply_all_fixes,
    has_h2,
    run_all_checks,
    SEVERITY_ORDER,
    SPEC,
)

AGENTS = {
    "docx": {
        "label":       "Word Document",
        "description": "PPM Impact Fund sector .docx files",
        "icon":        "📄",
        "accept":      ".docx",
        "run_checks":  run_all_checks,
        "apply_fixes": apply_all_fixes,
        "extra_meta":  lambda doc: {
            "h1_context": (
                f"H2 present → H1 expected at {SPEC['h1_pt_with_h2']}pt"
                if has_h2(doc)
                else f"No H2 → H1 expected at {SPEC['h1_pt_no_h2']}pt"
            )
        },
    },
    # "xlsx": { ... },   ← plug in later
    # "pptx": { ... },   ← plug in later
}

# ── App ─────────────────────────────────────────────────────────────────────

app = FastAPI(title="PPM Formatting Agent", version="1.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/", response_class=HTMLResponse)
async def root():
    return Path("static/index.html").read_text(encoding="utf-8")


@app.get("/agents")
async def list_agents():
    """Returns registered agent metadata so the UI can build the selector."""
    return {
        key: {k: v for k, v in meta.items() if not callable(v)}
        for key, meta in AGENTS.items()
    }


@app.get("/health")
async def health():
    return {"status": "ok", "agents": list(AGENTS.keys())}


# ── Shared helpers ───────────────────────────────────────────────────────────

def _build_report(filename, flags, agent_meta, doc):
    sev_summary = {
        sev: sum(1 for f in flags if f.get("severity") == sev)
        for sev in ("critical", "high", "medium", "low")
    }
    extra = agent_meta["extra_meta"](doc) if "extra_meta" in agent_meta else {}
    return {
        "file": filename,
        "total_flags": len(flags),
        "severity_summary": sev_summary,
        **extra,
        "flags": sorted(
            flags,
            key=lambda f: SEVERITY_ORDER.get(f.get("severity", "low"), 9),
        ),
    }


def _load_doc(agent_key: str, data: bytes, filename: str):
    """Load document using the right library for the agent type."""
    if agent_key == "docx":
        from docx import Document
        import io
        return Document(io.BytesIO(data))
    # Future: elif agent_key == "xlsx": ...
    raise HTTPException(400, f"No loader for agent '{agent_key}'")


def _save_doc(agent_key: str, doc, path: str):
    doc.save(path)


# ── Endpoints ────────────────────────────────────────────────────────────────

@app.post("/check/{agent_key}")
async def check_document(agent_key: str, file: UploadFile = File(...)):
    """Scan only — returns JSON flag report."""
    if agent_key not in AGENTS:
        raise HTTPException(404, f"Unknown agent '{agent_key}'. Available: {list(AGENTS)}")

    agent = AGENTS[agent_key]
    ext   = Path(file.filename).suffix.lower()
    if ext != agent["accept"]:
        raise HTTPException(400, f"Expected {agent['accept']} file, got {ext}")

    data = await file.read()
    try:
        doc   = _load_doc(agent_key, data, file.filename)
        flags = agent["run_checks"](doc)
        return JSONResponse(_build_report(file.filename, flags, agent, doc))
    except HTTPException:
        raise
    except Exception:
        raise HTTPException(500, f"Processing error:\n{traceback.format_exc()}")


@app.post("/fix/{agent_key}")
async def fix_document(agent_key: str, file: UploadFile = File(...)):
    """Fix + scan — returns zip with fixed file and JSON report."""
    if agent_key not in AGENTS:
        raise HTTPException(404, f"Unknown agent '{agent_key}'")

    agent = AGENTS[agent_key]
    ext   = Path(file.filename).suffix.lower()
    if ext != agent["accept"]:
        raise HTTPException(400, f"Expected {agent['accept']} file, got {ext}")

    data = await file.read()
    tmp  = tempfile.mkdtemp()
    try:
        stem        = Path(file.filename).stem
        fixed_path  = Path(tmp) / f"{stem}_fixed{ext}"
        report_path = Path(tmp) / f"{stem}_flags.json"
        zip_base    = Path(tmp) / f"{stem}_result"

        doc   = _load_doc(agent_key, data, file.filename)
        flags = agent["run_checks"](doc)

        report = _build_report(file.filename, flags, agent, doc)
        report_path.write_text(
            json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8"
        )

        agent["apply_fixes"](doc)
        _save_doc(agent_key, doc, str(fixed_path))

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
    # Note: tmp dir is cleaned by OS; for production add a BackgroundTask cleanup


# ── Vercel / Mangum adapter ──────────────────────────────────────────────────
# Vercel runs Python via its serverless runtime. Mangum wraps FastAPI for
# AWS Lambda-style handlers (which Vercel uses internally).
# Install: pip install mangum
# The `handler` name is what vercel.json points to.

try:
    from mangum import Mangum
    handler = Mangum(app, lifespan="off")
except ImportError:
    handler = None   # local run — mangum not needed


# ── Local dev entry point ────────────────────────────────────────────────────

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("server:app", host="0.0.0.0", port=port, reload=True)
