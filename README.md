# PPM Formatting Agent — Cloud + Local

## Files

```
ppm_cloud_app/
├── server.py               ← FastAPI app (plugin architecture)
├── ppm_format_agent.py     ← Word doc agent (spec + fixer)
├── requirements.txt        ← Python deps
├── vercel.json             ← Vercel deployment config
├── README.md               ← This file
└── static/
    └── index.html          ← Web UI (agent selector + flag table)
```

---

## Run locally

```bash
pip install -r requirements.txt
python server.py
# → open http://localhost:8000
```

---

## Deploy to Vercel (free tier, recommended)

```bash
# 1. Install Vercel CLI
npm install -g vercel

# 2. From inside this folder
vercel

# Follow the prompts:
#   Set up and deploy? Yes
#   Which scope? (your account)
#   Link to existing project? No
#   Project name: ppm-formatting-agent
#   Directory: ./  (current folder)

# 3. That's it. Vercel gives you a URL like:
#    https://ppm-formatting-agent.vercel.app
```

> **File size limit on Vercel free tier: 4.5 MB** per upload (serverless payload limit).
> Most PPM sector .docx files are well under this. If a file exceeds it, use Railway instead.

---

## Deploy to Railway (no file size limit, $5/month)

```bash
# 1. Install Railway CLI
npm install -g @railway/cli

# 2. Login and deploy
railway login
railway init       # name your project
railway up

# Railway auto-detects Python + requirements.txt.
# Set the start command in dashboard:
#   python server.py
```

---

## Deploy to Render (free tier with sleep)

1. Push this folder to a GitHub repo
2. Go to https://render.com → New Web Service
3. Connect your repo
4. Build command: `pip install -r requirements.txt`
5. Start command: `python server.py`
6. Done — Render gives you a public URL

> Free tier sleeps after 15 min of inactivity (cold start ~30s).
> Use the paid tier ($7/month) for always-on.

---

## Adding Excel / PowerPoint support later

1. Create `ppm_format_agent_xlsx.py` with the same interface:
   - `run_checks(doc)` → list of flag dicts
   - `apply_fixes(doc)` → modifies in-place

2. Register it in `server.py` → `AGENTS` dict:
   ```python
   from ppm_format_agent_xlsx import run_checks as xlsx_checks, apply_fixes as xlsx_fixes

   AGENTS["xlsx"] = {
       "label":       "Excel Spreadsheet",
       "description": "PPM financial model .xlsx files",
       "icon":        "📊",
       "accept":      ".xlsx",
       "run_checks":  xlsx_checks,
       "apply_fixes": xlsx_fixes,
   }
   ```

3. Remove `disabled` class from the Excel tab in `static/index.html`.

4. Redeploy — no other changes needed.

---

## Spec reference (Word)

| Element | Rule |
|---|---|
| Font | Lato |
| Body spacing | 1.15× |
| Table spacing | Single |
| H1 (H2 present) | 16pt max |
| H1 (no H2) | 14pt |
| H2 | 14pt bold |
| H3 | 12pt bold |
| H4/H5 | 11pt · #666666 |
| H6 | 11pt italic · #666666 |
| Header fill | #073763 |
| Header borders | Solid white · ½pt (sz=4) |
| Body cell borders | Dashed · #6fa8dc · ½pt (sz=4) |
| Header alignment | Center · Middle |
| Text cell alignment | Left · Middle |
| Numeric cell alignment | Right · Middle |
| Even row fill | #efefef |
| Outer table border | None |
