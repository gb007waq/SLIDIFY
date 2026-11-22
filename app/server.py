"""SLIDIFY app: a local web app that turns a topic, notes, Word files, PDFs
or pictures into a deck in the signature style, and exports it as
PowerPoint, PDF, slide pictures and Word.

The slide text is written by Claude Code on this computer (`claude -p`),
using your existing Claude sign-in. Run with start.ps1, then open
http://127.0.0.1:8765
"""

import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import uuid
from pathlib import Path

from flask import Flask, abort, jsonify, request, send_file, send_from_directory

ROOT = Path(__file__).resolve().parent
SKILL = ROOT.parent / "skills" / "slidify"
sys.path.insert(0, str(SKILL / "scripts"))
import build_deck as bd  # noqa: E402
import export_deck as ex  # noqa: E402

JOBS = ROOT / "jobs"
JOBS.mkdir(exist_ok=True)
STYLE = (SKILL / "STYLE.md").read_text(encoding="utf-8")

TEXT_TYPES = {".txt", ".md"}
DOC_TYPES = {".docx"}
READ_TYPES = {".pdf", ".png", ".jpg", ".jpeg", ".webp", ".gif"}
ALLOWED = TEXT_TYPES | DOC_TYPES | READ_TYPES
MAX_TRIES = 3
CLAUDE_TIMEOUT = 600

app = Flask(__name__, static_folder=str(ROOT / "static"), static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = 60 * 1024 * 1024

jobs = {}
lock = threading.Lock()


# ---- Helpers ---------------------------------------------------------------
def set_job(jid, **kw):
    with lock:
        jobs[jid].update(kw)


def safe_name(name):
    name = re.sub(r"[^\w .&()-]+", "", name).strip() or "file"
    return name[:80]


def docx_to_text(path):
    from docx import Document
    doc = Document(path)
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for t in doc.tables:
        for row in t.rows:
            parts.append(" | ".join(c.text.strip() for c in row.cells))
    return "\n".join(parts)


def claude_exe():
    exe = shutil.which("claude")
    if not exe:
        raise RuntimeError("Claude Code was not found. Install it and sign in, then restart the app.")
    return exe


def run_claude(prompt, cwd, session=None):
    """Run Claude Code headless. Returns (text, session_id)."""
    cmd = [claude_exe(), "-p", "--output-format", "json",
           "--tools", "Read", "--allowedTools", "Read",
           "--add-dir", str(cwd)]
    if session:
        cmd += ["--resume", session]
    env = {k: v for k, v in os.environ.items()
           if not (k.startswith("CLAUDECODE") or k.startswith("CLAUDE_CODE_"))}
    proc = subprocess.run(cmd, input=prompt, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", cwd=cwd, env=env,
                          timeout=CLAUDE_TIMEOUT)
    out = proc.stdout.strip()
    try:
        data = json.loads(out)
    except json.JSONDecodeError:
        raise RuntimeError("Claude Code did not answer: " + (proc.stderr or out)[-400:])
    if data.get("is_error"):
        raise RuntimeError("Claude Code reported an error: " + str(data.get("result"))[:400])
    return data.get("result", ""), data.get("session_id")


def extract_spec(text):
    m = re.search(r"```(?:json)?\s*(\{.*\})\s*```", text, re.S)
    raw = m.group(1) if m else text[text.find("{"): text.rfind("}") + 1]
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        raise RuntimeError("Could not read the deck Claude wrote. Try again.")


def first_prompt(job):
    today = dt.date.today().strftime("%b, %Y").upper()
    lines = [
        "You are SLIDIFY, the owner's presentation agent.",
        "Write a deck in the owner's signature style, described in this style guide:",
        "",
        "<style_guide>", STYLE, "</style_guide>",
        "",
        "Builder limits: 12 slides total (title, optional Sequence agenda and THANK YOU count);",
        "3-7 lines per slide or per column; at most 12 words per line; headings at most 5 words;",
        "every line must fit on one line of a 16:9 slide at 20 pt (two-column lines are half width).",
        "An agenda is only allowed with 7 or fewer content slides.",
        "",
        "Reply with ONLY the JSON spec in a ```json block, nothing else. Schema:",
        '{"title": str, "date": str, "agenda": bool, "slides": [{"heading": str, "lines": [str]}'
        ' or {"heading": str, "left": [str], "right": [str]}]}',
        "",
        f"Date to use: {job['date'] or today}.",
        f"Include a Sequence agenda slide: {'yes' if job['agenda'] else 'no'}.",
    ]
    if job["max_slides"]:
        lines.append(f"Total slides including title and THANK YOU: {job['max_slides']} at most.")
    if job["brief"]:
        lines += ["", "What the deck is about / instructions from the owner:", job["brief"]]
    if job["sources"]:
        lines += ["", "Source material - read every file with the Read tool, then build the deck from it:"]
        lines += [f"- {p}" for p in job["sources"]]
    return "\n".join(lines)


def make_deck(jid, prompt, session=None):
    job = jobs[jid]
    folder = JOBS / jid
    set_job(jid, status="working", message="SLIDIFY is reading and writing your slides...")
    text, session = run_claude(prompt, folder, session)
    spec = extract_spec(text)
    for attempt in range(1, MAX_TRIES + 1):
        spec.pop("closing", None)
        errors = bd.validate(spec)
        if not errors:
            break
        if attempt == MAX_TRIES:
            raise RuntimeError("Some lines still break the style: " + "; ".join(errors[:4]))
        set_job(jid, message=f"Tightening {len(errors)} line(s) that break the style...")
        fix = ("The builder rejected the deck:\n" + "\n".join("- " + e for e in errors) +
               "\nShorten those lines (never split them) and reply with the full corrected JSON spec only.")
        text, session = run_claude(fix, folder, session)
        spec = extract_spec(text)

    set_job(jid, message="Building PowerPoint, PDF, pictures & Word...")
    version = job.get("version", 0) + 1
    out_dir = folder / f"v{version}"
    stem = safe_name(bd.clean(spec["title"]))
    out = ex.export_all(spec, out_dir, stem=stem)
    (out_dir / "spec.json").write_text(json.dumps(spec, indent=2, ensure_ascii=False), encoding="utf-8")
    base = f"/files/{jid}/v{version}/"
    set_job(jid, status="done", message="Ready", session=session, spec=spec, version=version,
            title=spec["title"],
            slides=[base + "slides/" + p.name for p in out["png"]],
            downloads={k: base + out[k].name for k in ("pptx", "pdf", "zip", "docx")})


def worker(jid, prompt, session=None):
    try:
        make_deck(jid, prompt, session)
    except subprocess.TimeoutExpired:
        set_job(jid, status="error", message="SLIDIFY took too long. Try a shorter brief or fewer files.")
    except Exception as e:  # noqa: BLE001
        set_job(jid, status="error", message=str(e))


# ---- Routes ----------------------------------------------------------------
@app.get("/")
def index():
    return send_from_directory(ROOT / "static", "index.html")


@app.post("/api/jobs")
def create_job():
    brief = (request.form.get("brief") or "").strip()
    files = [f for f in request.files.getlist("files") if f and f.filename]
    if not brief and not files:
        return jsonify(error="Write a topic or add at least one file."), 400

    jid = uuid.uuid4().hex[:12]
    folder = JOBS / jid / "inputs"
    folder.mkdir(parents=True)
    sources = []
    for f in files:
        ext = Path(f.filename).suffix.lower()
        if ext not in ALLOWED:
            return jsonify(error=f"{f.filename}: use Word, PDF, pictures or text files."), 400
        path = folder / safe_name(f.filename)
        f.save(path)
        if ext in DOC_TYPES:
            txt = path.with_suffix(".txt")
            txt.write_text(docx_to_text(path), encoding="utf-8")
            sources.append(str(txt))
        else:
            sources.append(str(path))

    try:
        max_slides = max(4, min(12, int(request.form.get("max_slides") or 12)))
    except ValueError:
        max_slides = 12
    job = {"id": jid, "status": "queued", "message": "Starting...", "brief": brief,
           "sources": sources, "date": (request.form.get("date") or "").strip(),
           "agenda": request.form.get("agenda") == "true", "max_slides": max_slides}
    with lock:
        jobs[jid] = job
    threading.Thread(target=worker, args=(jid, first_prompt(job)), daemon=True).start()
    return jsonify(id=jid)


@app.post("/api/jobs/<jid>/revise")
def revise(jid):
    job = jobs.get(jid) or abort(404)
    if job["status"] == "working":
        return jsonify(error="Still working on the last change."), 409
    feedback = (request.json or {}).get("feedback", "").strip()
    if not feedback:
        return jsonify(error="Say what to change."), 400
    prompt = ("Change the deck as the owner asks, keeping every style rule:\n" + feedback +
              "\n\nCurrent deck:\n```json\n" + json.dumps(job.get("spec", {}), ensure_ascii=False) +
              "\n```\nReply with the full updated JSON spec only.")
    set_job(jid, status="working", message="Applying your changes...")
    threading.Thread(target=worker, args=(jid, prompt, job.get("session")), daemon=True).start()
    return jsonify(id=jid)


@app.get("/api/jobs/<jid>")
def status(jid):
    job = jobs.get(jid) or abort(404)
    keys = ("id", "status", "message", "title", "slides", "downloads", "version")
    return jsonify({k: job.get(k) for k in keys})


@app.get("/files/<jid>/<path:name>")
def files(jid, name):
    folder = (JOBS / jid).resolve()
    path = (folder / name).resolve()
    if folder not in path.parents or not path.is_file():
        abort(404)
    return send_file(path, as_attachment=request.args.get("download") == "1",
                     download_name=path.name)


if __name__ == "__main__":
    port = int(os.environ.get("SLIDIFY_PORT", "8765"))
    print(f"SLIDIFY is running: http://127.0.0.1:{port}")
    app.run(host="127.0.0.1", port=port, threaded=True)
