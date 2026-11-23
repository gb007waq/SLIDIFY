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
MAX_TRIES = 4
MAX_CUSTOM = 99        # largest custom slide count
BIG = 40               # above this many slides, plan headings first and write in batches
BATCH = 40             # content slides written per batch
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


def run_claude(prompt, cwd, session=None, timeout=CLAUDE_TIMEOUT):
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
                          timeout=timeout)
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
        raise RuntimeError("Could not read the presentation Claude wrote. Try again.")


def slide_limits(mode, n):
    """Turn the page's slide-count choice into (min, max) total slides."""
    if mode == "fewer":
        n = max(5, min(11, n or 8))
        return n, n
    if mode == "long":
        return 15, 20
    if mode == "custom":
        n = max(3, min(MAX_CUSTOM, n or 12))
        return n, n
    return 12, 12


def count_rule(job):
    lo, hi = job["limits"]
    return f"exactly {lo} slides in total." if lo == hi else f"between {lo} and {hi} slides in total."


def content_rule(job):
    lo, hi = job["limits"]
    extra = 2 + (1 if job["agenda"] else 0)
    a, b = max(1, lo - extra), max(1, hi - extra)
    if job["agenda"]:
        b = max(a, min(b, bd.MAX_AGENDA))
    return str(a) if a == b else f"{a} to {b}"


def first_prompt(job):
    today = dt.date.today().strftime("%b, %Y").upper()
    lines = [
        "You are SLIDIFY, the owner's presentation agent.",
        "Write a deck in the owner's signature style, described in this style guide:",
        "",
        "<style_guide>", STYLE, "</style_guide>",
        "",
        "Builder limits: 3-7 lines per slide or per column; at most 12 words per line;",
        "headings at most 5 words; every line must fit on one line of a 16:9 slide at 20 pt",
        "(two-column lines are half width). The agenda lists up to 14 content headings.",
        "",
        f"SLIDE COUNT: {count_rule(job)}",
        "The title slide, the Sequence agenda (if any) and the THANK YOU slide all count.",
        f"So the 'slides' array must hold {content_rule(job)} content slides.",
        "Spread the material evenly; never pad with filler slides.",
        "",
        "Reply with ONLY the JSON spec in a ```json block, nothing else. Schema:",
        '{"title": str, "date": str, "agenda": bool, "slides": [{"heading": str, "lines": [str]}'
        ' or {"heading": str, "left": [str], "right": [str]}]}',
        "",
        f"Date to use: {job['date'] or today}.",
        f"Include a Sequence agenda slide: {'yes' if job['agenda'] else 'no'}.",
    ]
    if job["brief"]:
        lines += ["", "What the deck is about / instructions from the owner:", job["brief"]]
        if job.get("spoken"):
            lines += ["(This was spoken and transcribed by speech recognition: ignore filler words,",
                      "fix obvious recognition mistakes. It may be in any language, e.g. Urdu or Pashto:",
                      "understand it, but always write the slides in English.)"]
    if job["sources"]:
        lines += ["", "Source material - read every file with the Read tool, then build the deck from it:"]
        lines += [f"- {p}" for p in job["sources"]]
    return "\n".join(lines)


SLIDE_SCHEMA = ('{"heading": str, "lines": [str]} or {"heading": str, "left": [str], "right": [str]}')


def per_slide_errors(spec):
    """{content index (0-based): [errors]} using the builder's own checks."""
    bad = {}
    for i, s in enumerate(spec.get("slides", [])):
        errs = bd.validate({"title": "x", "slides": [s], "style": spec.get("style")}, 1, 10**6)
        if errs:
            bad[i] = [e.split(": ", 1)[-1] for e in errs]
    return bad


def deck_errors(spec, limits):
    return [e for e in bd.validate(spec, *limits) if not e.startswith("Slide ")]


def fix_slides(jid, spec, session, folder):
    """Ask Claude to rewrite only the slides that break the rules (used for long decks)."""
    for attempt in range(MAX_TRIES):
        bad = per_slide_errors(spec)
        if not bad:
            return spec, session
        set_job(jid, message=f"Tightening {len(bad)} slide(s) that break the style...")
        listing = "\n".join(
            f"Slide {i + 1}: {json.dumps(spec['slides'][i], ensure_ascii=False)}\n  problems: " + "; ".join(e)
            for i, e in list(bad.items())[:60])
        text, session = run_claude(
            "These content slides break the style rules. Rewrite each one, keeping its heading and meaning;"
            " shorten lines (never split them).\n" + listing +
            '\nReply with ONLY JSON: {"slides": {"<slide number>": ' + SLIDE_SCHEMA + "}}",
            folder, session)
        for k, s in extract_spec(text).get("slides", {}).items():
            idx = int(k) - 1
            if 0 <= idx < len(spec["slides"]) and isinstance(s, dict):
                spec["slides"][idx] = s
    bad = per_slide_errors(spec)
    if bad:
        first = next(iter(bad.items()))
        raise RuntimeError(f"{len(bad)} slide(s) still break the style, e.g. slide {first[0] + 1}: {first[1][0]}")
    return spec, session


def write_big(jid, job, folder):
    """Long decks: plan every heading first, then write the slides in batches."""
    total = job["limits"][0]
    k = total - 2 - (1 if job["agenda"] else 0)
    set_job(jid, message=f"Planning {k} slide headings...")
    plan = first_prompt(job).replace(
        "Reply with ONLY the JSON spec in a ```json block, nothing else. Schema:",
        "First, plan the deck only. Reply with ONLY JSON in a ```json block, nothing else. Schema:")
    plan = plan.replace(
        '{"title": str, "date": str, "agenda": bool, "slides": [{"heading": str, "lines": [str]}'
        ' or {"heading": str, "left": [str], "right": [str]}]}',
        '{"title": str, "date": str, "agenda": bool, "headings": [str]}  '
        f"- exactly {k} distinct headings, in presentation order, at most 5 words each.")
    text, session = run_claude(plan, folder, None, CLAUDE_TIMEOUT + 2 * k)
    outline = extract_spec(text)
    heads = [h for h in outline.get("headings", []) if str(h).strip()]
    if len(heads) != k:
        text, session = run_claude(
            f"You gave {len(heads)} headings; I need exactly {k}. Reply with the corrected JSON plan only.",
            folder, session, CLAUDE_TIMEOUT + 2 * k)
        outline = extract_spec(text)
        heads = [h for h in outline.get("headings", []) if str(h).strip()]
    if len(heads) < k:
        raise RuntimeError(f"Could only plan {len(heads)} of {k} slides. Try a broader topic or fewer slides.")
    heads = heads[:k]

    slides = []
    for start in range(0, k, BATCH):
        part = heads[start:start + BATCH]
        set_job(jid, message=f"Writing slides {start + 1}-{start + len(part)} of {k}...")
        ask = ("Now write the content slides for these headings, in this order, using the headings exactly:\n" +
               "\n".join(f"{start + i + 1}. {h}" for i, h in enumerate(part)) +
               f"\nReply with ONLY JSON: {{\"slides\": [{SLIDE_SCHEMA}, ...]}} - exactly {len(part)} slides.")
        got = []
        for _ in range(2):
            text, session = run_claude(ask, folder, session)
            got = [s for s in extract_spec(text).get("slides", []) if isinstance(s, dict)]
            if len(got) >= len(part):
                break
            ask = f"That had {len(got)} slides; I need exactly {len(part)}. " + ask
        if len(got) < len(part):
            raise RuntimeError(f"SLIDIFY stopped early while writing slides {start + 1}-{start + len(part)}.")
        for h, s in zip(part, got):
            s["heading"] = h
            slides.append(s)

    spec = {"title": outline.get("title") or "Presentation", "date": outline.get("date") or job["date"],
            "agenda": bool(job["agenda"]), "slides": slides, "style": job.get("style")}
    return spec, session


def apply_edits(spec, edits):
    """Apply {"op": replace|delete|insert_after, "n": slide number, ...} edits (1-based content numbers)."""
    slides = spec["slides"]
    for e in sorted(edits, key=lambda e: int(e.get("n", 0)), reverse=True):
        n, op = int(e.get("n", 0)), e.get("op")
        if op == "replace" and 1 <= n <= len(slides) and isinstance(e.get("slide"), dict):
            slides[n - 1] = e["slide"]
        elif op == "delete" and 1 <= n <= len(slides):
            del slides[n - 1]
        elif op == "insert_after" and 0 <= n <= len(slides):
            new = [s for s in e.get("slides", []) if isinstance(s, dict)]
            slides[n:n] = new
    return spec


def revise_big(jid, job, folder, feedback):
    """Long decks: Claude reads the current deck from a file and replies with edits only."""
    current = folder / "current.json"
    current.write_text(json.dumps(job["spec"], indent=1, ensure_ascii=False), encoding="utf-8")
    text, session = run_claude(
        f"The current deck is in {current} (read it with the Read tool; 'slides' are numbered from 1)."
        f"\nChange it as the owner asks, keeping every style rule:\n{feedback}\n"
        "Reply with ONLY JSON listing the edits, nothing else:\n"
        '{"title": str (optional, only if it changes), "agenda": bool (optional), "edits": ['
        '{"op": "replace", "n": int, "slide": ' + SLIDE_SCHEMA + '}, {"op": "delete", "n": int}, '
        '{"op": "insert_after", "n": int (0 = at the start), "slides": [' + SLIDE_SCHEMA + "]}]}"
        "\nNumbers refer to the current deck before any edit.",
        folder, job.get("session"), CLAUDE_TIMEOUT * 2)
    reply = extract_spec(text)
    spec = json.loads(json.dumps(job["spec"]))
    spec = apply_edits(spec, reply.get("edits", []))
    if reply.get("title"):
        spec["title"] = reply["title"]
    if "agenda" in reply:
        spec["agenda"] = bool(reply["agenda"])
    if spec.get("agenda") and len(spec["slides"]) > bd.MAX_AGENDA:
        spec["agenda"] = False
    return spec, session


def make_deck(jid, prompt, session=None, limits=(1, bd.MAX_SLIDES), feedback=None):
    job = jobs[jid]
    folder = JOBS / jid
    set_job(jid, status="working", message="SLIDIFY is reading and writing your slides...")
    wait = CLAUDE_TIMEOUT + 12 * min(limits[1], 300)   # more slides, more time
    big_new = feedback is None and limits[0] == limits[1] and limits[0] > BIG
    big_edit = feedback is not None and len(job.get("spec", {}).get("slides", [])) > BIG

    if big_new or big_edit:
        spec, session = (write_big(jid, job, folder) if big_new
                         else revise_big(jid, job, folder, feedback))
        spec, session = fix_slides(jid, spec, session, folder)
        errors = deck_errors(spec, limits)
        if errors:
            raise RuntimeError("; ".join(errors))
    else:
        text, session = run_claude(prompt, folder, session, wait)
        spec = extract_spec(text)
        for attempt in range(1, MAX_TRIES + 1):
            spec.pop("closing", None)
            spec["style"] = job.get("style")
            errors = bd.validate(spec, *limits)
            if not errors:
                break
            if attempt == MAX_TRIES:
                raise RuntimeError("Some lines still break the style: " + "; ".join(errors[:4]))
            set_job(jid, message=f"Tightening {len(errors)} line(s) that break the style...")
            fix = ("The builder rejected the deck:\n" + "\n".join("- " + e for e in errors) +
                   "\nShorten those lines (never split them) and reply with the full corrected JSON spec only.")
            text, session = run_claude(fix, folder, session, wait)
            spec = extract_spec(text)

    spec.pop("closing", None)
    spec["style"] = job.get("style")
    set_job(jid, message=f"Building PowerPoint, PDF, pictures & Word ({bd.slide_count(spec)} slides)...")
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


def worker(jid, prompt, session=None, limits=(1, bd.MAX_SLIDES), feedback=None):
    try:
        make_deck(jid, prompt, session, limits, feedback)
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
        n = int(request.form.get("slides_n") or 0)
    except ValueError:
        n = 0
    limits = slide_limits(request.form.get("slides_mode", "standard"), n)
    default_style, styles = bd.load_styles()
    style = request.form.get("style") or default_style
    if style not in styles or not bd.style_ready(styles[style]):
        style = default_style
    agenda = request.form.get("agenda") == "true"
    if agenda and limits[0] - 3 > bd.MAX_AGENDA:
        agenda = False  # too many slides to list on one agenda slide
    job = {"id": jid, "status": "queued", "message": "Starting...", "brief": brief,
           "sources": sources, "date": (request.form.get("date") or "").strip(),
           "agenda": agenda, "limits": limits, "spoken": request.form.get("spoken") == "true",
           "style": style}
    with lock:
        jobs[jid] = job
    threading.Thread(target=worker, args=(jid, first_prompt(job), None, limits), daemon=True).start()
    return jsonify(id=jid)


@app.get("/api/styles")
def styles_list():
    """Font styles for the picker; a style without a font shows as 'coming soon'."""
    default, styles = bd.load_styles()
    return jsonify(default=default, styles=[
        {"id": k, "label": v.get("label", k), "font": v.get("font"), "ready": bd.style_ready(v),
         "title_pt": v.get("title_pt"), "heading_pt": v.get("heading_pt"), "body_pt": v.get("body_pt")}
        for k, v in styles.items()])


@app.post("/api/translate")
def translate():
    """Faithful translation of spoken text into English (no summarising, nothing added)."""
    data = request.json or {}
    text = (data.get("text") or "").strip()
    lang = (data.get("language") or "the speaker's language").strip()[:60]
    if not text:
        return jsonify(error="Nothing to translate."), 400
    folder = JOBS / "translate"
    folder.mkdir(exist_ok=True)
    prompt = (
        "The text below is what a speech recogniser heard. Work out what the speaker actually said "
        "and give it in English.\n"
        f"The recogniser was set to {lang}, but the speaker may have used any language in the world, "
        "or a mix of languages (for example Pashto with English words). If the recogniser was set "
        "to a different language from the one spoken, it will have written the sounds as look-alike "
        "words or in the wrong script (for example Pashto 'mala da 12 slides presentation jor ka' "
        "heard as English words). Read it phonetically, work out which language was really spoken, "
        "recover the speaker's real words, and translate them.\n"
        "Rules: translate exactly what was said - same meaning, details, names and numbers. "
        "Do not summarise, explain, answer, add or leave out anything. Keep parts that are "
        "already English as they are. Write every number as digits (10, not ten). "
        "Reply with ONLY the English, no quotes, no notes.\n\n"
        f"<speech>\n{text}\n</speech>")
    try:
        english, _ = run_claude(prompt, folder, None, 180)
    except Exception as e:  # noqa: BLE001
        return jsonify(error=f"Translation failed: {e}"), 502
    return jsonify(english=english.strip().strip('"'))


@app.post("/api/jobs/<jid>/revise")
def revise(jid):
    job = jobs.get(jid) or abort(404)
    if job["status"] == "working":
        return jsonify(error="Still working on the last change."), 409
    feedback = (request.json or {}).get("feedback", "").strip()
    if not feedback:
        return jsonify(error="Say what to change."), 400
    prompt = ("Change the deck as the owner asks, keeping every style rule:\n" + feedback +
              f"\n\nKeep the slide count rule ({count_rule(job)}) unless the owner asks for a different number"
              f" (3 to {MAX_CUSTOM} slides allowed)."
              "\n\nCurrent deck:\n```json\n" + json.dumps(job.get("spec", {}), ensure_ascii=False) +
              "\n```\nReply with the full updated JSON spec only.")
    set_job(jid, status="working", message="Applying your changes...")
    threading.Thread(target=worker, args=(jid, prompt, job.get("session"), (3, MAX_CUSTOM), feedback),
                     daemon=True).start()
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
