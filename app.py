#!/usr/bin/env python3
"""
app.py -- hosted web front end for the redirect-map scripts.

This is the browser/multi-user replacement for redirect_map_gui.py. It does
NOT reimplement any scraping logic -- it shells out to the exact same
sitemap_to_redirect_map.py / match_new_site.py scripts via subprocess,
the same way the tkinter GUI already did, and streams their stdout/stderr
into the browser instead of a tkinter log widget.

Run locally:
    pip install -r requirements.txt
    python app.py
    -> open http://localhost:5000

Deploy (Render/Railway/Fly/a VPS -- anything that runs a persistent
process, NOT a serverless platform with short request timeouts):
    gunicorn -w 1 --threads 8 -k gthread app:app

IMPORTANT: -w 1 (a single worker process) is required as-is, because job
state lives in this process's memory (a plain dict). Running more than one
worker would split jobs across processes that can't see each other's
state. If you outgrow a single worker later, move `jobs` into something
shared like Redis.

Environment variables:
    APP_PASSWORD  - shared password gate for the whole app. Strongly
                    recommended once this is hosted publicly, since this
                    tool fetches arbitrary URLs on someone's behalf.
                    Leave unset only for fully-trusted/local use.
    SECRET_KEY    - Flask session signing key. Set this explicitly in
                    production (any random string) so sessions survive a
                    restart; otherwise a new random one is generated each
                    time the process starts and everyone gets logged out.
    PORT          - port to listen on when run directly (default 5000).
"""

import ipaddress
import os
import shutil
import socket
import subprocess
import sys
import threading
import time
import uuid
from functools import wraps
from pathlib import Path
from urllib.parse import urlparse

from flask import (
    Flask, abort, jsonify, redirect, render_template, request,
    send_file, session, url_for,
)

BASE_DIR = Path(__file__).resolve().parent
JOBS_DIR = BASE_DIR / "jobs"
JOBS_DIR.mkdir(exist_ok=True)

STAGE1_SCRIPT = BASE_DIR / "sitemap_to_redirect_map.py"
STAGE2_SCRIPT = BASE_DIR / "match_new_site.py"

APP_PASSWORD = os.environ.get("APP_PASSWORD", "")
SECRET_KEY = os.environ.get("SECRET_KEY") or uuid.uuid4().hex

JOB_TTL_SECONDS = 6 * 60 * 60  # delete job folders/output files after 6h

app = Flask(__name__)
app.secret_key = SECRET_KEY

jobs = {}  # job_id -> dict(status, log:list[str], output_path, download_name, created)
jobs_lock = threading.Lock()


# --------------------------------------------------------------------- #
# Auth -- a single shared password, good enough for an internal/client
# tool. Swap for real accounts later if this needs per-user tracking.
# --------------------------------------------------------------------- #
def require_login(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if APP_PASSWORD and not session.get("logged_in"):
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapped


@app.route("/login", methods=["GET", "POST"])
def login():
    if not APP_PASSWORD:
        session["logged_in"] = True
        return redirect(url_for("index"))
    error = None
    if request.method == "POST":
        if request.form.get("password") == APP_PASSWORD:
            session["logged_in"] = True
            return redirect(request.args.get("next") or url_for("index"))
        error = "Incorrect password."
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# --------------------------------------------------------------------- #
# Small helpers
# --------------------------------------------------------------------- #
def _looks_like_url(u: str) -> bool:
    u = (u or "").strip()
    return u.startswith("http://") or u.startswith("https://")


def _resolve_is_public(url: str) -> bool:
    """
    Reject URLs whose hostname resolves to a private/loopback/link-local/
    reserved address. This app fetches whatever URL a logged-in user gives
    it -- without this check, someone could point it at the host's own
    internal network or a cloud metadata endpoint (a classic SSRF risk for
    any hosted "fetch this URL for me" tool). Legitimate client/business
    sites are always public, so this never gets in the way of real use.
    """
    host = urlparse(url).hostname
    if not host:
        return False
    try:
        infos = socket.getaddrinfo(host, None)
    except socket.gaierror:
        return False
    for info in infos:
        ip_str = info[4][0]
        try:
            ip = ipaddress.ip_address(ip_str)
        except ValueError:
            continue
        if ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved or ip.is_multicast:
            return False
    return True


def domain_to_filename(url: str) -> str:
    """Turn https://www.example.com/ into Example301RW.xlsx as a friendly default."""
    if not url:
        return "redirect_map.xlsx"
    cleaned = url.strip()
    for prefix in ("https://", "http://"):
        if cleaned.startswith(prefix):
            cleaned = cleaned[len(prefix):]
            break
    cleaned = cleaned.split("/")[0]
    if cleaned.startswith("www."):
        cleaned = cleaned[4:]
    core = cleaned.split(".")[0]
    core = "".join(ch for ch in core if ch.isalnum())
    if not core:
        return "redirect_map.xlsx"
    return f"{core[0].upper()}{core[1:]}301RW.xlsx"


# --------------------------------------------------------------------- #
# Job management -- each job runs the existing CLI script as a real
# subprocess in a background thread, exactly like the tkinter GUI did.
# --------------------------------------------------------------------- #
def new_job(download_name):
    job_id = uuid.uuid4().hex[:12]
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    with jobs_lock:
        jobs[job_id] = {
            "status": "running",
            "log": [],
            "output_path": None,
            "download_name": download_name,
            "created": time.time(),
        }
    return job_id, job_dir


def append_log(job_id, line):
    with jobs_lock:
        if job_id in jobs:
            jobs[job_id]["log"].append(line)


def run_subprocess_job(job_id, args, output_path, download_name):
    try:
        process = subprocess.Popen(
            args, cwd=str(BASE_DIR), stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT, text=True, bufsize=1,
        )
        for line in process.stdout:
            append_log(job_id, line.rstrip("\n"))
        process.wait()
        with jobs_lock:
            if process.returncode == 0 and output_path.exists():
                jobs[job_id]["status"] = "done"
                jobs[job_id]["output_path"] = str(output_path)
                jobs[job_id]["download_name"] = download_name
            else:
                jobs[job_id]["status"] = "error"
    except Exception as exc:  # noqa: BLE001
        append_log(job_id, f"Unexpected error: {exc}")
        with jobs_lock:
            jobs[job_id]["status"] = "error"


def cleanup_old_jobs():
    while True:
        time.sleep(30 * 60)
        cutoff = time.time() - JOB_TTL_SECONDS
        with jobs_lock:
            stale = [jid for jid, j in jobs.items() if j["created"] < cutoff]
        for jid in stale:
            job_dir = JOBS_DIR / jid
            try:
                if job_dir.exists():
                    for f in job_dir.iterdir():
                        f.unlink(missing_ok=True)
                    job_dir.rmdir()
            except OSError:
                pass
            with jobs_lock:
                jobs.pop(jid, None)


threading.Thread(target=cleanup_old_jobs, daemon=True).start()


# --------------------------------------------------------------------- #
# Routes
# --------------------------------------------------------------------- #
@app.route("/")
@require_login
def index():
    return render_template("index.html")


@app.route("/api/start-stage1", methods=["POST"])
@require_login
def start_stage1():
    old_url = (request.form.get("old_url") or "").strip()
    if not _looks_like_url(old_url):
        return jsonify(error="Enter a valid old-site URL starting with http:// or https://"), 400
    if not _resolve_is_public(old_url):
        return jsonify(error="That host doesn't resolve to a public website."), 400

    max_pages = request.form.get("max_pages", "300").strip() or "300"
    no_crawl = request.form.get("no_crawl") == "on"
    no_rest_api = request.form.get("no_rest_api") == "on"
    extra_urls_raw = request.form.get("extra_urls", "")
    extra_urls = [line.strip() for line in extra_urls_raw.splitlines() if line.strip()]

    download_name = domain_to_filename(old_url)
    job_id, job_dir = new_job(download_name)
    output_path = job_dir / "redirect_map.xlsx"

    args = [sys.executable, str(STAGE1_SCRIPT), old_url, "-o", str(output_path),
            "--max-pages", max_pages]
    if no_crawl:
        args.append("--no-crawl")
    if no_rest_api:
        args.append("--no-rest-api")
    if extra_urls:
        args += ["--extra-urls", *extra_urls]

    threading.Thread(
        target=run_subprocess_job,
        args=(job_id, args, output_path, download_name),
        daemon=True,
    ).start()

    return jsonify(job_id=job_id, download_name=download_name)


@app.route("/api/start-stage2", methods=["POST"])
@require_login
def start_stage2():
    new_url = (request.form.get("new_url") or "").strip()
    if not _looks_like_url(new_url):
        return jsonify(error="Enter a valid new-site URL starting with http:// or https://"), 400
    if not _resolve_is_public(new_url):
        return jsonify(error="That host doesn't resolve to a public website."), 400

    source_job_id = (request.form.get("source_job_id") or "").strip()
    uploaded = request.files.get("workbook_file")

    max_pages = request.form.get("max_pages2", "300").strip() or "300"
    no_crawl = request.form.get("no_crawl2") == "on"

    job_id, job_dir = new_job(None)
    workbook_path = job_dir / "redirect_map.xlsx"

    if uploaded and uploaded.filename:
        if not uploaded.filename.lower().endswith(".xlsx"):
            return jsonify(error="Please upload a .xlsx workbook."), 400
        uploaded.save(workbook_path)
        download_name = uploaded.filename
    elif source_job_id:
        with jobs_lock:
            src = jobs.get(source_job_id)
        if not src or not src.get("output_path") or not os.path.exists(src["output_path"]):
            return jsonify(error="That Step 1 workbook isn't available anymore -- please upload it instead."), 400
        shutil.copy(src["output_path"], workbook_path)
        download_name = src.get("download_name") or "redirect_map.xlsx"
    else:
        return jsonify(error="Provide the workbook from Step 1 or upload one."), 400

    with jobs_lock:
        jobs[job_id]["download_name"] = download_name

    args = [sys.executable, str(STAGE2_SCRIPT), str(workbook_path), new_url,
            "--max-pages", max_pages]
    if no_crawl:
        args.append("--no-crawl")

    threading.Thread(
        target=run_subprocess_job,
        args=(job_id, args, workbook_path, download_name),
        daemon=True,
    ).start()

    return jsonify(job_id=job_id, download_name=download_name)


@app.route("/api/status/<job_id>")
@require_login
def status(job_id):
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return jsonify(error="Unknown job"), 404
        return jsonify(
            status=job["status"],
            log="\n".join(job["log"]),
            download_ready=job["status"] == "done",
        )


@app.route("/download/<job_id>")
@require_login
def download(job_id):
    with jobs_lock:
        job = jobs.get(job_id)
    if not job or job["status"] != "done" or not job["output_path"]:
        abort(404)
    return send_file(
        job["output_path"], as_attachment=True,
        download_name=job.get("download_name") or "redirect_map.xlsx",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
