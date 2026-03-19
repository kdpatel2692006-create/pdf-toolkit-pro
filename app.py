"""
PDF Toolkit Pro - Complete Server
Works both locally and on cloud (Render/Railway)
"""

import os
import io
import sys
import uuid
import json
import shutil
import tempfile
import threading
import webbrowser
from pathlib import Path
from functools import wraps

from flask import (
    Flask, request, jsonify, send_file, render_template,
    session, redirect, url_for, make_response
)
from werkzeug.utils import secure_filename
from pdf_engine import PDFEngine

app = Flask(__name__)

# Secret key for sessions (password protection)
app.secret_key = os.environ.get(
    "SECRET_KEY",
    "pdf-toolkit-local-dev-key-change-in-production"
)

# Password protection (set APP_PASSWORD on Render to enable)
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "pdftoolkit_up")
OUTPUT_DIR = os.path.join(tempfile.gettempdir(), "pdftoolkit_out")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

app.config["MAX_CONTENT_LENGTH"] = 150 * 1024 * 1024  # 150 MB

engine = PDFEngine()
edit_sessions = {}


# ───────── Password Protection ─────────
def check_auth():
    """Check if user is authenticated (only if password is set)."""
    if not APP_PASSWORD:
        return True  # No password set = open access
    return session.get("authenticated", False)


@app.before_request
def require_login():
    """Block all requests if password is set and user not logged in."""
    # Allow these paths without login
    open_paths = ["/login", "/static/", "/manifest.json"]
    if any(request.path.startswith(p) for p in open_paths):
        return
    if APP_PASSWORD and not session.get("authenticated"):
        if request.path.startswith("/api/"):
            return jsonify(error="Not authenticated"), 401
        return redirect("/login")


@app.route("/login", methods=["GET", "POST"])
def login():
    if not APP_PASSWORD:
        return redirect("/")
    if request.method == "POST":
        pwd = request.form.get("password", "")
        if pwd == APP_PASSWORD:
            session["authenticated"] = True
            session.permanent = True
            return redirect("/")
        return '''<!DOCTYPE html><html><head>
            <meta name="viewport" content="width=device-width,initial-scale=1">
            <title>PDF Toolkit Pro - Login</title>
            <style>
            *{margin:0;padding:0;box-sizing:border-box}
            body{font-family:'Segoe UI',system-ui,sans-serif;background:#f0f2f5;
                 display:flex;align-items:center;justify-content:center;min-height:100vh}
            .box{background:#fff;border-radius:20px;padding:40px;max-width:400px;
                 width:90%;box-shadow:0 4px 20px rgba(0,0,0,.1);text-align:center}
            h1{font-size:1.5rem;margin-bottom:8px;color:#1a1a2e}
            p{color:#888;margin-bottom:24px;font-size:.9rem}
            .err{background:#ffe0e0;color:#e74c3c;padding:10px;border-radius:8px;
                 margin-bottom:16px;font-size:.85rem}
            input{width:100%;padding:14px;border:2px solid #ddd;border-radius:10px;
                  font-size:1rem;margin-bottom:16px;text-align:center}
            input:focus{border-color:#667eea;outline:none}
            button{width:100%;padding:14px;background:linear-gradient(135deg,#667eea,#764ba2);
                   color:#fff;border:none;border-radius:10px;font-size:1rem;font-weight:700;cursor:pointer}
            </style></head><body>
            <div class="box">
            <h1>📄 PDF Toolkit Pro</h1>
            <p>Enter password to continue</p>
            <div class="err">❌ Wrong password. Try again.</div>
            <form method="POST">
            <input type="password" name="password" placeholder="Enter password" autofocus required>
            <button type="submit">🔓 Unlock</button>
            </form></div></body></html>'''
    return '''<!DOCTYPE html><html><head>
        <meta name="viewport" content="width=device-width,initial-scale=1">
        <title>PDF Toolkit Pro - Login</title>
        <style>
        *{margin:0;padding:0;box-sizing:border-box}
        body{font-family:'Segoe UI',system-ui,sans-serif;background:#f0f2f5;
             display:flex;align-items:center;justify-content:center;min-height:100vh}
        .box{background:#fff;border-radius:20px;padding:40px;max-width:400px;
             width:90%;box-shadow:0 4px 20px rgba(0,0,0,.1);text-align:center}
        h1{font-size:1.5rem;margin-bottom:8px;color:#1a1a2e}
        p{color:#888;margin-bottom:24px;font-size:.9rem}
        input{width:100%;padding:14px;border:2px solid #ddd;border-radius:10px;
              font-size:1rem;margin-bottom:16px;text-align:center}
        input:focus{border-color:#667eea;outline:none}
        button{width:100%;padding:14px;background:linear-gradient(135deg,#667eea,#764ba2);
               color:#fff;border:none;border-radius:10px;font-size:1rem;font-weight:700;cursor:pointer}
        </style></head><body>
        <div class="box">
        <h1>📄 PDF Toolkit Pro</h1>
        <p>Enter password to continue</p>
        <form method="POST">
        <input type="password" name="password" placeholder="Enter password" autofocus required>
        <button type="submit">🔓 Unlock</button>
        </form></div></body></html>'''


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")


# ───────── Helpers ─────────
def _save(file):
    name = secure_filename(file.filename) or "upload"
    unique = f"{uuid.uuid4().hex}_{name}"
    path = os.path.join(UPLOAD_DIR, unique)
    file.save(path)
    return path


def _out(original_name, suffix="", ext=None):
    stem = Path(original_name).stem
    if ext is None:
        ext = Path(original_name).suffix
    return os.path.join(OUTPUT_DIR, f"{uuid.uuid4().hex}_{stem}{suffix}{ext}")


def _cleanup(*paths, delay=5):
    def _do():
        import time
        time.sleep(delay)
        for p in paths:
            try:
                if os.path.isdir(p):
                    shutil.rmtree(p, ignore_errors=True)
                elif os.path.isfile(p):
                    os.remove(p)
            except OSError:
                pass
    threading.Thread(target=_do, daemon=True).start()


def _send_and_cleanup(filepath, download_name, mimetype=None):
    resp = send_file(filepath, as_attachment=True,
                     download_name=download_name, mimetype=mimetype)
    _cleanup(filepath, delay=15)
    return resp


# ───────── Pages ─────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/manifest.json")
def manifest():
    return send_file("static/manifest.json",
                     mimetype="application/manifest+json")


# ───────── API Routes ─────────
@app.route("/api/preview", methods=["POST"])
def preview():
    f = request.files.get("file")
    page = int(request.form.get("page", 0))
    if not f:
        return jsonify(error="No file"), 400
    path = _save(f)
    try:
        img = engine.get_preview(path, page_num=page)
        img_path = _out(f.filename, f"_p{page}", ".png")
        with open(img_path, "wb") as fp:
            fp.write(img)
        _cleanup(path)
        _cleanup(img_path, delay=60)
        return send_file(img_path, mimetype="image/png")
    except Exception as e:
        _cleanup(path)
        return jsonify(error=str(e)), 500


@app.route("/api/metadata", methods=["POST"])
def metadata():
    f = request.files.get("file")
    if not f:
        return jsonify(error="No file"), 400
    path = _save(f)
    try:
        meta = engine.get_metadata(path)
        _cleanup(path)
        return jsonify(meta)
    except Exception as e:
        _cleanup(path)
        return jsonify(error=str(e)), 500


@app.route("/api/merge", methods=["POST"])
def merge():
    files = request.files.getlist("files")
    if len(files) < 2:
        return jsonify(error="Upload at least 2 PDFs"), 400
    paths = [_save(f) for f in files]
    output = _out("merged", ext=".pdf")
    try:
        engine.merge(paths, output)
        _cleanup(*paths)
        return _send_and_cleanup(output, "merged.pdf")
    except Exception as e:
        _cleanup(*paths)
        return jsonify(error=str(e)), 500


@app.route("/api/split", methods=["POST"])
def split():
    f = request.files.get("file")
    mode = request.form.get("mode", "all")
    if not f:
        return jsonify(error="No file"), 400
    path = _save(f)
    try:
        pages = None
        every_n = 1
        if mode == "ranges":
            raw = request.form.get("ranges", "1-1")
            pages = []
            for r in raw.split(","):
                parts = r.strip().split("-")
                pages.append((int(parts[0]), int(parts[1])))
        elif mode == "every_n":
            every_n = int(request.form.get("every_n", 2))
        elif mode == "extract":
            raw = request.form.get("pages", "1")
            pages = [int(x.strip()) for x in raw.split(",")]
        result = engine.split(path, mode=mode, pages=pages, every_n=every_n)
        _cleanup(path)
        dl = "split_result.zip" if result.endswith(".zip") else "split_result.pdf"
        return _send_and_cleanup(result, dl)
    except Exception as e:
        _cleanup(path)
        return jsonify(error=str(e)), 500


@app.route("/api/reorder", methods=["POST"])
def reorder():
    f = request.files.get("file")
    order_str = request.form.get("order", "")
    if not f or not order_str:
        return jsonify(error="File and order required"), 400
    path = _save(f)
    output = _out(f.filename, "_reordered")
    try:
        order = [int(x.strip()) for x in order_str.split(",")]
        engine.reorder(path, order, output)
        _cleanup(path)
        return _send_and_cleanup(output, "reordered.pdf")
    except Exception as e:
        _cleanup(path)
        return jsonify(error=str(e)), 500


@app.route("/api/convert", methods=["POST"])
def convert():
    f = request.files.get("file")
    target = request.form.get("target", "docx")
    if not f:
        return jsonify(error="No file"), 400
    path = _save(f)
    ext = Path(f.filename).suffix.lower()
    try:
        if ext == ".pdf" and target == "docx":
            output = _out(f.filename, "_converted", ".docx")
            engine.pdf_to_word(path, output)
            dl = Path(f.filename).stem + ".docx"
        elif ext == ".pdf" and target == "xlsx":
            output = _out(f.filename, "_converted", ".xlsx")
            engine.pdf_to_excel(path, output)
            dl = Path(f.filename).stem + ".xlsx"
        elif ext == ".pdf" and target == "pptx":
            output = _out(f.filename, "_converted", ".pptx")
            engine.pdf_to_ppt(path, output)
            dl = Path(f.filename).stem + ".pptx"
        elif ext in (".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt"):
            output = engine.office_to_pdf(path, OUTPUT_DIR)
            dl = Path(f.filename).stem + ".pdf"
        else:
            _cleanup(path)
            return jsonify(error=f"Unsupported: {ext} to {target}"), 400
        _cleanup(path)
        return _send_and_cleanup(output, dl)
    except Exception as e:
        _cleanup(path)
        return jsonify(error=str(e)), 500


@app.route("/api/watermark", methods=["POST"])
def watermark():
    f = request.files.get("file")
    if not f:
        return jsonify(error="No file"), 400
    path = _save(f)
    output = _out(f.filename, "_watermarked")
    try:
        engine.add_watermark(
            path, output,
            text=request.form.get("text", "CONFIDENTIAL"),
            fontsize=int(request.form.get("fontsize", 60)),
            opacity=float(request.form.get("opacity", 0.3)),
            rotation=int(request.form.get("rotation", 45)),
            font_name=request.form.get("font", "Helvetica"),
        )
        _cleanup(path)
        return _send_and_cleanup(output, "watermarked.pdf")
    except Exception as e:
        _cleanup(path)
        return jsonify(error=str(e)), 500


@app.route("/api/page-numbers", methods=["POST"])
def page_numbers():
    f = request.files.get("file")
    if not f:
        return jsonify(error="No file"), 400
    path = _save(f)
    output = _out(f.filename, "_numbered")
    try:
        engine.add_page_numbers(
            path, output,
            position=request.form.get("position", "bottom-center"),
            start_num=int(request.form.get("start", 1)),
            fontsize=int(request.form.get("fontsize", 12)),
            format_str=request.form.get("format", "{n}"),
        )
        _cleanup(path)
        return _send_and_cleanup(output, "numbered.pdf")
    except Exception as e:
        _cleanup(path)
        return jsonify(error=str(e)), 500


@app.route("/api/header-footer", methods=["POST"])
def header_footer():
    f = request.files.get("file")
    if not f:
        return jsonify(error="No file"), 400
    path = _save(f)
    output = _out(f.filename, "_hf")
    try:
        engine.add_header_footer(
            path, output,
            header=request.form.get("header", ""),
            footer=request.form.get("footer", ""),
            fontsize=int(request.form.get("fontsize", 10)),
        )
        _cleanup(path)
        return _send_and_cleanup(output, "header_footer.pdf")
    except Exception as e:
        _cleanup(path)
        return jsonify(error=str(e)), 500


@app.route("/api/ocr", methods=["POST"])
def ocr():
    f = request.files.get("file")
    if not f:
        return jsonify(error="No file"), 400
    path = _save(f)
    output = _out(f.filename, "_ocr")
    try:
        engine.ocr_pdf(
            path, output,
            language=request.form.get("language", "eng"),
        )
        _cleanup(path)
        return _send_and_cleanup(output, "ocr_searchable.pdf")
    except Exception as e:
        _cleanup(path)
        return jsonify(error=str(e)), 500


@app.route("/api/redact", methods=["POST"])
def redact():
    f = request.files.get("file")
    texts = request.form.get("texts", "")
    if not f or not texts.strip():
        return jsonify(error="File and text required"), 400
    path = _save(f)
    output = _out(f.filename, "_redacted")
    try:
        search = [t.strip() for t in texts.split(",") if t.strip()]
        engine.redact_text(path, output, search)
        _cleanup(path)
        return _send_and_cleanup(output, "redacted.pdf")
    except Exception as e:
        _cleanup(path)
        return jsonify(error=str(e)), 500


@app.route("/api/repair", methods=["POST"])
def repair():
    f = request.files.get("file")
    if not f:
        return jsonify(error="No file"), 400
    path = _save(f)
    output = _out(f.filename, "_repaired")
    try:
        engine.repair_pdf(path, output)
        _cleanup(path)
        return _send_and_cleanup(output, "repaired.pdf")
    except Exception as e:
        _cleanup(path)
        return jsonify(error=str(e)), 500


@app.route("/api/compress", methods=["POST"])
def compress():
    f = request.files.get("file")
    if not f:
        return jsonify(error="No file"), 400
    path = _save(f)
    output = _out(f.filename, "_compressed")
    try:
        engine.compress(path, output,
                        quality=request.form.get("quality", "medium"))
        _cleanup(path)
        return _send_and_cleanup(output, "compressed.pdf")
    except Exception as e:
        _cleanup(path)
        return jsonify(error=str(e)), 500


@app.route("/api/rotate", methods=["POST"])
def rotate():
    f = request.files.get("file")
    if not f:
        return jsonify(error="No file"), 400
    path = _save(f)
    output = _out(f.filename, "_rotated")
    try:
        angle = int(request.form.get("angle", 90))
        pages_str = request.form.get("pages", "").strip()
        page_numbers_list = None
        if pages_str:
            page_numbers_list = [int(x.strip()) for x in pages_str.split(",")]
        engine.rotate(path, output, angle=angle,
                      page_numbers=page_numbers_list)
        _cleanup(path)
        return _send_and_cleanup(output, "rotated.pdf")
    except Exception as e:
        _cleanup(path)
        return jsonify(error=str(e)), 500


@app.route("/api/pdf-to-images", methods=["POST"])
def pdf_to_images():
    f = request.files.get("file")
    if not f:
        return jsonify(error="No file"), 400
    path = _save(f)
    try:
        fmt = request.form.get("format", "png")
        dpi = int(request.form.get("dpi", 200))
        result = engine.pdf_to_images(path, img_format=fmt, dpi=dpi)
        _cleanup(path)
        return _send_and_cleanup(result, "pdf_images.zip")
    except Exception as e:
        _cleanup(path)
        return jsonify(error=str(e)), 500


@app.route("/api/images-to-pdf", methods=["POST"])
def images_to_pdf():
    files = request.files.getlist("files")
    if not files:
        return jsonify(error="No images"), 400
    paths = [_save(f) for f in files]
    output = _out("combined_images", ext=".pdf")
    try:
        engine.images_to_pdf(paths, output)
        _cleanup(*paths)
        return _send_and_cleanup(output, "images_combined.pdf")
    except Exception as e:
        _cleanup(*paths)
        return jsonify(error=str(e)), 500


@app.route("/api/protect", methods=["POST"])
def protect():
    f = request.files.get("file")
    password = request.form.get("password", "")
    if not f or not password:
        return jsonify(error="File and password required"), 400
    path = _save(f)
    output = _out(f.filename, "_protected")
    try:
        engine.protect(path, output, password)
        _cleanup(path)
        return _send_and_cleanup(output, "protected.pdf")
    except Exception as e:
        _cleanup(path)
        return jsonify(error=str(e)), 500


@app.route("/api/unlock", methods=["POST"])
def unlock():
    f = request.files.get("file")
    password = request.form.get("password", "")
    if not f or not password:
        return jsonify(error="File and password required"), 400
    path = _save(f)
    output = _out(f.filename, "_unlocked")
    try:
        engine.unlock(path, output, password)
        _cleanup(path)
        return _send_and_cleanup(output, "unlocked.pdf")
    except Exception as e:
        _cleanup(path)
        return jsonify(error=str(e)), 500


@app.route("/api/bates", methods=["POST"])
def bates():
    f = request.files.get("file")
    if not f:
        return jsonify(error="No file"), 400
    path = _save(f)
    output = _out(f.filename, "_bates")
    try:
        engine.bates_number(
            path, output,
            prefix=request.form.get("prefix", "DOC"),
            start=int(request.form.get("start", 1)),
            digits=int(request.form.get("digits", 6)),
            position=request.form.get("position", "bottom-right"),
        )
        _cleanup(path)
        return _send_and_cleanup(output, "bates_numbered.pdf")
    except Exception as e:
        _cleanup(path)
        return jsonify(error=str(e)), 500


# ── EDITOR ENDPOINTS ──
@app.route("/api/editor/upload", methods=["POST"])
def editor_upload():
    f = request.files.get("file")
    if not f:
        return jsonify(error="No file"), 400
    path = _save(f)
    session_id = uuid.uuid4().hex
    edit_sessions[session_id] = path
    page_count = engine.get_page_count(path)
    return jsonify(session_id=session_id, pages=page_count)


@app.route("/api/editor/page/<session_id>/<int:page_num>")
def editor_page(session_id, page_num):
    path = edit_sessions.get(session_id)
    if not path:
        return jsonify(error="Session expired"), 404
    try:
        img = engine.get_preview(path, page_num=page_num, dpi=150)
        return send_file(io.BytesIO(img), mimetype="image/png")
    except Exception as e:
        return jsonify(error=str(e)), 500


@app.route("/api/editor/save", methods=["POST"])
def editor_save():
    data = request.get_json()
    if not data:
        return jsonify(error="No data"), 400
    session_id = data.get("session_id")
    annotations = data.get("annotations", [])
    path = edit_sessions.get(session_id)
    if not path:
        return jsonify(error="Session expired"), 404
    output = _out("edited", ext=".pdf")
    try:
        engine.apply_edits(path, output, annotations, dpi=150)
        _cleanup(path)
        if session_id in edit_sessions:
            del edit_sessions[session_id]
        return _send_and_cleanup(output, "edited.pdf")
    except Exception as e:
        return jsonify(error=str(e)), 500


# ───────── Launch ─────────
def main():
    port = int(os.environ.get("PORT", 5000))
    is_cloud = "RENDER" in os.environ or "PORT" in os.environ

    if is_cloud:
        app.run(host="0.0.0.0", port=port, debug=False)
    else:
        print(f"\n{'=' * 50}")
        print("  PDF Toolkit Pro")
        print(f"  Open -> http://127.0.0.1:{port}")
        if APP_PASSWORD:
            print(f"  Password: {APP_PASSWORD}")
        else:
            print("  No password (open access)")
        print(f"{'=' * 50}\n")
        webbrowser.open(f"http://127.0.0.1:{port}")
        app.run(host="0.0.0.0", port=port, debug=True)


if __name__ == "__main__":
    main()