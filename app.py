"""
PDF Toolkit Pro - Complete Server
"""
import os, io, sys, uuid, json, shutil, tempfile, threading, webbrowser
from pathlib import Path
from flask import (Flask, request, jsonify, send_file,
                   render_template, session, redirect)
from werkzeug.utils import secure_filename
from pdf_engine import PDFEngine

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "pdf-toolkit-pro-local-2025")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "pdftoolkit_up")
OUTPUT_DIR = os.path.join(tempfile.gettempdir(), "pdftoolkit_out")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
app.config["MAX_CONTENT_LENGTH"] = 150 * 1024 * 1024

engine = PDFEngine()
edit_sessions = {}
delete_sessions = {}

@app.before_request
def require_login():
    open_paths = ["/login", "/static/", "/manifest.json"]
    if any(request.path.startswith(p) for p in open_paths): return
    if APP_PASSWORD and not session.get("authenticated"):
        if request.path.startswith("/api/"): return jsonify(error="Not authenticated"), 401
        return redirect("/login")

@app.route("/login", methods=["GET", "POST"])
def login():
    if not APP_PASSWORD: return redirect("/")
    err = ""
    if request.method == "POST":
        if request.form.get("password", "") == APP_PASSWORD:
            session["authenticated"] = True; session.permanent = True; return redirect("/")
        err = '<div style="background:#ffe0e0;color:#e74c3c;padding:10px;border-radius:8px;margin-bottom:16px">Wrong password</div>'
    return f'''<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>PDF Toolkit Pro</title>
<style>*{{margin:0;padding:0;box-sizing:border-box}}body{{font-family:'Segoe UI',system-ui,sans-serif;background:#f0f2f5;display:flex;align-items:center;justify-content:center;min-height:100vh}}.box{{background:#fff;border-radius:20px;padding:40px;max-width:400px;width:90%;box-shadow:0 4px 20px rgba(0,0,0,.1);text-align:center}}.logo{{font-size:1.6rem;font-weight:900;background:linear-gradient(135deg,#667eea,#764ba2);-webkit-background-clip:text;-webkit-text-fill-color:transparent;margin-bottom:16px}}input{{width:100%;padding:14px;border:2px solid #ddd;border-radius:10px;font-size:1rem;margin-bottom:16px;text-align:center}}input:focus{{border-color:#667eea;outline:none}}button{{width:100%;padding:14px;background:linear-gradient(135deg,#667eea,#764ba2);color:#fff;border:none;border-radius:10px;font-size:1rem;font-weight:700;cursor:pointer}}</style></head><body>
<div class="box"><div class="logo">PDF Toolkit Pro</div>{err}<form method="POST"><input type="password" name="password" placeholder="Enter password" autofocus required><button type="submit">Unlock</button></form></div></body></html>'''

def _save(file):
    name = secure_filename(file.filename) or "upload"
    path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4().hex}_{name}")
    file.save(path); return path

def _out(name, suffix="", ext=None):
    stem = Path(name).stem
    if ext is None: ext = Path(name).suffix
    return os.path.join(OUTPUT_DIR, f"{uuid.uuid4().hex}_{stem}{suffix}{ext}")

def _cleanup(*paths, delay=5):
    def _do():
        import time; time.sleep(delay)
        for p in paths:
            try:
                if os.path.isdir(p): shutil.rmtree(p, ignore_errors=True)
                elif os.path.isfile(p): os.remove(p)
            except OSError: pass
    threading.Thread(target=_do, daemon=True).start()

def _send_and_cleanup(fp, dl, mt=None):
    resp = send_file(fp, as_attachment=True, download_name=dl, mimetype=mt)
    _cleanup(fp, delay=15); return resp

@app.route("/")
def index(): return render_template("index.html")

@app.route("/manifest.json")
def manifest(): return send_file("static/manifest.json", mimetype="application/manifest+json")

@app.route("/api/preview", methods=["POST"])
def preview():
    f = request.files.get("file"); page = int(request.form.get("page", 0))
    if not f: return jsonify(error="No file"), 400
    path = _save(f)
    try:
        img = engine.get_preview(path, page_num=page)
        ip = _out(f.filename, f"_p{page}", ".png")
        with open(ip, "wb") as fp: fp.write(img)
        _cleanup(path); _cleanup(ip, delay=60)
        return send_file(ip, mimetype="image/png")
    except Exception as e: _cleanup(path); return jsonify(error=str(e)), 500

@app.route("/api/metadata", methods=["POST"])
def metadata():
    f = request.files.get("file")
    if not f: return jsonify(error="No file"), 400
    path = _save(f)
    try: meta = engine.get_metadata(path); _cleanup(path); return jsonify(meta)
    except Exception as e: _cleanup(path); return jsonify(error=str(e)), 500

@app.route("/api/merge", methods=["POST"])
def merge():
    files = request.files.getlist("files")
    if len(files) < 2: return jsonify(error="Need 2+ PDFs"), 400
    paths = [_save(f) for f in files]; output = _out("merged", ext=".pdf")
    try: engine.merge(paths, output); _cleanup(*paths); return _send_and_cleanup(output, "merged.pdf")
    except Exception as e: _cleanup(*paths); return jsonify(error=str(e)), 500

@app.route("/api/split", methods=["POST"])
def split():
    f = request.files.get("file"); mode = request.form.get("mode", "all")
    if not f: return jsonify(error="No file"), 400
    path = _save(f)
    try:
        pages = None; every_n = 1
        if mode == "ranges":
            raw = request.form.get("ranges", "1-1"); pages = []
            for r in raw.split(","): parts = r.strip().split("-"); pages.append((int(parts[0]), int(parts[1])))
        elif mode == "every_n": every_n = int(request.form.get("every_n", 2))
        elif mode == "extract": pages = [int(x.strip()) for x in request.form.get("pages", "1").split(",")]
        result = engine.split(path, mode=mode, pages=pages, every_n=every_n); _cleanup(path)
        return _send_and_cleanup(result, "split.zip" if result.endswith(".zip") else "split.pdf")
    except Exception as e: _cleanup(path); return jsonify(error=str(e)), 500

@app.route("/api/reorder", methods=["POST"])
def reorder():
    f = request.files.get("file"); order_str = request.form.get("order", "")
    if not f or not order_str: return jsonify(error="File and order required"), 400
    path = _save(f); output = _out(f.filename, "_reordered")
    try: engine.reorder(path, [int(x.strip()) for x in order_str.split(",")], output); _cleanup(path); return _send_and_cleanup(output, "reordered.pdf")
    except Exception as e: _cleanup(path); return jsonify(error=str(e)), 500

@app.route("/api/convert", methods=["POST"])
def convert():
    f = request.files.get("file"); target = request.form.get("target", "docx")
    if not f: return jsonify(error="No file"), 400
    path = _save(f); ext = Path(f.filename).suffix.lower()
    try:
        if ext == ".pdf" and target == "docx": output = _out(f.filename, "_converted", ".docx"); engine.pdf_to_word(path, output); dl = Path(f.filename).stem+".docx"
        elif ext == ".pdf" and target == "xlsx": output = _out(f.filename, "_converted", ".xlsx"); engine.pdf_to_excel(path, output); dl = Path(f.filename).stem+".xlsx"
        elif ext == ".pdf" and target == "pptx": output = _out(f.filename, "_converted", ".pptx"); engine.pdf_to_ppt(path, output); dl = Path(f.filename).stem+".pptx"
        elif ext in (".docx",".doc",".xlsx",".xls",".pptx",".ppt"): output = engine.office_to_pdf(path, OUTPUT_DIR); dl = Path(f.filename).stem+".pdf"
        else: _cleanup(path); return jsonify(error=f"Unsupported: {ext}"), 400
        _cleanup(path); return _send_and_cleanup(output, dl)
    except Exception as e: _cleanup(path); return jsonify(error=str(e)), 500

@app.route("/api/watermark", methods=["POST"])
def watermark():
    f = request.files.get("file")
    if not f: return jsonify(error="No file"), 400
    path = _save(f); output = _out(f.filename, "_wm")
    try:
        engine.add_watermark(path, output, text=request.form.get("text", "CONFIDENTIAL"), fontsize=int(request.form.get("fontsize", 60)), opacity=float(request.form.get("opacity", 0.3)), rotation=int(request.form.get("rotation", 45)), font_name=request.form.get("font", "Helvetica"))
        _cleanup(path); return _send_and_cleanup(output, "watermarked.pdf")
    except Exception as e: _cleanup(path); return jsonify(error=str(e)), 500

@app.route("/api/page-numbers", methods=["POST"])
def page_numbers():
    f = request.files.get("file")
    if not f: return jsonify(error="No file"), 400
    path = _save(f); output = _out(f.filename, "_num")
    try:
        engine.add_page_numbers(path, output, position=request.form.get("position", "bottom-center"), start_num=int(request.form.get("start", 1)), fontsize=int(request.form.get("fontsize", 12)), format_str=request.form.get("format", "{n}"))
        _cleanup(path); return _send_and_cleanup(output, "numbered.pdf")
    except Exception as e: _cleanup(path); return jsonify(error=str(e)), 500

@app.route("/api/header-footer", methods=["POST"])
def header_footer():
    f = request.files.get("file")
    if not f: return jsonify(error="No file"), 400
    path = _save(f); output = _out(f.filename, "_hf")
    try:
        engine.add_header_footer(path, output, header=request.form.get("header", ""), footer=request.form.get("footer", ""), fontsize=int(request.form.get("fontsize", 10)))
        _cleanup(path); return _send_and_cleanup(output, "header_footer.pdf")
    except Exception as e: _cleanup(path); return jsonify(error=str(e)), 500

@app.route("/api/ocr", methods=["POST"])
def ocr():
    f = request.files.get("file")
    if not f: return jsonify(error="No file"), 400
    path = _save(f); output = _out(f.filename, "_ocr")
    try: engine.ocr_pdf(path, output, language=request.form.get("language", "eng")); _cleanup(path); return _send_and_cleanup(output, "ocr_searchable.pdf")
    except Exception as e: _cleanup(path); return jsonify(error=str(e)), 500

@app.route("/api/redact", methods=["POST"])
def redact():
    f = request.files.get("file"); texts = request.form.get("texts", "")
    if not f or not texts.strip(): return jsonify(error="File and text required"), 400
    path = _save(f); output = _out(f.filename, "_redacted")
    try: engine.redact_text(path, output, [t.strip() for t in texts.split(",") if t.strip()]); _cleanup(path); return _send_and_cleanup(output, "redacted.pdf")
    except Exception as e: _cleanup(path); return jsonify(error=str(e)), 500

@app.route("/api/repair", methods=["POST"])
def repair():
    f = request.files.get("file")
    if not f: return jsonify(error="No file"), 400
    path = _save(f); output = _out(f.filename, "_repaired")
    try: engine.repair_pdf(path, output); _cleanup(path); return _send_and_cleanup(output, "repaired.pdf")
    except Exception as e: _cleanup(path); return jsonify(error=str(e)), 500

@app.route("/api/compress", methods=["POST"])
def compress():
    f = request.files.get("file")
    if not f: return jsonify(error="No file"), 400
    path = _save(f); output = _out(f.filename, "_compressed")
    try: engine.compress(path, output, quality=request.form.get("quality", "medium")); _cleanup(path); return _send_and_cleanup(output, "compressed.pdf")
    except Exception as e: _cleanup(path); return jsonify(error=str(e)), 500

@app.route("/api/rotate", methods=["POST"])
def rotate():
    f = request.files.get("file")
    if not f: return jsonify(error="No file"), 400
    path = _save(f); output = _out(f.filename, "_rotated")
    try:
        ps = request.form.get("pages", "").strip()
        engine.rotate(path, output, angle=int(request.form.get("angle", 90)), page_numbers=[int(x.strip()) for x in ps.split(",")] if ps else None)
        _cleanup(path); return _send_and_cleanup(output, "rotated.pdf")
    except Exception as e: _cleanup(path); return jsonify(error=str(e)), 500

@app.route("/api/pdf-to-images", methods=["POST"])
def pdf_to_images():
    f = request.files.get("file")
    if not f: return jsonify(error="No file"), 400
    path = _save(f)
    try: result = engine.pdf_to_images(path, img_format=request.form.get("format", "png"), dpi=int(request.form.get("dpi", 200))); _cleanup(path); return _send_and_cleanup(result, "pdf_images.zip")
    except Exception as e: _cleanup(path); return jsonify(error=str(e)), 500

@app.route("/api/images-to-pdf", methods=["POST"])
def images_to_pdf():
    files = request.files.getlist("files")
    if not files: return jsonify(error="No images"), 400
    paths = [_save(f) for f in files]; output = _out("images", ext=".pdf")
    try: engine.images_to_pdf(paths, output); _cleanup(*paths); return _send_and_cleanup(output, "images.pdf")
    except Exception as e: _cleanup(*paths); return jsonify(error=str(e)), 500

@app.route("/api/protect", methods=["POST"])
def protect():
    f = request.files.get("file"); pw = request.form.get("password", "")
    if not f or not pw: return jsonify(error="File and password required"), 400
    path = _save(f); output = _out(f.filename, "_protected")
    try: engine.protect(path, output, pw); _cleanup(path); return _send_and_cleanup(output, "protected.pdf")
    except Exception as e: _cleanup(path); return jsonify(error=str(e)), 500

@app.route("/api/unlock", methods=["POST"])
def unlock():
    f = request.files.get("file"); pw = request.form.get("password", "")
    if not f or not pw: return jsonify(error="File and password required"), 400
    path = _save(f); output = _out(f.filename, "_unlocked")
    try: engine.unlock(path, output, pw); _cleanup(path); return _send_and_cleanup(output, "unlocked.pdf")
    except Exception as e: _cleanup(path); return jsonify(error=str(e)), 500

@app.route("/api/bates", methods=["POST"])
def bates():
    f = request.files.get("file")
    if not f: return jsonify(error="No file"), 400
    path = _save(f); output = _out(f.filename, "_bates")
    try:
        engine.bates_number(path, output, prefix=request.form.get("prefix", "DOC"), start=int(request.form.get("start", 1)), digits=int(request.form.get("digits", 6)), position=request.form.get("position", "bottom-right"))
        _cleanup(path); return _send_and_cleanup(output, "bates.pdf")
    except Exception as e: _cleanup(path); return jsonify(error=str(e)), 500

@app.route("/api/editor/upload", methods=["POST"])
def editor_upload():
    f = request.files.get("file")
    if not f: return jsonify(error="No file"), 400
    path = _save(f); sid = uuid.uuid4().hex; edit_sessions[sid] = path
    return jsonify(session_id=sid, pages=engine.get_page_count(path))

@app.route("/api/editor/page/<sid>/<int:pn>")
def editor_page(sid, pn):
    path = edit_sessions.get(sid)
    if not path: return jsonify(error="Session expired"), 404
    try: return send_file(io.BytesIO(engine.get_preview(path, page_num=pn, dpi=150)), mimetype="image/png")
    except Exception as e: return jsonify(error=str(e)), 500

@app.route("/api/editor/save", methods=["POST"])
def editor_save():
    data = request.get_json()
    if not data: return jsonify(error="No data"), 400
    sid = data.get("session_id"); path = edit_sessions.get(sid)
    if not path: return jsonify(error="Session expired"), 404
    output = _out("edited", ext=".pdf")
    try:
        engine.apply_edits(path, output, data.get("annotations", []), dpi=150)
        _cleanup(path); edit_sessions.pop(sid, None)
        return _send_and_cleanup(output, "edited.pdf")
    except Exception as e: return jsonify(error=str(e)), 500

@app.route("/api/delete-pages/upload", methods=["POST"])
def del_upload():
    f = request.files.get("file")
    if not f: return jsonify(error="No file"), 400
    path = _save(f); sid = uuid.uuid4().hex; delete_sessions[sid] = path
    return jsonify(session_id=sid, pages=engine.get_page_count(path))

@app.route("/api/delete-pages/thumb/<sid>/<int:pn>")
def del_thumb(sid, pn):
    path = delete_sessions.get(sid)
    if not path: return jsonify(error="Session expired"), 404
    try: return send_file(io.BytesIO(engine.get_preview(path, page_num=pn, dpi=100)), mimetype="image/png")
    except Exception as e: return jsonify(error=str(e)), 500

@app.route("/api/delete-pages/process", methods=["POST"])
def del_process():
    data = request.get_json()
    if not data: return jsonify(error="No data"), 400
    sid = data.get("session_id"); pages = data.get("pages_to_delete", [])
    path = delete_sessions.get(sid)
    if not path: return jsonify(error="Session expired"), 404
    if not pages: return jsonify(error="No pages selected"), 400
    output = _out("pages_deleted", ext=".pdf")
    try:
        engine.delete_pages(path, output, pages); _cleanup(path); delete_sessions.pop(sid, None)
        return _send_and_cleanup(output, "pages_deleted.pdf")
    except Exception as e: return jsonify(error=str(e)), 500

def main():
    port = int(os.environ.get("PORT", 5000))
    is_cloud = "RENDER" in os.environ or "RENDER_EXTERNAL_URL" in os.environ
    if is_cloud: app.run(host="0.0.0.0", port=port, debug=False)
    else:
        print(f"\n{'='*50}")
        print("  PDF Toolkit Pro")
        print(f"  Open -> http://127.0.0.1:{port}")
        print(f"{'='*50}\n")
        webbrowser.open(f"http://127.0.0.1:{port}")
        app.run(host="0.0.0.0", port=port, debug=True)

if __name__ == "__main__": main()