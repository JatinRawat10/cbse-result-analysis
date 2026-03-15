from flask import Flask, render_template, request, send_file, redirect, url_for
from analysis import process_result
import os
import uuid
import threading
import time
import tempfile

app = Flask(__name__)

# =========================
# Config
# =========================
upload_tracker = {}
UPLOAD_LIMIT = 10    # max uploads
UPLOAD_WINDOW = 60     # seconds

upload_tracker_lock = threading.Lock()

app.config["MAX_CONTENT_LENGTH"] = 2 * 1024 * 1024

EXPIRY_SECONDS = 300
MAX_SESSIONS = 1
CLEANUP_INTERVAL = 10
PROCESSING_LIMIT = 6

processing_limit = threading.Semaphore(PROCESSING_LIMIT)
storage_lock = threading.Lock()

tmp_base = tempfile.gettempdir()
UPLOAD_DIR = os.path.join(tmp_base, "cbse_uploads")
OUTPUT_DIR = os.path.join(tmp_base, "cbse_outputs")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# {
#   file_id: {
#       "txt_path": str,
#       "excel_path": str,
#       "word_path": str,
#       "subject_inputs": dict|None,
#       "teacher_inputs": dict|None,
#       "analytics": dict|None,
#       "created_at": float,
#       "last_accessed": float,
#       "download_started": float|None,
#   }
# }
temporary_storage = {}


# =========================
# Error handler
# =========================

@app.errorhandler(413)
def file_too_large(e):
    return """
    <script>
    alert("File must be smaller than 2 MB.");
    window.history.back();
    </script>
    """, 413


# =========================
# Helpers
# =========================

def check_rate_limit(ip):
    now = time.time()

    with upload_tracker_lock:

        uploads = upload_tracker.get(ip, [])

        # remove old timestamps
        uploads = [t for t in uploads if now - t < UPLOAD_WINDOW]

        if uploads:
            upload_tracker[ip] = uploads
        else:
            upload_tracker.pop(ip, None)

        if len(uploads) >= UPLOAD_LIMIT:
            retry_after = int(UPLOAD_WINDOW - (now - uploads[0]))
            return False, retry_after

        uploads.append(now)
        upload_tracker[ip] = uploads

    return True, 0

def save_output_file(obj, path):
    if obj is None:
        return False

    data = None

    if isinstance(obj, (bytes, bytearray)):
        data = obj
    elif hasattr(obj, "getvalue"):
        data = obj.getvalue()
    elif hasattr(obj, "read"):
        try:
            pos = obj.tell() if hasattr(obj, "tell") else None
            data = obj.read()
            if pos is not None and hasattr(obj, "seek"):
                obj.seek(pos)
        except (OSError, ValueError):
            data = None

    if data is None:
        return False

    with open(path, "wb") as f:
        f.write(data)

    return True


def delete_session_files(session):
    for key in ("txt_path", "excel_path", "word_path"):
        path = session.get(key)
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except OSError:
                pass


def _session_last_seen(session):
    return (
        session.get("last_accessed")
        or session.get("created_at")
        or session.get("download_started")
        or 0
    )


def _cleanup_expired_locked(now):

    expired_ids = []

    for file_id, session in temporary_storage.items():

        started = session.get("download_started")
        created = session.get("created_at")

        # If download page was opened
        if started:
            if now - started > EXPIRY_SECONDS:
                expired_ids.append(file_id)

        # If user uploaded but never opened download page
        elif created:
            if now - created > EXPIRY_SECONDS:
                expired_ids.append(file_id)

    removed = []

    for file_id in expired_ids:
        popped = temporary_storage.pop(file_id, None)
        if popped:
            removed.append(popped)

    return removed


def _evict_for_capacity_locked():
    if len(temporary_storage) < MAX_SESSIONS:
        return []

    # Keep one slot for the incoming session.
    evict_count = len(temporary_storage) - MAX_SESSIONS + 1
    if evict_count <= 0:
        return []

    oldest_ids = sorted(
        temporary_storage,
        key=lambda file_id: _session_last_seen(temporary_storage[file_id])
    )

    removed = []
    for file_id in oldest_ids[:evict_count]:
        popped = temporary_storage.pop(file_id, None)
        if popped:
            removed.append(popped)

    return removed


def cleanup_expired():
    now = time.time()

    with storage_lock:
        removed = _cleanup_expired_locked(now)

    for session in removed:
        delete_session_files(session)


def background_cleanup():
    while True:
        time.sleep(CLEANUP_INTERVAL)
        cleanup_expired()


def get_entry(file_id):
    if not file_id:
        return None

    with storage_lock:
        entry = temporary_storage.get(file_id)
        if entry is not None:
            entry["last_accessed"] = time.time()

    return entry


def touch_session(file_id):
    with storage_lock:
        entry = temporary_storage.get(file_id)
        if entry is not None:
            now = time.time()
            entry["last_accessed"] = now
            if entry.get("download_started") is None:
                entry["download_started"] = now


def remove_session(file_id):
    with storage_lock:
        session = temporary_storage.pop(file_id, None)

    if session:
        delete_session_files(session)


def render_expired():
    return render_template("expired_session.html"), 410


def default_analytics():
    return {
        "school_avg": 0,
        "highest_percent": 0,
        "all_A1": 0,
        "top5": {}
    }


# =========================
# Routes
# =========================

@app.route("/")
def home():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    ip = request.remote_addr
    allowed, wait_time = check_rate_limit(ip)

    if not allowed:
        return f"""
        <script>
        alert("Too many uploads. Please wait {wait_time} seconds before uploading again.");
        window.history.back();
        </script>
        """, 429
    
    file = request.files.get("file")

    if not file:
        return "No file uploaded.", 400

    if not file.filename.lower().endswith(".txt"):
        return "Only .txt files are allowed.", 400

    file_id = str(uuid.uuid4())

    txt_path = os.path.join(UPLOAD_DIR, f"{file_id}.txt")
    excel_path = os.path.join(OUTPUT_DIR, f"{file_id}.xlsx")
    word_path = os.path.join(OUTPUT_DIR, f"{file_id}.docx")

    file.save(txt_path)

    now = time.time()
    with storage_lock:
        removed = _cleanup_expired_locked(now)
        removed.extend(_evict_for_capacity_locked())

        temporary_storage[file_id] = {
            "txt_path": txt_path,
            "excel_path": excel_path,
            "word_path": word_path,
            "subject_inputs": None,
            "teacher_inputs": None,
            "analytics": None,
            "created_at": now,
            "last_accessed": now,
            "download_started": None,
        }

    for session in removed:
        delete_session_files(session)

    try:
        with processing_limit:
            with open(txt_path, "rb") as f:
                result = process_result(f)
    except Exception as e:
        remove_session(file_id)
        return f"Processing error: {e}", 500

    if not isinstance(result, dict):
        remove_session(file_id)
        return "Processing error: unexpected result format", 500

    if "missing_subjects" in result:
        return render_template(
            "missing_subjects.html",
            codes=result["missing_subjects"],
            file_id=file_id
        )

    if "missing_teachers" in result:
        with storage_lock:
            entry = temporary_storage.get(file_id)
            if entry:
                entry["analytics"] = result.get("analytics")

        return render_template(
            "missing_teachers.html",
            subjects=result["missing_teachers"],
            file_id=file_id
        )

    with storage_lock:
        entry = temporary_storage.get(file_id)

    if not entry:
        return render_expired()

    excel_ok = save_output_file(result.get("excel_file"), entry["excel_path"])
    word_ok = save_output_file(result.get("word_file"), entry["word_path"])
    
    if not excel_ok or not word_ok:
        remove_session(file_id)
        return "Processing error: failed to save output files", 500

    with storage_lock:
        entry = temporary_storage.get(file_id)
        if not entry:
            return render_expired()
        entry["analytics"] = result.get("analytics")

    return redirect(url_for("download_page", file_id=file_id))


@app.route("/submit_subjects/<file_id>", methods=["POST"])
def submit_subjects(file_id):
    entry = get_entry(file_id)

    if not entry:
        return render_expired()

    subject_inputs = dict(request.form)

    with storage_lock:
        entry["subject_inputs"] = subject_inputs

    try:
        with processing_limit:
            with open(entry["txt_path"], "rb") as f:
                result = process_result(
                    f,
                    subject_inputs=subject_inputs,
                    teacher_inputs=entry.get("teacher_inputs")
                )
    except Exception as e:
        return f"Processing error: {e}", 500

    if not isinstance(result, dict):
        return "Processing error: unexpected result format", 500

    if "missing_subjects" in result:
        return render_template(
            "missing_subjects.html",
            codes=result["missing_subjects"],
            file_id=file_id
        )

    if "missing_teachers" in result:
        with storage_lock:
            entry = temporary_storage.get(file_id)
            if entry:
                entry["analytics"] = result.get("analytics")

        return render_template(
            "missing_teachers.html",
            subjects=result["missing_teachers"],
            file_id=file_id
        )

    with storage_lock:
        entry = temporary_storage.get(file_id)

    if not entry:
        return render_expired()

    excel_ok = save_output_file(result.get("excel_file"), entry["excel_path"])
    word_ok = save_output_file(result.get("word_file"), entry["word_path"])
    
    if not excel_ok or not word_ok:
        remove_session(file_id)
        return "Processing error: failed to save output files", 500

    with storage_lock:
        entry = temporary_storage.get(file_id)
        if not entry:
            return render_expired()
        entry["analytics"] = result.get("analytics")

    return redirect(url_for("download_page", file_id=file_id))


@app.route("/submit_teachers/<file_id>", methods=["POST"])
def submit_teachers(file_id):
    entry = get_entry(file_id)

    if not entry:
        return render_expired()

    teacher_inputs = dict(request.form)

    with storage_lock:
        entry["teacher_inputs"] = teacher_inputs

    try:
        with processing_limit:
            with open(entry["txt_path"], "rb") as f:
                result = process_result(
                    f,
                    subject_inputs=entry.get("subject_inputs"),
                    teacher_inputs=teacher_inputs
                )
    except Exception as e:
        return f"Processing error: {e}", 500

    if not isinstance(result, dict):
        return "Processing error: unexpected result format", 500

    if "missing_teachers" in result:
        return render_template(
            "missing_teachers.html",
            subjects=result["missing_teachers"],
            file_id=file_id
        )

    with storage_lock:
        entry = temporary_storage.get(file_id)

    if not entry:
        return render_expired()

    excel_ok = save_output_file(result.get("excel_file"), entry["excel_path"])
    word_ok = save_output_file(result.get("word_file"), entry["word_path"])
    
    if not excel_ok or not word_ok:
        remove_session(file_id)
        return "Processing error: failed to save output files", 500

    with storage_lock:
        entry = temporary_storage.get(file_id)
        if not entry:
            return render_expired()
        entry["analytics"] = result.get("analytics")

    return redirect(url_for("download_page", file_id=file_id))


@app.route("/download/<file_id>")
def download_page(file_id):
    entry = get_entry(file_id)

    if not entry:
        return render_expired()

    with storage_lock:
        if entry["download_started"] is None:
            now = time.time()
            entry["download_started"] = now
            entry["last_accessed"] = now

        remaining = EXPIRY_SECONDS - (time.time() - entry["download_started"])

    return render_template(
        "download.html",
        file_id=file_id,
        analytics=entry.get("analytics") or default_analytics(),
        remaining=int(max(0, remaining))
    )


@app.route("/download_excel/<file_id>")
def download_excel(file_id):
    entry = get_entry(file_id)

    if not entry:
        return render_expired()

    touch_session(file_id)

    path = entry.get("excel_path")

    if not path or not os.path.exists(path):
        return render_expired()

    return send_file(
        path,
        as_attachment=True,
        download_name="CBSE_Result.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/download_word/<file_id>")
def download_word(file_id):
    entry = get_entry(file_id)

    if not entry:
        return render_expired()

    touch_session(file_id)

    path = entry.get("word_path")

    if not path or not os.path.exists(path):
        return render_expired()

    return send_file(
        path,
        as_attachment=True,
        download_name="CBSE_Forms.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/report/<file_id>")
def report_page(file_id):
    entry = get_entry(file_id)

    if not entry:
        return render_expired()

    touch_session(file_id)

    return render_template(
        "report.html",
        file_id=file_id,
        analytics=entry.get("analytics") or default_analytics()
    )


@app.route("/result/<file_id>")
def result_page(file_id):
    return redirect(url_for("download_page", file_id=file_id))


# =========================
# Start cleanup thread
# =========================

cleanup_thread = threading.Thread(target=background_cleanup, daemon=True)
cleanup_thread.start()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
