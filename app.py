from flask import Flask, render_template, request, send_file, redirect, url_for
from analysis import process_result
import os
import uuid
import threading
import time
import tempfile

app = Flask(__name__)

# Upload limit: 2 MB
app.config["MAX_CONTENT_LENGTH"] = 2 * 1024 * 1024

@app.errorhandler(413)
def file_too_large(e):
    return """
    <script>
    alert("File must be smaller than 2 MB.");
    window.history.back();
    </script>
    """, 413


# =========================
# Config
# =========================

EXPIRY_SECONDS = 300  # 5 minutes
MAX_SESSIONS = 100
CLEANUP_INTERVAL = 10  # seconds
PROCESSING_LIMIT = 2   # max simultaneous processing jobs

processing_limit = threading.Semaphore(PROCESSING_LIMIT)
storage_lock = threading.Lock()

BASE_TMP_DIR = tempfile.gettempdir()
UPLOAD_DIR = os.path.join(BASE_TMP_DIR, "cbse_uploads")
OUTPUT_DIR = os.path.join(BASE_TMP_DIR, "cbse_outputs")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


# =========================
# In-memory session store
# =========================
# {
#   file_id: {
#       "txt_path": ...,
#       "excel_path": ...,
#       "word_path": ...,
#       "subject_inputs": {...} or None,
#       "teacher_inputs": {...} or None,
#       "analytics": {...} or None,
#       "created_at": timestamp,
#       "download_started": timestamp or None
#   }
# }
temporary_storage = {}


# =========================
# Helper Functions
# =========================

def save_output_file(output_obj, out_path):
    """
    Save excel/word result to disk.
    Supports bytes, bytearray, and file-like objects with getvalue()/read().
    """
    if output_obj is None:
        return False

    data = None
    if isinstance(output_obj, (bytes, bytearray)):
        data = output_obj
    elif hasattr(output_obj, "getvalue"):
        data = output_obj.getvalue()
    elif hasattr(output_obj, "read"):
        try:
            pos = output_obj.tell() if hasattr(output_obj, "tell") else None
            data = output_obj.read()
            if pos is not None and hasattr(output_obj, "seek"):
                output_obj.seek(pos)
        except Exception:
            data = None

    if data is None:
        return False

    with open(out_path, "wb") as f:
        f.write(data)

    return True


def delete_session_files(session):
    """
    Delete all files linked to a session.
    """
    for path_key in ("txt_path", "excel_path", "word_path"):
        path = session.get(path_key)
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except Exception:
                pass


def cleanup_expired():
    """
    Delete:
    1) abandoned uploads that never opened the download page
    2) active download sessions after expiry
    Also removes their physical files.
    """
    now = time.time()
    expired_ids = []

    with storage_lock:
        for file_id, data in temporary_storage.items():
            created_at = data.get("created_at")
            started_at = data.get("download_started")

            # Abandoned upload: never reached download page
            if started_at is None:
                if created_at and (now - created_at > EXPIRY_SECONDS):
                    expired_ids.append(file_id)

            # Normal session after download page opened
            else:
                if now - started_at > EXPIRY_SECONDS:
                    expired_ids.append(file_id)

        removed_sessions = []
        for file_id in expired_ids:
            session = temporary_storage.pop(file_id, None)
            if session:
                removed_sessions.append(session)

    # Delete files outside the lock
    for session in removed_sessions:
        delete_session_files(session)


def background_cleanup():
    while True:
        time.sleep(CLEANUP_INTERVAL)
        cleanup_expired()


def get_entry(file_id):
    if not file_id:
        return None

    cleanup_expired()

    with storage_lock:
        return temporary_storage.get(file_id)


def render_expired():
    return render_template("expired_session.html"), 410


# =========================
# Routes
# =========================

@app.route("/")
def home():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    file = request.files.get("file")

    if not file:
        return "No file uploaded.", 400

    if not file.filename.lower().endswith(".txt"):
        return "Only .txt files are allowed.", 400

    file_id = str(uuid.uuid4())

    txt_path = os.path.join(UPLOAD_DIR, f"{file_id}.txt")
    excel_path = os.path.join(OUTPUT_DIR, f"{file_id}.xlsx")
    word_path = os.path.join(OUTPUT_DIR, f"{file_id}.docx")

    # Save uploaded file to disk, not RAM
    file.save(txt_path)

    with storage_lock:
        if len(temporary_storage) >= MAX_SESSIONS:
            cleanup_expired()

        temporary_storage[file_id] = {
            "txt_path": txt_path,
            "excel_path": excel_path,
            "word_path": word_path,
            "subject_inputs": None,
            "teacher_inputs": None,
            "analytics": None,
            "created_at": time.time(),
            "download_started": None
        }

    try:
        with processing_limit:
            with open(txt_path, "rb") as f:
                result = process_result(f)

    except Exception as e:
        with storage_lock:
            session = temporary_storage.pop(file_id, None)
        if session:
            delete_session_files(session)
        return f"Processing error: {e}", 500

    # Missing subject mapping
    if isinstance(result, dict) and "missing_subjects" in result:
        return render_template(
            "missing_subjects.html",
            codes=result["missing_subjects"],
            file_id=file_id
        )

    # Missing teacher mapping
    if isinstance(result, dict) and "missing_teachers" in result:
        with storage_lock:
            entry = temporary_storage.get(file_id)
            if entry:
                entry["analytics"] = result.get("analytics")

        return render_template(
            "missing_teachers.html",
            subjects=result["missing_teachers"],
            file_id=file_id
        )

    # Save generated files to disk
    with storage_lock:
        entry = temporary_storage.get(file_id)

    if not entry:
        return render_expired()

    excel_file = result.get("excel_file") if isinstance(result, dict) else None
    word_file = result.get("word_file") if isinstance(result, dict) else None

    if excel_file:
        save_output_file(excel_file, entry["excel_path"])

    if word_file:
        save_output_file(word_file, entry["word_path"])

    with storage_lock:
        entry = temporary_storage.get(file_id)
        if not entry:
            return render_expired()
        entry["analytics"] = result.get("analytics") if isinstance(result, dict) else None

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

    if isinstance(result, dict) and "missing_subjects" in result:
        return render_template(
            "missing_subjects.html",
            codes=result["missing_subjects"],
            file_id=file_id
        )

    if isinstance(result, dict) and "missing_teachers" in result:
        with storage_lock:
            entry["analytics"] = result.get("analytics")

        return render_template(
            "missing_teachers.html",
            subjects=result["missing_teachers"],
            file_id=file_id
        )

    excel_file = result.get("excel_file") if isinstance(result, dict) else None
    word_file = result.get("word_file") if isinstance(result, dict) else None

    if excel_file:
        save_output_file(excel_file, entry["excel_path"])

    if word_file:
        save_output_file(word_file, entry["word_path"])

    with storage_lock:
        entry = temporary_storage.get(file_id)
        if not entry:
            return render_expired()
        entry["analytics"] = result.get("analytics") if isinstance(result, dict) else None

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

    if isinstance(result, dict) and "missing_teachers" in result:
        return render_template(
            "missing_teachers.html",
            subjects=result["missing_teachers"],
            file_id=file_id
        )

    excel_file = result.get("excel_file") if isinstance(result, dict) else None
    word_file = result.get("word_file") if isinstance(result, dict) else None

    if excel_file:
        save_output_file(excel_file, entry["excel_path"])

    if word_file:
        save_output_file(word_file, entry["word_path"])

    with storage_lock:
        entry = temporary_storage.get(file_id)
        if not entry:
            return render_expired()
        entry["analytics"] = result.get("analytics") if isinstance(result, dict) else None

    return redirect(url_for("download_page", file_id=file_id))


@app.route("/download/<file_id>")
def download_page(file_id):
    entry = get_entry(file_id)

    if not entry:
        return render_expired()

    with storage_lock:
        if entry["download_started"] is None:
            entry["download_started"] = time.time()
        remaining = EXPIRY_SECONDS - (time.time() - entry["download_started"])

    analytics = entry.get("analytics") or {
        "school_avg": 0,
        "highest_percent": 0,
        "all_A1": 0,
        "top5": {}
    }

    return render_template(
        "download.html",
        file_id=file_id,
        analytics=analytics,
        remaining=int(max(0, remaining))
    )


@app.route("/download_excel/<file_id>")
def download_excel(file_id):
    entry = get_entry(file_id)

    if not entry:
        return render_expired()

    with storage_lock:
        started = entry.get("download_started")
        if started and time.time() - started > EXPIRY_SECONDS:
            session = temporary_storage.pop(file_id, None)
            if session:
                delete_session_files(session)
            return render_expired()

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

    with storage_lock:
        started = entry.get("download_started")
        if started and time.time() - started > EXPIRY_SECONDS:
            session = temporary_storage.pop(file_id, None)
            if session:
                delete_session_files(session)
            return render_expired()

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

    with storage_lock:
        started = entry.get("download_started")
        if started and time.time() - started > EXPIRY_SECONDS:
            session = temporary_storage.pop(file_id, None)
            if session:
                delete_session_files(session)
            return render_expired()

    analytics = entry.get("analytics") or {
        "school_avg": 0,
        "highest_percent": 0,
        "all_A1": 0,
        "top5": {}
    }

    return render_template(
        "report.html",
        file_id=file_id,
        analytics=analytics
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
    app.run(
        host="0.0.0.0",
        port=port,
        debug=False
    )
