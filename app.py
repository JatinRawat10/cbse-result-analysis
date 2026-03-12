from flask import Flask, render_template, request, send_file, redirect, url_for
from analysis import process_result
import os
import io
import uuid
import threading
import time

app = Flask(__name__)

# Limit upload size (2MB)
app.config["MAX_CONTENT_LENGTH"] = 2 * 1024 * 1024

# =========================
# Global Storage
# =========================

temporary_storage = {}
storage_lock = threading.Lock()

EXPIRY_SECONDS = 60  # 5 minutes
MAX_SESSIONS = 100

# CPU processing limiter
processing_limit = threading.Semaphore(2)


# =========================
# Helper Functions
# =========================

def cleanup_expired():
    now = time.time()

    with storage_lock:
        expired = []

        for file_id, data in temporary_storage.items():

            start = data.get("download_started")

            if start and now - start > EXPIRY_SECONDS:
                expired.append(file_id)

        for key in expired:
            temporary_storage.pop(key, None)


def background_cleanup():
    while True:
        time.sleep(10)
        cleanup_expired()


def get_entry(file_id):

    if not file_id:
        return None

    with storage_lock:
        entry = temporary_storage.get(file_id)

    return entry


def render_expired():
    return render_template("expired_session.html"), 410


def to_bytes(file_obj):

    if file_obj is None:
        return None

    if isinstance(file_obj, bytes):
        return file_obj

    if hasattr(file_obj, "getvalue"):
        return file_obj.getvalue()

    try:
        return bytes(file_obj)
    except Exception:
        return None


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

    uploaded_bytes = file.read()

    file_id = str(uuid.uuid4())

    with storage_lock:

        if len(temporary_storage) > MAX_SESSIONS:
            cleanup_expired()

        temporary_storage[file_id] = {
            "uploaded_bytes": uploaded_bytes,
            "subject_inputs": None,
            "teacher_inputs": None,
            "excel": None,
            "word": None,
            "analytics": None,
            "download_started": None
        }

    try:
        with processing_limit:
            result = process_result(io.BytesIO(uploaded_bytes))

    except Exception as e:

        with storage_lock:
            temporary_storage.pop(file_id, None)

        return f"Processing error: {e}", 500

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

        entry["excel"] = to_bytes(result.get("excel_file"))
        entry["word"] = to_bytes(result.get("word_file"))
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
            result = process_result(
                io.BytesIO(entry["uploaded_bytes"]),
                subject_inputs=subject_inputs,
                teacher_inputs=entry.get("teacher_inputs")
            )

    except Exception as e:
        return f"Processing error: {e}", 500

    if "missing_subjects" in result:
        return render_template(
            "missing_subjects.html",
            codes=result["missing_subjects"],
            file_id=file_id
        )

    if "missing_teachers" in result:

        with storage_lock:
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

        entry["excel"] = to_bytes(result.get("excel_file"))
        entry["word"] = to_bytes(result.get("word_file"))
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
            result = process_result(
                io.BytesIO(entry["uploaded_bytes"]),
                subject_inputs=entry.get("subject_inputs"),
                teacher_inputs=teacher_inputs
            )

    except Exception as e:
        return f"Processing error: {e}", 500

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

        entry["excel"] = to_bytes(result.get("excel_file"))
        entry["word"] = to_bytes(result.get("word_file"))
        entry["analytics"] = result.get("analytics")

    return redirect(url_for("download_page", file_id=file_id))


@app.route("/download/<file_id>")
def download_page(file_id):

    entry = get_entry(file_id)

    if not entry or not entry.get("excel") or not entry.get("word"):
        return render_expired()

    with storage_lock:
        if entry["download_started"] is None:
            entry["download_started"] = time.time()

    default_analytics = {
        "school_avg": 0,
        "highest_percent": 0,
        "all_A1": 0,
        "top5": {}
    }

    analytics = entry.get("analytics") or default_analytics

    remaining = EXPIRY_SECONDS - (time.time() - entry["download_started"])

    return render_template(
        "download.html",
        file_id=file_id,
        analytics=analytics,
        remaining=int(max(0, remaining))
    )


@app.route("/download_excel/<file_id>")
def download_excel(file_id):

    entry = get_entry(file_id)

    if not entry or not entry.get("excel"):
        return render_expired()

    return send_file(
        io.BytesIO(entry["excel"]),
        as_attachment=True,
        download_name="CBSE_Result.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/download_word/<file_id>")
def download_word(file_id):

    entry = get_entry(file_id)

    if not entry or not entry.get("word"):
        return render_expired()

    return send_file(
        io.BytesIO(entry["word"]),
        as_attachment=True,
        download_name="CBSE_Forms.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/report/<file_id>")
def report_page(file_id):

    entry = get_entry(file_id)

    if not entry:
        return render_expired()

    default_analytics = {
        "school_avg": 0,
        "highest_percent": 0,
        "all_A1": 0,
        "top5": {}
    }

    analytics = entry.get("analytics") or default_analytics

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
