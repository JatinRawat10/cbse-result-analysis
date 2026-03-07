from flask import Flask, render_template, request, send_file, redirect, url_for
from analysis import process_result
import io
import uuid
import threading
import time

app = Flask(__name__)

# In-memory temporary store:
# { file_id: { uploaded_bytes, subject_inputs, teacher_inputs, excel, word, analytics, created_at } }
temporary_storage = {}
storage_lock = threading.Lock()

EXPIRY_SECONDS = 600  # 10 minutes


# =========================
# Helper Functions
# =========================

def cleanup_expired():
    now = time.time()
    with storage_lock:
        expired_keys = [
            file_id for file_id, data in temporary_storage.items()
            if now - data.get("created_at", 0) > EXPIRY_SECONDS
        ]
        for key in expired_keys:
            del temporary_storage[key]


def get_entry(file_id):
    cleanup_expired()
    if not file_id:
        return None
    with storage_lock:
        return temporary_storage.get(file_id)


def render_expired():
    return render_template("expired_session.html"), 410


def to_bytes(file_obj):
    if file_obj is None:
        return None
    if isinstance(file_obj, bytes):
        return file_obj
    if hasattr(file_obj, "getvalue"):
        return file_obj.getvalue()
    # attempt to convert
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

    # create storage entry early so file_id is available when templates request mapping
    file_id = str(uuid.uuid4())
    with storage_lock:
        temporary_storage[file_id] = {
            "uploaded_bytes": uploaded_bytes,
            "subject_inputs": None,
            "teacher_inputs": None,
            "excel": None,
            "word": None,
            "analytics": None,
            "created_at": time.time(),
        }

    try:
        # call with no mappings initially
        result = process_result(io.BytesIO(uploaded_bytes))
    except Exception as e:
        # remove created entry on error
        with storage_lock:
            temporary_storage.pop(file_id, None)
        return f"Processing error: {e}", 500

    # If mapping for subject codes needed
    if "missing_subjects" in result:
        return render_template(
            "missing_subjects.html",
            codes=result["missing_subjects"],
            file_id=file_id
        )

    # If teacher mapping needed
    if "missing_teachers" in result:
        # store analytics if provided (may be useful)
        with storage_lock:
            entry = temporary_storage.get(file_id)
            if entry:
                entry["analytics"] = result.get("analytics")
                entry["created_at"] = time.time()
        return render_template(
            "missing_teachers.html",
            subjects=result["missing_teachers"],
            file_id=file_id
        )

    # Success case - save files and analytics
    with storage_lock:
        entry = temporary_storage.get(file_id)
        if not entry:
            return render_expired()
        entry["excel"] = to_bytes(result.get("excel_file"))
        entry["word"] = to_bytes(result.get("word_file"))
        entry["analytics"] = result.get("analytics")
        entry["created_at"] = time.time()

    return redirect(url_for("download_page", file_id=file_id))


@app.route("/submit_subjects/<file_id>", methods=["POST"])
def submit_subjects(file_id):
    entry = get_entry(file_id)
    if not entry:
        return render_expired()

    # collect non-empty form inputs
    subject_inputs = {k: v for k, v in request.form.items() if v.strip()}

    # store subject inputs so they persist
    with storage_lock:
        entry["subject_inputs"] = subject_inputs
        entry["created_at"] = time.time()

    try:
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
        # persist analytics if any
        with storage_lock:
            entry["analytics"] = result.get("analytics")
            entry["created_at"] = time.time()
        return render_template(
            "missing_teachers.html",
            subjects=result["missing_teachers"],
            file_id=file_id
        )

    # success - save outputs and analytics
    with storage_lock:
        entry = temporary_storage.get(file_id)
        if not entry:
            return render_expired()
        entry["excel"] = to_bytes(result.get("excel_file"))
        entry["word"] = to_bytes(result.get("word_file"))
        entry["analytics"] = result.get("analytics")
        entry["created_at"] = time.time()

    return redirect(url_for("download_page", file_id=file_id))


@app.route("/submit_teachers/<file_id>", methods=["POST"])
def submit_teachers(file_id):
    entry = get_entry(file_id)
    if not entry:
        return render_expired()

    teacher_inputs = {k: v for k, v in request.form.items() if v.strip()}

    # store teacher inputs before processing
    with storage_lock:
        entry["teacher_inputs"] = teacher_inputs
        entry["created_at"] = time.time()

    try:
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

    # success - save outputs and analytics
    with storage_lock:
        entry = temporary_storage.get(file_id)
        if not entry:
            return render_expired()
        entry["excel"] = to_bytes(result.get("excel_file"))
        entry["word"] = to_bytes(result.get("word_file"))
        entry["analytics"] = result.get("analytics")
        entry["created_at"] = time.time()

    return redirect(url_for("download_page", file_id=file_id))


@app.route("/download/<file_id>")
def download_page(file_id):
    entry = get_entry(file_id)
    if not entry or not entry.get("excel") or not entry.get("word"):
        return render_expired()

    # defensive/default analytics so templates won't crash
    default_analytics = {
        "school_avg": 0,
        "highest_percent": 0,
        "all_A1": 0,
        "top5": {}
    }
    analytics = entry.get("analytics") or default_analytics

    return render_template(
        "download.html",
        file_id=file_id,
        analytics=analytics
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
    # redirect to main download page which contains tabs (Download / Report / Result)
    return redirect(url_for("download_page", file_id=file_id))


if __name__ == "__main__":
    app.run(debug=True)
