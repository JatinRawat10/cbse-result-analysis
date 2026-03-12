from flask import Flask, render_template, request, send_file, redirect, url_for
from analysis import process_result
import os
import uuid
import threading
import time
import tempfile

app = Flask(__name__)

# =========================
# Configuration
# =========================

app.config["MAX_CONTENT_LENGTH"] = 2 * 1024 * 1024  # 2MB upload limit

EXPIRY_SECONDS = 300
MAX_SESSIONS = 100

processing_limit = threading.Semaphore(2)

# Temporary directories
UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "cbse_uploads")
OUTPUT_DIR = os.path.join(tempfile.gettempdir(), "cbse_outputs")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# =========================
# Global session store
# =========================

temporary_storage = {}
storage_lock = threading.Lock()

# =========================
# Error Handling
# =========================

@app.errorhandler(413)
def file_too_large(e):
    return """
    <script>
    alert("File must be smaller than 2 MB.");
    window.history.back();
    </script>
    """

# =========================
# Helper Functions
# =========================

def cleanup_expired():

    now = time.time()

    with storage_lock:

        expired = []

        for file_id, data in temporary_storage.items():

            created = data["created_at"]
            started = data.get("download_started")

            if started is None:
                if now - created > EXPIRY_SECONDS:
                    expired.append(file_id)

            else:
                if now - started > EXPIRY_SECONDS:
                    expired.append(file_id)

        for key in expired:

            session = temporary_storage.pop(key, None)

            if session:

                for path in [session.get("txt_path"),
                             session.get("excel_path"),
                             session.get("word_path")]:

                    if path and os.path.exists(path):
                        try:
                            os.remove(path)
                        except:
                            pass


def background_cleanup():
    while True:
        time.sleep(10)
        cleanup_expired()


def get_entry(file_id):

    cleanup_expired()

    if not file_id:
        return None

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

    file.save(txt_path)

    with storage_lock:

        if len(temporary_storage) > MAX_SESSIONS:
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

            result = process_result(txt_path)

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

        if result.get("excel_file"):
            with open(entry["excel_path"], "wb") as f:
                f.write(result["excel_file"].getvalue())

        if result.get("word_file"):
            with open(entry["word_path"], "wb") as f:
                f.write(result["word_file"].getvalue())

        entry["analytics"] = result.get("analytics")

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

    path = entry.get("excel_path")

    if not path or not os.path.exists(path):
        return render_expired()

    return send_file(path, as_attachment=True)


@app.route("/download_word/<file_id>")
def download_word(file_id):

    entry = get_entry(file_id)

    if not entry:
        return render_expired()

    path = entry.get("word_path")

    if not path or not os.path.exists(path):
        return render_expired()

    return send_file(path, as_attachment=True)


@app.route("/report/<file_id>")
def report_page(file_id):

    entry = get_entry(file_id)

    if not entry:
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
# Start Cleanup Thread
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
