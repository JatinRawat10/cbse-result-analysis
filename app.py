from flask import Flask, render_template, request, send_file, url_for, redirect
from analysis import process_result
import io
import uuid
import threading
import time

app = Flask(__name__)
app.secret_key = "replace-this-with-a-strong-secret"

# In-memory temporary store: { file_id: { uploaded_bytes, subject_inputs, teacher_inputs, excel, word, timer } }
temporary_storage = {}
storage_lock = threading.Lock()

# Helper: delete an entry safely (also cancels timer if present)
def delete_entry(file_id):
    with storage_lock:
        entry = temporary_storage.pop(file_id, None)
        if entry and "timer" in entry and entry["timer"] is not None:
            try:
                entry["timer"].cancel()
            except Exception:
                pass

# Helper: start / restart expiry timer for a file_id (delay in seconds)
def start_expiry_timer(file_id, delay_seconds=600):
    def expire():
        delete_entry(file_id)

    with storage_lock:
        entry = temporary_storage.get(file_id)
        if not entry:
            return
        # cancel previous timer if any
        old_timer = entry.get("timer")
        if old_timer:
            try:
                old_timer.cancel()
            except Exception:
                pass
        timer = threading.Timer(delay_seconds, expire)
        entry["timer"] = timer
        timer.daemon = True
        timer.start()

# Home: your upload page (render index.html - keep your existing template)
@app.route("/")
def home():
    return render_template("index.html")


# Upload file
@app.route("/upload", methods=["POST"])
def upload():
    file = request.files.get("file")
    if not file:
        return "No file uploaded.", 400

    filename = file.filename or ""
    if not filename.lower().endswith(".txt"):
        return "Only .txt files are allowed.", 400

    # Read bytes and process using your existing logic
    uploaded_bytes = file.read()
    try:
        result = process_result(io.BytesIO(uploaded_bytes))
    except Exception as e:
        # Return error for debugging; in production you'd show friendly message and log
        return f"Processing error: {e}", 500

    # Create file_id and store uploaded bytes (so user won't need to re-upload when mapping)
    file_id = str(uuid.uuid4())
    with storage_lock:
        temporary_storage[file_id] = {
            "uploaded_bytes": uploaded_bytes,
            "subject_inputs": None,
            "teacher_inputs": None,
            "excel": None,
            "word": None,
            "timer": None
        }
    # start expiry for 10 minutes
    start_expiry_timer(file_id, delay_seconds=600)

    # If process_result reports missing subject codes, render mapping form
    if "missing_subjects" in result:
        codes = result["missing_subjects"]
        # Render a template that posts to /submit_subjects/<file_id>
        return render_template("missing_subjects.html", codes=codes, file_id=file_id)

    # If process_result reports missing teachers, render teacher form
    if "missing_teachers" in result:
        subjects = result["missing_teachers"]
        return render_template("missing_teachers.html", subjects=subjects, file_id=file_id)

    # Success: store excel/word in memory for downloads and show download page
    with storage_lock:
        temporary_storage[file_id]["excel"] = result["excel_file"]
        temporary_storage[file_id]["word"] = result["word_file"]
    # restart timer so user gets full 10 minutes from now
    start_expiry_timer(file_id, delay_seconds=600)

    return render_template("download.html", file_id=file_id)


# Submit subject code mappings (form posts here)
@app.route("/submit_subjects/<file_id>", methods=["POST"])
def submit_subjects(file_id):
    with storage_lock:
        entry = temporary_storage.get(file_id)
    if not entry:
        return "Session expired or invalid upload. Please upload again.", 410

    # Collect mappings submitted by the user: form fields like name="184" value="English"
    subject_inputs = {k: v for k, v in request.form.items() if v.strip()}
    # store provided subject inputs
    with storage_lock:
        temporary_storage[file_id]["subject_inputs"] = subject_inputs

    # Re-run processing with subject_inputs
    try:
        result = process_result(io.BytesIO(entry["uploaded_bytes"]), subject_inputs=subject_inputs)
    except Exception as e:
        return f"Processing error: {e}", 500

    if "missing_subjects" in result:
        # still missing (unlikely) — ask again
        return render_template("missing_subjects.html", codes=result["missing_subjects"], file_id=file_id)

    if "missing_teachers" in result:
        # ask for teacher names
        subjects = result["missing_teachers"]
        return render_template("missing_teachers.html", subjects=subjects, file_id=file_id)

    # Success: store outputs and restart timer
    with storage_lock:
        temporary_storage[file_id]["excel"] = result["excel_file"]
        temporary_storage[file_id]["word"] = result["word_file"]
    start_expiry_timer(file_id, delay_seconds=600)

    return render_template("download.html", file_id=file_id)


# Submit teacher name mappings (form posts here)
@app.route("/submit_teachers/<file_id>", methods=["POST"])
def submit_teachers(file_id):
    with storage_lock:
        entry = temporary_storage.get(file_id)
    if not entry:
        return "Session expired or invalid upload. Please upload again.", 410

    teacher_inputs = {k: v for k, v in request.form.items() if v.strip()}
    with storage_lock:
        temporary_storage[file_id]["teacher_inputs"] = teacher_inputs

    # Use previously stored subject_inputs if any
    subject_inputs = entry.get("subject_inputs")

    try:
        result = process_result(
            io.BytesIO(entry["uploaded_bytes"]),
            subject_inputs=subject_inputs,
            teacher_inputs=teacher_inputs
        )
    except Exception as e:
        return f"Processing error: {e}", 500

    if "missing_teachers" in result:
        return render_template("missing_teachers.html", subjects=result["missing_teachers"], file_id=file_id)

    # Success: store outputs and restart timer
    with storage_lock:
        temporary_storage[file_id]["excel"] = result["excel_file"]
        temporary_storage[file_id]["word"] = result["word_file"]
    start_expiry_timer(file_id, delay_seconds=600)

    return render_template("download.html", file_id=file_id)


# Download endpoints — do NOT remove entry on download (user wanted multiple downloads/refresh)
@app.route("/download_excel")
def download_excel():
    file_id = request.args.get("file_id")

    with storage_lock:
        entry = temporary_storage.get(file_id) if file_id else None
    if not entry or not entry.get("excel"):
        return "File expired or not found. Please upload again.", 410

    stored_excel = entry["excel"]
    stored_excel.seek(0)

    return send_file(
        stored_excel,
        as_attachment=True,
        download_name="CBSE_Result.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/download_word")
def download_word():
    file_id = request.args.get("file_id")

    with storage_lock:
        entry = temporary_storage.get(file_id) if file_id else None
    if not entry or not entry.get("word"):
        return "File expired or not found. Please upload again.", 410

    stored_word = entry["word"]
    stored_word.seek(0)

    return send_file(
        stored_word,
        as_attachment=True,
        download_name="CBSE_Forms.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# Optional: a route to check status (debugging)
@app.route("/status/<file_id>")
def status(file_id):
    with storage_lock:
        entry = temporary_storage.get(file_id)
    if not entry:
        return {"status": "expired_or_missing"}
    return {
        "status": "present",
        "has_excel": bool(entry.get("excel")),
        "has_word": bool(entry.get("word"))
    }


if __name__ == "__main__":
    app.run(debug=True)



