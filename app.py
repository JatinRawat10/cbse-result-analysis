from flask import Flask, render_template, request, send_file, redirect, url_for
from analysis import process_result
import io
import uuid
import threading

app = Flask(__name__)
app.secret_key = "replace-this-with-a-strong-secret"

# In-memory temporary store: { file_id: { uploaded_bytes, subject_inputs, teacher_inputs, excel, word, timer } }
temporary_storage = {}
storage_lock = threading.Lock()


def get_entry(file_id):
    if not file_id:
        return None
    with storage_lock:
        return temporary_storage.get(file_id)


# Helper: delete an entry safely (also cancels timer if present)
def delete_entry(file_id):
    with storage_lock:
        entry = temporary_storage.pop(file_id, None)
        if entry and entry.get("timer") is not None:
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


def render_reupload_message():
    return render_template("expired_session.html"), 410


def _to_bytes(file_obj):
    if file_obj is None:
        return None
    if isinstance(file_obj, bytes):
        return file_obj
    if hasattr(file_obj, "getvalue"):
        return file_obj.getvalue()
    return bytes(file_obj)


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/download/<file_id>")
def download_page(file_id):
    entry = get_entry(file_id)
    if not entry or not entry.get("excel") or not entry.get("word"):
        return render_reupload_message()
    return render_template("download.html", file_id=file_id)


@app.route("/upload", methods=["POST"])
def upload():
    file = request.files.get("file")
    if not file:
        return "No file uploaded.", 400

    filename = file.filename or ""
    if not filename.lower().endswith(".txt"):
        return "Only .txt files are allowed.", 400

    uploaded_bytes = file.read()
    try:
        result = process_result(io.BytesIO(uploaded_bytes))
    except Exception as e:
        return f"Processing error: {e}", 500

    file_id = str(uuid.uuid4())
    with storage_lock:
        temporary_storage[file_id] = {
            "uploaded_bytes": uploaded_bytes,
            "subject_inputs": None,
            "teacher_inputs": None,
            "excel": None,
            "word": None,
            "timer": None,
        }
    start_expiry_timer(file_id, delay_seconds=600)

    if "missing_subjects" in result:
        return render_template("missing_subjects.html", codes=result["missing_subjects"], file_id=file_id)

    if "missing_teachers" in result:
        return render_template("missing_teachers.html", subjects=result["missing_teachers"], file_id=file_id)

    with storage_lock:
        entry = temporary_storage.get(file_id)
        if not entry:
            return render_reupload_message()
        entry["excel"] = _to_bytes(result["excel_file"])
        entry["word"] = _to_bytes(result["word_file"])

    start_expiry_timer(file_id, delay_seconds=600)
    return redirect(url_for("download_page", file_id=file_id))


@app.route("/submit_subjects/<file_id>", methods=["POST"])
def submit_subjects(file_id):
    entry = get_entry(file_id)
    if not entry:
        return render_reupload_message()

    subject_inputs = {k: v for k, v in request.form.items() if v.strip()}

    with storage_lock:
        latest_entry = temporary_storage.get(file_id)
        if not latest_entry:
            return render_reupload_message()
        latest_entry["subject_inputs"] = subject_inputs
        uploaded_bytes = latest_entry["uploaded_bytes"]

    try:
        result = process_result(io.BytesIO(uploaded_bytes), subject_inputs=subject_inputs)
    except Exception as e:
        return f"Processing error: {e}", 500

    if "missing_subjects" in result:
        return render_template("missing_subjects.html", codes=result["missing_subjects"], file_id=file_id)

    if "missing_teachers" in result:
        return render_template("missing_teachers.html", subjects=result["missing_teachers"], file_id=file_id)

    with storage_lock:
        latest_entry = temporary_storage.get(file_id)
        if not latest_entry:
            return render_reupload_message()
        latest_entry["excel"] = _to_bytes(result["excel_file"])
        latest_entry["word"] = _to_bytes(result["word_file"])

    start_expiry_timer(file_id, delay_seconds=600)
    return redirect(url_for("download_page", file_id=file_id))


@app.route("/submit_teachers/<file_id>", methods=["POST"])
def submit_teachers(file_id):
    entry = get_entry(file_id)
    if not entry:
        return render_reupload_message()

    teacher_inputs = {k: v for k, v in request.form.items() if v.strip()}

    with storage_lock:
        latest_entry = temporary_storage.get(file_id)
        if not latest_entry:
            return render_reupload_message()
        latest_entry["teacher_inputs"] = teacher_inputs
        uploaded_bytes = latest_entry["uploaded_bytes"]
        subject_inputs = latest_entry.get("subject_inputs")

    try:
        result = process_result(
            io.BytesIO(uploaded_bytes),
            subject_inputs=subject_inputs,
            teacher_inputs=teacher_inputs,
        )
    except Exception as e:
        return f"Processing error: {e}", 500

    if "missing_teachers" in result:
        return render_template("missing_teachers.html", subjects=result["missing_teachers"], file_id=file_id)

    with storage_lock:
        latest_entry = temporary_storage.get(file_id)
        if not latest_entry:
            return render_reupload_message()
        latest_entry["excel"] = _to_bytes(result["excel_file"])
        latest_entry["word"] = _to_bytes(result["word_file"])

    start_expiry_timer(file_id, delay_seconds=600)
    return redirect(url_for("download_page", file_id=file_id))


@app.route("/download_excel")
def download_excel():
    file_id = request.args.get("file_id")
    entry = get_entry(file_id)
    if not entry or not entry.get("excel"):
        return render_reupload_message()

    stored_excel = io.BytesIO(entry["excel"])

    return send_file(
        stored_excel,
        as_attachment=True,
        download_name="CBSE_Result.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/download_word")
def download_word():
    file_id = request.args.get("file_id")
    entry = get_entry(file_id)
    if not entry or not entry.get("word"):
        return render_reupload_message()

    stored_word = io.BytesIO(entry["word"])

    return send_file(
        stored_word,
        as_attachment=True,
        download_name="CBSE_Forms.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/status/<file_id>")
def status(file_id):
    entry = get_entry(file_id)
    if not entry:
        return {"status": "expired_or_missing"}

    return {
        "status": "present",
        "has_excel": bool(entry.get("excel")),
        "has_word": bool(entry.get("word")),
    }


if __name__ == "__main__":
    app.run(debug=True)
