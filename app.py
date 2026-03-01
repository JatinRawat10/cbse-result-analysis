from flask import Flask, render_template, request, send_file
from analysis import process_result
import io

app = Flask(__name__)

stored_file_content = None
stored_subject_inputs = {}
stored_excel = None
stored_word = None


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    global stored_file_content, stored_subject_inputs

    file = request.files.get("file")

    if not file:
        return "No file uploaded."

    stored_file_content = file.read()
    stored_subject_inputs = {}

    result = process_result(io.BytesIO(stored_file_content))

    # 🔴 Missing subject codes
    if "missing_subjects" in result:
        return render_template(
            "missing_subjects.html",
            codes=result["missing_subjects"]
        )

    # 🔴 Missing teachers
    if "missing_teachers" in result:
        return render_template(
            "missing_teachers.html",
            subjects=result["missing_teachers"]
        )

    return handle_success(result)


@app.route("/submit_subjects", methods=["POST"])
def submit_subjects():
    global stored_file_content, stored_subject_inputs

    subject_inputs = {}

    for code in request.form:
        subject_inputs[code] = request.form[code]

    stored_subject_inputs.update(subject_inputs)

    result = process_result(
        io.BytesIO(stored_file_content),
        subject_inputs=stored_subject_inputs
    )

    if "missing_teachers" in result:
        return render_template(
            "missing_teachers.html",
            subjects=result["missing_teachers"]
        )

    return handle_success(result)


@app.route("/submit_teachers", methods=["POST"])
def submit_teachers():
    global stored_file_content, stored_subject_inputs

    teacher_inputs = {}

    for subject in request.form:
        teacher_inputs[subject] = request.form[subject]

    result = process_result(
        io.BytesIO(stored_file_content),
        subject_inputs=stored_subject_inputs,
        teacher_inputs=teacher_inputs
    )

    return handle_success(result)


def handle_success(result):
    global stored_excel, stored_word

    if "excel_file" in result:
        stored_excel = result["excel_file"]
        stored_word = result["word_file"]
        return render_template("download.html")

    return "Unexpected error."


@app.route("/download_excel")
def download_excel():
    stored_excel.seek(0)
    return send_file(
        stored_excel,
        as_attachment=True,
        download_name="CBSE_Result.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/download_word")
def download_word():
    stored_word.seek(0)
    return send_file(
        stored_word,
        as_attachment=True,
        download_name="CBSE_Forms.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


if __name__ == "__main__":
    app.run(debug=True)
