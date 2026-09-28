import os
import secrets
import subprocess
import tempfile
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

app = Flask(__name__)

@app.route("/", methods=["GET"])
def home():
    return jsonify({
        "status": "online",
        "service": "Python PDF API"
    })


@app.route("/health", methods=["GET"])
def health():
    return jsonify({
        "status": "ok",
        "service": "pptx-pdf-worker"
    })

def is_authorized():
    expected_key = os.getenv("API_KEY", "")
    received_key = request.headers.get("X-API-Key", "")

    return (
        expected_key
        and received_key
        and secrets.compare_digest(expected_key, received_key)
    )


def replace_paragraph_text(paragraph, replacements):
    original_text = "".join(run.text for run in paragraph.runs)
    updated_text = original_text

    for placeholder, value in replacements.items():
        updated_text = updated_text.replace(placeholder, value)

    if updated_text != original_text:
        if paragraph.runs:
            paragraph.runs[0].text = updated_text

            for run in paragraph.runs[1:]:
                run.text = ""
        else:
            paragraph.text = updated_text


def replace_shape_text(shape, replacements):
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for child_shape in shape.shapes:
            replace_shape_text(child_shape, replacements)

    if getattr(shape, "has_text_frame", False):
        for paragraph in shape.text_frame.paragraphs:
            replace_paragraph_text(paragraph, replacements)

    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                for paragraph in cell.text_frame.paragraphs:
                    replace_paragraph_text(paragraph, replacements)


@app.get("/")
def home():
    return jsonify({
        "status": "online",
        "service": "Python PDF API"
    })


@app.get("/health")
def health():
    return jsonify({
        "status": "ok",
        "service": "pptx-pdf-worker"
    })


@app.route("/generate-certificate", methods=["POST"])
def generate_certificate():
def generate_certificate():
def generate_certificate():
    if not is_authorized():
        return jsonify({
            "success": False,
            "message": "Unauthorized"
        }), 401

    if "pptx_file" not in request.files:
        return jsonify({
            "success": False,
            "message": "pptx_file is required"
        }), 400

    execution_id = request.form.get("execution_id", "").strip()
    name = request.form.get("name", "").strip()
    ic_number = request.form.get("ic_number", "").strip()
    email = request.form.get("email", "").strip()

    if not execution_id or not name:
        return jsonify({
            "success": False,
            "message": "execution_id and name are required"
        }), 400

    uploaded_file = request.files["pptx_file"]

    replacements = {
        "<name>": name,
        "<ic_number>": ic_number,
        "<email>": email,
        "{{name}}": name,
        "{{ic_number}}": ic_number,
        "{{email}}": email
    }

    try:
        with tempfile.TemporaryDirectory() as temporary_directory:
            workdir = Path(temporary_directory)

            input_path = workdir / "template.pptx"
            output_pptx = workdir / f"{execution_id}.pptx"

            uploaded_file.save(input_path)

            presentation = Presentation(input_path)

            for slide in presentation.slides:
                for shape in slide.shapes:
                    replace_shape_text(shape, replacements)

            presentation.save(output_pptx)

            profile_directory = workdir / "libreoffice_profile"
profile_directory.mkdir(exist_ok=True)

profile_uri = profile_directory.resolve().as_uri()

subprocess.run(
    [
        "libreoffice",
        f"-env:UserInstallation={profile_uri}",
        "--headless",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        "--convert-to",
        "pdf:impress_pdf_Export",
        "--outdir",
        str(workdir),
        str(output_pptx)
    ],
    check=True,
    capture_output=True,
    text=True,
    timeout=120,
    env={
        **os.environ,
        "HOME": str(workdir)
    }
)

            output_pdf = workdir / f"{execution_id}.pdf"

            if not output_pdf.exists():
                raise RuntimeError(
                    "LibreOffice did not produce the PDF file."
                )

            return send_file(
                output_pdf,
                mimetype="application/pdf",
                as_attachment=True,
                download_name=f"{execution_id}.pdf"
            )

    except subprocess.TimeoutExpired:
        return jsonify({
            "success": False,
            "execution_id": execution_id,
            "message": "PDF conversion timed out"
        }), 504

    except Exception as error:
        app.logger.exception("PDF generation failed")

        return jsonify({
            "success": False,
            "execution_id": execution_id,
            "message": str(error)
        }), 500
