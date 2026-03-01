"""
EMCC USA Certificate Generator
Deploy on Railway.

The PPTX template is stored in a PRIVATE Dropbox folder.
Authentication uses app key + secret + refresh token (never expires).

Requirements: pip install flask python-pptx dropbox gunicorn
LibreOffice must be installed (via Dockerfile on Railway).
"""

import io
import os
import subprocess
import tempfile
from datetime import datetime

import dropbox
from flask import Flask, jsonify, request, send_file
from pptx import Presentation

app = Flask(__name__)

# ── Environment variables (set these in Railway) ──────────────────────────────
# API_SECRET              : shared secret to protect your endpoint
# DROPBOX_APP_KEY         : from dropbox.com/developers/apps → Settings tab
# DROPBOX_APP_SECRET      : from dropbox.com/developers/apps → Settings tab
# DROPBOX_REFRESH_TOKEN   : generated once using get_dropbox_token.py
# DROPBOX_FILE_PATH       : path to PPTX in your Dropbox
#                           e.g. /certificates/EMCC_Certificate_TEMPLATE.pptx
# ─────────────────────────────────────────────────────────────────────────────
API_SECRET            = os.environ.get("API_SECRET", "change-me-in-env")
DROPBOX_APP_KEY       = os.environ.get("DROPBOX_APP_KEY", "")
DROPBOX_APP_SECRET    = os.environ.get("DROPBOX_APP_SECRET", "")
DROPBOX_REFRESH_TOKEN = os.environ.get("DROPBOX_REFRESH_TOKEN", "")
DROPBOX_FILE_PATH     = os.environ.get("DROPBOX_FILE_PATH", "/EMCC_Certificate_TEMPLATE.pptx")

# Slide dimensions in EMUs → A4 landscape (297mm x 210mm)
# Slide cx=10693400, cy=7562850 EMUs = 297.04mm x 210.08mm
SLIDE_WIDTH_MM  = 297
SLIDE_HEIGHT_MM = 210


def download_template() -> bytes:
    """Download the PPTX template from private Dropbox using refresh token."""
    dbx = dropbox.Dropbox(
        oauth2_refresh_token=DROPBOX_REFRESH_TOKEN,
        app_key=DROPBOX_APP_KEY,
        app_secret=DROPBOX_APP_SECRET,
    )
    _, response = dbx.files_download(DROPBOX_FILE_PATH)
    return response.content


def replace_placeholders(prs: Presentation, replacements: dict) -> Presentation:
    """Replace {{PLACEHOLDER}} tokens in all text runs of a presentation."""
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    for placeholder, value in replacements.items():
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, value)
    return prs


def convert_pptx_to_pdf(pptx_path: str, output_dir: str) -> str:
    """Convert PPTX to PDF using LibreOffice headless with exact page dimensions."""

    # Write a LibreOffice filter config that forces exact A4-landscape page size
    # This prevents LibreOffice from guessing/rescaling the slide canvas
    filter_options = (
        "impress_pdf_Export:"
        "PageRange=1,"
        f"PageWidth={SLIDE_WIDTH_MM * 100},"   # in 1/100 mm units
        f"PageHeight={SLIDE_HEIGHT_MM * 100}"   # in 1/100 mm units
    )

    result = subprocess.run(
        [
            "soffice",
            "--headless",
            "--infilter=Impress MS PowerPoint 2007 XML",
            f"--convert-to=pdf:{filter_options}",
            "--outdir", output_dir,
            pptx_path,
        ],
        capture_output=True,
        text=True,
        timeout=60,
        env={**os.environ, "HOME": "/tmp"},  # prevent soffice profile conflicts
    )

    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")

    base_name = os.path.splitext(os.path.basename(pptx_path))[0]
    pdf_path = os.path.join(output_dir, base_name + ".pdf")
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF not found after conversion: {pdf_path}")
    return pdf_path


def format_date(dt: datetime) -> str:
    """Format a datetime as '27 February 2026'."""
    return dt.strftime("%-d %B %Y")


@app.route("/generate-certificate", methods=["POST"])
def generate_certificate():
    # Authenticate
    if request.headers.get("X-API-Secret", "") != API_SECRET:
        return jsonify({"error": "Unauthorized"}), 401

    # Parse input
    data = request.get_json(force=True)
    first_name = data.get("first_name", "").strip()
    last_name  = data.get("last_name", "").strip()
    issued_date_str = data.get("issued_date")

    if not first_name or not last_name:
        return jsonify({"error": "first_name and last_name are required"}), 400

    # Compute dates
    issued_dt = datetime.fromisoformat(issued_date_str) if issued_date_str else datetime.today()
    valid_dt  = issued_dt.replace(year=issued_dt.year + 1)

    replacements = {
        "{{FIRST_NAME}}":  first_name,
        "{{LAST_NAME}}":   last_name,
        "{{ISSUED_DATE}}": format_date(issued_dt),
        "{{VALID_DATE}}":  format_date(valid_dt),
    }

    # Download template from private Dropbox
    try:
        template_bytes = download_template()
    except Exception as e:
        return jsonify({"error": f"Failed to download template: {str(e)}"}), 500

    # Generate certificate
    with tempfile.TemporaryDirectory() as tmpdir:
        prs = Presentation(io.BytesIO(template_bytes))
        prs = replace_placeholders(prs, replacements)

        pptx_out = os.path.join(tmpdir, f"certificate_{first_name}_{last_name}.pptx")
        prs.save(pptx_out)

        try:
            pdf_path = convert_pptx_to_pdf(pptx_out, tmpdir)
        except Exception as e:
            return jsonify({"error": f"PDF conversion failed: {str(e)}"}), 500

        return send_file(
            pdf_path,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=f"EMCC_USA_Certificate_{first_name}_{last_name}.pdf",
        )


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
