import io
import os
import subprocess
import tempfile
from datetime import datetime

import dropbox
from flask import Flask, jsonify, request, send_file
from pptx import Presentation

app = Flask(__name__)
API_SECRET           = os.environ.get("API_SECRET", "change-me-in-env")
DROPBOX_ACCESS_TOKEN = os.environ.get("DROPBOX_ACCESS_TOKEN", "")
DROPBOX_FILE_PATH    = os.environ.get("DROPBOX_FILE_PATH", "/EMCC_Certificate_TEMPLATE.pptx")


def download_template() -> bytes:
    """Download the PPTX template from private Dropbox and return raw bytes."""
    
    dbx = dropbox.Dropbox(DROPBOX_ACCESS_TOKEN)
    print("=== Checking account info ===")
    account = dbx.users_get_current_account()
    print(f"Logged in as: {account.email}")
    
    print("\n=== Listing root folder contents ===")
    try:
        result = dbx.files_list_folder("")
        if result.entries:
            for entry in result.entries:
                print(f"  {entry.path_display}")
        else:
            print("  (root folder is empty — app likely has App Folder access only)")
    except Exception as e:
        print(f"  Error listing root: {e}")

    print("\n=== Trying to access the file directly ===")
    try:
        metadata = dbx.files_get_metadata(DROPBOX_FILE_PATH)
        print(f"  Found: {metadata.path_display} ({metadata.size} bytes)")
    except Exception as e:
        print(f"  Error: {e}")
    
    print("\n=== Listing /certificates/ folder ===")
    try:
        result = dbx.files_list_folder("/certificates")
        for entry in result.entries:
            print(f"  {entry.path_display}")
    except Exception as e:
        print(f"  Error listing /certificates: {e}")
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
    """Convert PPTX to PDF using LibreOffice headless."""
    result = subprocess.run(
        [
            "soffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            pptx_path,
        ],
        capture_output=True,
        text=True,
        timeout=60,
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
    issued_date_str = data.get("issued_date")  # ISO format: 2026-02-27

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
