# -*- coding: utf-8 -*-
"""entop_sdl_app.py — enTop Clients – Service Delivery Hub with DWR 4-step form."""

from flask import Flask, request, redirect, send_file
import os
import pandas as pd
import numpy as np
from werkzeug.utils import secure_filename
from daily_internal_qc_entop_sdl_prod import select_random_series_from_stream  # Import function

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "generated_reports")
BIND_UPLOAD_DIR = os.path.join(UPLOAD_DIR, "dwr_bind_steps")

STEP_TEMPLATES = {
    1: os.path.join(BASE_DIR, "Daily Work Report Step 1 Template - enTop v1.0.0.docx"),
    2: os.path.join(BASE_DIR, "Daily Work Report Step 2 Template - enTop v1.0.0.docx"),
    3: os.path.join(BASE_DIR, "Daily Work Report Step 3 Template - enTop v1.0.0.docx"),
    4: os.path.join(BASE_DIR, "Daily Work Report Step 4 Template - enTop v1.0.0.docx"),
}

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(BIND_UPLOAD_DIR, exist_ok=True)


@app.route('/')
def home():
    return '''
    <html>
        <head><title>enTop Service Hub</title></head>
        <body>
            <h1>enTop Clients – Service Delivery Hub</h1>
            <h3>Select an action below:</h3>
            <form action="execute" method="post" enctype="multipart/form-data">
                <input type="file" name="file">
                <button type="submit" name="script" value="daily_qc">SDL Daily QC Jobs</button>
            </form>
            <form action="execute" method="post" enctype="multipart/form-data">
                <input type="file" name="file">
                <button type="submit" name="script" value="draft_report">Draft SDL Daily Report</button>
            </form>
            <form action="execute" method="post">
                <button type="submit" name="script" value="script3">Run Script 3</button>
                <button type="submit" name="script" value="script4">Go to Daily Work Report</button>
            </form>
        </body>
    </html>
    '''


@app.route('/daily-work-report', methods=['GET'])
def daily_work_report():
    """Display the Daily Work Report upload/bind page."""
    return '''
    <html>
        <head>
            <title>Generate SDL Daily Work Report</title>
            <style>
                body {
                    font-family: Arial, Helvetica, sans-serif;
                    margin: 28px 32px;
                    color: #000;
                    background: #fff;
                }
                h1 {
                    font-family: "Times New Roman", Times, serif;
                    font-size: 42px;
                    margin: 0 0 92px 0;
                    font-weight: 700;
                    line-height: 1.05;
                }
                p {
                    font-size: 20px;
                    margin: 0 0 16px 0;
                    line-height: 1.3;
                }
                .filename-sample {
                    font-family: "Times New Roman", Times, serif;
                    font-style: italic;
                    white-space: nowrap;
                }
                .dwr-panel {
                    margin-top: 8px;
                    max-width: 720px;
                }
                .upload-row {
                    display: block;
                    margin: 0 0 14px 0;
                }
                .upload-row input[type="file"] {
                    font-size: 18px;
                    font-family: Arial, Helvetica, sans-serif;
                }
                .helper-text {
                    font-size: 18px;
                    margin: 0 0 16px 0;
                    color: #222;
                }
                .step-grid {
                    display: grid;
                    grid-template-columns: 380px 230px;
                    column-gap: 64px;
                    row-gap: 13px;
                    align-items: center;
                    margin-top: 8px;
                }
                .step-label {
                    font-size: 18px;
                    line-height: 1.24;
                    min-height: 34px;
                    max-width: 360px;
                    display: flex;
                    align-items: center;
                    color: #000;
                }
                .step-title {
                    font-weight: 700;
                }
                .step-btn {
                    width: 220px;
                    min-height: 34px;
                    font-family: Arial, Helvetica, sans-serif;
                    font-size: 18px;
                    padding: 4px 12px;
                    border: 1px solid #777;
                    border-radius: 3px;
                    color: #111;
                    cursor: pointer;
                    text-align: center;
                    box-shadow: 0 1px 1px rgba(0,0,0,0.15);
                }
                .step1 { background: #ffff99; }
                .step2 { background: #efe3a0; }
                .step3 { background: #ffb27d; }
                .step4 { background: #58d68d; }
                .bind-panel {
                    margin-top: 42px;
                    max-width: 720px;
                    padding-top: 18px;
                    border-top: 1px solid #ddd;
                }
                .bind-heading {
                    font-family: "Times New Roman", Times, serif;
                    font-size: 28px;
                    font-weight: 700;
                    margin: 0 0 10px 0;
                }
                .bind-note {
                    font-size: 17px;
                    margin: 0 0 14px 0;
                    color: #222;
                }
                .bind-grid {
                    display: grid;
                    grid-template-columns: 130px 1fr;
                    column-gap: 14px;
                    row-gap: 10px;
                    align-items: center;
                    margin-bottom: 16px;
                }
                .bind-label {
                    font-size: 17px;
                    font-weight: 700;
                }
                .bind-grid input[type="file"] {
                    font-size: 16px;
                    font-family: Arial, Helvetica, sans-serif;
                }
                .bind-btn {
                    min-width: 220px;
                    min-height: 34px;
                    font-family: Arial, Helvetica, sans-serif;
                    font-size: 18px;
                    padding: 4px 14px;
                    border: 1px solid #777;
                    border-radius: 3px;
                    color: #111;
                    cursor: pointer;
                    text-align: center;
                    background: #c9daf8;
                    box-shadow: 0 1px 1px rgba(0,0,0,0.15);
                }
                .back-link {
                    display: inline-block;
                    margin-top: 54px;
                    font-family: "Times New Roman", Times, serif;
                    font-size: 22px;
                }
            </style>
        </head>
        <body>
            <h1>Generate SDL Daily Work Report</h1>

            <p>Upload the daily Excel file named like <span class="filename-sample">DD-MM-YYYY_SDL_Daily_Work_Report.xlsx</span>.</p>

            <form class="dwr-panel" action="generate-dwr-step" method="post" enctype="multipart/form-data">
                <div class="upload-row">
                    <input type="file" name="file" accept=".xlsx" required>
                </div>
                <p class="helper-text">Use the same uploaded Excel file to generate any required DWR step.</p>

                <div class="step-grid">
                    <div class="step-label"><span><span class="step-title">Step 1:</span> Operational Tables</span></div>
                    <button class="step-btn step1" type="submit" name="step" value="1">Generate Step 1 DWR</button>

                    <div class="step-label"><span><span class="step-title">Step 2:</span> Daily Metrics at a Glance, top ordering/quoted customer sections</span></div>
                    <button class="step-btn step2" type="submit" name="step" value="2">Generate Step 2 DWR</button>

                    <div class="step-label"><span><span class="step-title">Step 3:</span> Leaderboard, descriptive statistics, histogram-related chart</span></div>
                    <button class="step-btn step3" type="submit" name="step" value="3">Generate Step 3 DWR</button>

                    <div class="step-label"><span><span class="step-title">Step 4:</span> Key Observations</span></div>
                    <button class="step-btn step4" type="submit" name="step" value="4">Generate Step 4 DWR</button>
                </div>
            </form>

            <form class="bind-panel" action="bind-dwr-step-files" method="post" enctype="multipart/form-data">
                <div class="bind-heading">Bind DWR Step Files</div>
                <p class="bind-note">Upload the four generated Word files, then bind them into one full DWR document in Step 1 → Step 2 → Step 3 → Step 4 order.</p>
                <div class="bind-grid">
                    <label class="bind-label" for="step1_docx">Step 1 file</label>
                    <input id="step1_docx" type="file" name="step1_docx" accept=".docx" required>

                    <label class="bind-label" for="step2_docx">Step 2 file</label>
                    <input id="step2_docx" type="file" name="step2_docx" accept=".docx" required>

                    <label class="bind-label" for="step3_docx">Step 3 file</label>
                    <input id="step3_docx" type="file" name="step3_docx" accept=".docx" required>

                    <label class="bind-label" for="step4_docx">Step 4 file</label>
                    <input id="step4_docx" type="file" name="step4_docx" accept=".docx" required>
                </div>
                <button class="bind-btn" type="submit">Bind DWR Step Files</button>
            </form>

            <a class="back-link" href="./">Back to Dashboard</a>
        </body>
    </html>
    '''

def _handle_dwr_step(step_number):
    """Shared handler for DWR step upload and generation."""
    back_link = "daily-work-report"

    if 'file' not in request.files:
        return f"<h3>Error: No file uploaded.</h3><br><a href='{back_link}'>Go Back</a>"

    file = request.files['file']
    if file.filename == '':
        return f"<h3>Error: No file selected.</h3><br><a href='{back_link}'>Go Back</a>"

    filename = secure_filename(file.filename)
    if not filename.lower().endswith('.xlsx'):
        return f"<h3>Error: Please upload an .xlsx file.</h3><br><a href='{back_link}'>Go Back</a>"

    input_path = os.path.join(UPLOAD_DIR, filename)
    file.save(input_path)

    template_path = STEP_TEMPLATES.get(step_number)
    if not template_path or not os.path.exists(template_path):
        return (
            f"<h3>Error: Step {step_number} template file was not found on the server.</h3>"
            f"<p>Expected path: {template_path}</p>"
            f"<br><a href='{back_link}'>Go Back</a>"
        )

    try:
        if step_number == 1:
            from sdl_dwr_step1_generator import generate_dwr_step1
            output_path = generate_dwr_step1(input_path, template_path=template_path, output_dir=OUTPUT_DIR)
        elif step_number == 2:
            from sdl_dwr_step2_generator import generate_dwr_step2
            output_path = generate_dwr_step2(input_path, template_path=template_path, output_dir=OUTPUT_DIR)
        elif step_number == 3:
            from sdl_dwr_step3_generator import generate_dwr_step3
            output_path = generate_dwr_step3(input_path, template_path=template_path, output_dir=OUTPUT_DIR)
        elif step_number == 4:
            from sdl_dwr_step4_generator import generate_dwr_step4
            output_path = generate_dwr_step4(input_path, template_path=template_path, output_dir=OUTPUT_DIR)
        else:
            return f"<h3>Error: Invalid DWR step.</h3><br><a href='{back_link}'>Go Back</a>"

        return send_file(output_path, as_attachment=True, download_name=os.path.basename(output_path))

    except ModuleNotFoundError as exc:
        return (
            f"<h3>DWR Step {step_number} generator is not deployed yet.</h3>"
            f"<p>Upload <code>sdl_dwr_step{step_number}_generator.py</code> to the app folder, then restart the Python app.</p>"
            f"<pre>{exc}</pre>"
            f"<br><a href='{back_link}'>Go Back</a>"
        )
    except Exception as exc:
        return (
            f"<h3>Error generating DWR Step {step_number}:</h3>"
            f"<pre>{exc}</pre>"
            f"<br><a href='{back_link}'>Go Back</a>"
        )



@app.route('/generate-dwr-step', methods=['POST'])
def generate_dwr_step_route():
    """Generate the selected DWR step using one shared uploaded Excel file."""
    try:
        step_number = int(request.form.get('step', '0'))
    except ValueError:
        step_number = 0

    if step_number not in (1, 2, 3, 4):
        return "<h3>Error: Invalid DWR step selected.</h3><br><a href='daily-work-report'>Go Back</a>"

    return _handle_dwr_step(step_number)


@app.route('/generate-dwr-step1', methods=['POST'])
def generate_dwr_step1_route():
    return _handle_dwr_step(1)


@app.route('/generate-dwr-step2', methods=['POST'])
def generate_dwr_step2_route():
    return _handle_dwr_step(2)


@app.route('/generate-dwr-step3', methods=['POST'])
def generate_dwr_step3_route():
    return _handle_dwr_step(3)


@app.route('/generate-dwr-step4', methods=['POST'])
def generate_dwr_step4_route():
    return _handle_dwr_step(4)


def _save_bind_upload(field_name, label, run_dir):
    """Save one uploaded Step .docx file for the binder form."""
    if field_name not in request.files:
        raise ValueError(f"{label}: no file uploaded.")
    file = request.files[field_name]
    if file.filename == '':
        raise ValueError(f"{label}: no file selected.")
    filename = secure_filename(file.filename)
    if not filename.lower().endswith('.docx'):
        raise ValueError(f"{label}: please upload a .docx file.")
    os.makedirs(run_dir, exist_ok=True)
    path = os.path.join(run_dir, filename)
    file.save(path)
    return path


@app.route('/bind-dwr-step-files', methods=['POST'])
def bind_dwr_step_files_route():
    """Bind uploaded Step 1-4 DWR .docx files into one full DWR .docx."""
    back_link = "daily-work-report"
    try:
        from sdl_dwr_binder import bind_dwr_step_files

        run_stamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
        run_dir = os.path.join(BIND_UPLOAD_DIR, run_stamp)

        step1_path = _save_bind_upload('step1_docx', 'Step 1', run_dir)
        step2_path = _save_bind_upload('step2_docx', 'Step 2', run_dir)
        step3_path = _save_bind_upload('step3_docx', 'Step 3', run_dir)
        step4_path = _save_bind_upload('step4_docx', 'Step 4', run_dir)

        output_path = bind_dwr_step_files(
            step1_path,
            step2_path,
            step3_path,
            step4_path,
            output_dir=OUTPUT_DIR,
        )
        return send_file(output_path, as_attachment=True, download_name=os.path.basename(output_path))

    except ModuleNotFoundError as exc:
        return (
            "<h3>DWR binder is not fully deployed.</h3>"
            "<p>Confirm <code>sdl_dwr_binder.py</code> is in the app folder and <code>docxcompose</code> is installed.</p>"
            f"<pre>{exc}</pre>"
            f"<br><a href='{back_link}'>Go Back</a>"
        )
    except Exception as exc:
        return (
            "<h3>Error binding DWR step files:</h3>"
            f"<pre>{exc}</pre>"
            f"<br><a href='{back_link}'>Go Back</a>"
        )


@app.route('/execute', methods=['POST'])
def execute():
    script_name = request.form.get('script')

    if script_name == "daily_qc":
        if 'file' not in request.files:
            return "<h3>Error: No file uploaded.</h3><br><a href='./'>Go Back</a>"

        file = request.files['file']
        if file.filename == '':
            return "<h3>Error: No file selected.</h3><br><a href='./'>Go Back</a>"

        selected_series = select_random_series_from_stream(file.stream)
        return f"<h3>Selected Series: {selected_series}</h3><br><a href='./'>Go Back</a>"

    elif script_name == "draft_report":
        if 'file' not in request.files:
            return "<h3>Error: No file uploaded.</h3><br><a href='./'>Go Back</a>"

        file = request.files['file']
        if file.filename == '':
            return "<h3>Error: No file selected.</h3><br><a href='./'>Go Back</a>"

        input_file_path = os.path.join(UPLOAD_DIR, "uploaded_input.xlsx")
        output_file_path = os.path.join(OUTPUT_DIR, "transformed_output.xlsx")

        file.save(input_file_path)
        from draft_sdl_daily_report import transform_sdl_report  # Import transformation function

        transform_sdl_report(input_file_path, output_file_path)

        return send_file(output_file_path, as_attachment=True, download_name="Draft_SDL_Report.xlsx")

    elif script_name == "script3":
        return "<h3>Script 3 executed successfully.</h3><br><a href='./'>Go Back</a>"

    elif script_name == "script4":
        return redirect('daily-work-report')

    return "<h3>Error: Invalid selection.</h3><br><a href='./'>Go Back</a>"


# Passenger WSGI entry point
if __name__ != '__main__':
    application = app
