"""
Configuration-driven Flask application for reconciliation tool.
All processing logic is now driven by configuration files.
"""

import os
import pandas as pd
import tempfile
import shutil
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
from config import ReconciliationConfig
from processors import ReconciliationProcessor
from rate_tool_integration import run_rate_analysis, save_uploaded_file

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "uploads"
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)

# Global instances
config = ReconciliationConfig()
processor = ReconciliationProcessor()
last_output = None

@app.route("/", methods=["GET", "POST"])
def index():
    """Main route - handles both form display and processing"""
    global last_output
    result = None
    error_message = None
    recon_type = request.form.get("recon_type")

    if request.method == "POST":
        try:
            # Validate reconciliation type
            if not config.validate_reconciliation_type(recon_type):
                error_message = "Please select a valid reconciliation type."
            else:
                # Process the reconciliation
                result = processor.process(recon_type, request.files)
                last_output = pd.DataFrame(result) if result else None
                
        except Exception as e:
            error_message = f"Error processing files: {str(e)}"
    
    # Get all available reconciliation types for the form
    reconciliation_types = config.get_all_types()
    
    return render_template("index_dynamic.html", 
                         result=result, 
                         recon_type=recon_type, 
                         error_message=error_message,
                         reconciliation_types=reconciliation_types,
                         result_config=config.RESULT_TABLES.get(recon_type, {}))

@app.route("/api/reconciliation-types")
def get_reconciliation_types():
    """API endpoint to get all reconciliation types configuration"""
    return jsonify(config.get_all_types())

@app.route("/rates-file", methods=["GET", "POST"])
def rates_file():
    """Integrated rates calculator tab"""
    report = None
    error_message = None

    if request.method == "POST":
        temp_dir = tempfile.mkdtemp(prefix="rates_tab_")
        try:
            summary_file = request.files.get("summary_file")
            if not summary_file or summary_file.filename == "":
                raise ValueError("Summary file is required.")

            summary_path = save_uploaded_file(summary_file, temp_dir)
            card_path = save_uploaded_file(request.files.get("card_file"), temp_dir)
            international_path = save_uploaded_file(request.files.get("international_file"), temp_dir)
            domestic_path = save_uploaded_file(request.files.get("domestic_file"), temp_dir)
            dispute_path = save_uploaded_file(request.files.get("dispute_file"), temp_dir)

            file_paths = {
                "summary": summary_path,
                "card": card_path,
                "international": international_path,
                "domestic": domestic_path,
                "dispute": dispute_path
            }

            report = run_rate_analysis(file_paths)
        except Exception as exc:
            error_message = str(exc)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    return render_template("rates_tab.html", report=report, error_message=error_message)

@app.route("/download")
def download():
    """Download reconciliation results as Excel file"""
    global last_output
    if last_output is not None:
        path = "reconciliation_output.xlsx"
        last_output.to_excel(path, index=False)
        return send_file(path, as_attachment=True)
    return "No reconciliation results available to download.", 404

@app.route("/health")
def health_check():
    """Health check endpoint"""
    return jsonify({
        "status": "healthy",
        "available_types": list(config.get_all_types().keys()),
        "version": "2.0-config-driven"
    })

@app.errorhandler(404)
def not_found_error(error):
    return render_template("error.html", error="Page not found"), 404

@app.errorhandler(500)
def internal_error(error):
    return render_template("error.html", error="Internal server error"), 500

if __name__ == "__main__":
    app.run(debug=True, port=5000)

