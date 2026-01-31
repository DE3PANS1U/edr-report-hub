"""
EDR Report Generator Web Application
Flask backend for uploading Excel files and generating PowerPoint reports
"""
from flask import Flask, request, render_template, send_file, jsonify
import os
import pandas as pd
from datetime import datetime
from werkzeug.utils import secure_filename
import traceback
from edr_report_generator_custom import generate_edr_report

app = Flask(__name__)

# Configuration
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB

# Create necessary directories
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE


def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def validate_excel_format(filepath):
    """
    Validate that Excel file has required columns
    Returns (is_valid, error_message)
    """
    try:
        df = pd.read_excel(filepath)
        
        required_columns = [
            'Entity',
            'Alert Trend',
            'Filename',
            'Alert Severity',
            'Alert Status',
            'Alert Efficiency'
        ]
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            return False, f"Missing required columns: {', '.join(missing_columns)}"
        
        if len(df) == 0:
            return False, "Excel file is empty"
        
        return True, None
        
    except Exception as e:
        return False, f"Error reading Excel file: {str(e)}"


@app.route('/')
def index():
    """Render the main upload page"""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and report generation"""
    try:
        # Check if file is present
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file type. Please upload .xlsx or .xls file'}), 400
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        saved_filename = f"{timestamp}_{filename}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], saved_filename)
        file.save(filepath)
        
        # Validate Excel format
        is_valid, error_msg = validate_excel_format(filepath)
        if not is_valid:
            os.remove(filepath)  # Clean up invalid file
            return jsonify({'error': error_msg}), 400
        
        # Call the report generator
        generated_path = generate_edr_report(
            excel_file_path=filepath,
            date_range="01st Dec 25 to 07th Dec 25",
            output_dir=app.config['OUTPUT_FOLDER']
        )
        
        output_filename = os.path.basename(generated_path)
        
        # Clean up uploaded file
        os.remove(filepath)
        
        return jsonify({
            'success': True,
            'filename': output_filename,
            'message': 'Report generated successfully!'
        })
        
    except Exception as e:
        # Log error for debugging
        print(f"Error: {str(e)}")
        print(traceback.format_exc())
        return jsonify({'error': f'Error generating report: {str(e)}'}), 500


@app.route('/download/<filename>')
def download_file(filename):
    """Download generated report"""
    try:
        filepath = os.path.join(app.config['OUTPUT_FOLDER'], secure_filename(filename))
        
        if not os.path.exists(filepath):
            return jsonify({'error': 'File not found'}), 404
        
        return send_file(
            filepath,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        
    except Exception as e:
        return jsonify({'error': f'Error downloading file: {str(e)}'}), 500


@app.route('/cleanup/<filename>', methods=['POST'])
def cleanup_file(filename):
    """Clean up generated file after download"""
    try:
        filepath = os.path.join(app.config['OUTPUT_FOLDER'], secure_filename(filename))
        if os.path.exists(filepath):
            os.remove(filepath)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
