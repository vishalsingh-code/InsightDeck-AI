#!/usr/bin/env python3
"""
Web Dashboard for CSV/Excel-to-PowerPoint AI Analyzer
Provides a web interface for uploading files and generating presentations
"""

import os
import tempfile
import uuid
from datetime import datetime
from flask import Flask, request, render_template, send_file, flash, redirect, url_for, jsonify
from werkzeug.utils import secure_filename
from advanced_ppt_generator import CSVPPTGenerator

app = Flask(__name__, static_url_path='/static', static_folder='static')
app.secret_key = 'your-secret-key-change-this-in-production'

# Configuration
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'csv', 'xlsx', 'xls'}
MAX_FILE_SIZE = 16 * 1024 * 1024  # 16MB max file size

# Create upload directory if it doesn't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    """Check if the uploaded file has an allowed extension"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_file_info(file_path):
    """Get information about uploaded file"""
    try:
        generator = CSVPPTGenerator()
        file_type = generator.detect_file_type(file_path)
        
        if file_type == 'excel':
            excel_info = generator.load_excel_info(file_path)
            return {
                'type': 'excel',
                'sheets': excel_info['sheets'],
                'total_sheets': excel_info['total_sheets'],
                'sheets_with_data': excel_info['sheets_with_data']
            }
        else:
            # For CSV files, just return basic info
            import pandas as pd
            df = pd.read_csv(file_path)
            return {
                'type': 'csv',
                'rows': len(df),
                'columns': len(df.columns),
                'column_names': df.columns.tolist()
            }
    except Exception as e:
        return {'error': str(e)}

@app.route('/')
def index():
    """Main dashboard page"""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and show file information"""
    if 'file' not in request.files:
        flash('No file selected')
        return redirect(url_for('index'))
    
    file = request.files['file']
    if file.filename == '':
        flash('No file selected')
        return redirect(url_for('index'))
    
    if not allowed_file(file.filename):
        flash('Invalid file type. Please upload CSV, XLSX, or XLS files only.')
        return redirect(url_for('index'))
    
    # Save uploaded file
    filename = secure_filename(file.filename)
    unique_filename = f"{uuid.uuid4()}_{filename}"
    file_path = os.path.join(UPLOAD_FOLDER, unique_filename)
    file.save(file_path)
    
    # Get file information
    file_info = get_file_info(file_path)
    
    return render_template('file_info.html', 
                         filename=filename, 
                         file_path=unique_filename,
                         file_info=file_info)

@app.route('/generate', methods=['POST'])
def generate_presentation():
    """Generate PowerPoint presentation from uploaded file"""
    try:
        file_path = request.form.get('file_path')
        sheet_name = request.form.get('sheet_name', None)
        output_filename = request.form.get('output_filename', None)
        
        if not file_path:
            return jsonify({'error': 'No file specified'}), 400
        
        full_file_path = os.path.join(UPLOAD_FOLDER, file_path)
        
        if not os.path.exists(full_file_path):
            return jsonify({'error': 'File not found'}), 404
        
        # Generate presentation
        generator = CSVPPTGenerator()
        
        # Create output filename if not provided
        if not output_filename:
            base_name = os.path.splitext(file_path)[0]
            output_filename = f"{base_name}_presentation.pptx"
        elif not output_filename.endswith('.pptx'):
            output_filename += '.pptx'
        
        output_path = os.path.join(UPLOAD_FOLDER, output_filename)
        
        # Generate the presentation
        result_path = generator.create_presentation_from_csv(
            full_file_path, 
            output_filename=output_path,
            sheet_name=sheet_name if sheet_name else None
        )
        
        return jsonify({
            'success': True,
            'download_url': url_for('download_file', filename=output_filename),
            'message': 'Presentation generated successfully!'
        })
        
    except Exception as e:
        return jsonify({'error': f'Error generating presentation: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    """Download generated presentation"""
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        flash('File not found')
        return redirect(url_for('index'))

@app.route('/api/file-info/<filename>')
def api_file_info(filename):
    """API endpoint to get file information"""
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    if os.path.exists(file_path):
        info = get_file_info(file_path)
        return jsonify(info)
    else:
        return jsonify({'error': 'File not found'}), 404

@app.route('/cleanup')
def cleanup_files():
    """Clean up old uploaded files (for development)"""
    try:
        count = 0
        for filename in os.listdir(UPLOAD_FOLDER):
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            if os.path.isfile(file_path):
                os.remove(file_path)
                count += 1
        
        flash(f'Cleaned up {count} files')
        return redirect(url_for('index'))
    except Exception as e:
        flash(f'Error cleaning up files: {str(e)}')
        return redirect(url_for('index'))

if __name__ == '__main__':
    print("🚀 Starting CSV/Excel-to-PowerPoint Dashboard...")
    print("📊 Upload CSV or Excel files to generate presentations")
    print("🌐 Access the dashboard at: http://localhost:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)
