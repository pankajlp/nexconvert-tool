import os
import sys
import uuid
import time
import threading
import webbrowser
import pythoncom
import win32com.client
from flask import Flask, render_template, request, jsonify, send_file
from pdf2docx import Converter

app = Flask(__name__)

# Configure upload and converted directories inside the project folder
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
CONVERTED_FOLDER = os.path.join(BASE_DIR, 'converted')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['CONVERTED_FOLDER'] = CONVERTED_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB upload limit

# Store metadata of converted files for the download endpoint
# Structure: { file_id: { "filename": "example.docx", "path": "/path/to/example.docx" } }
conversion_registry = {}

def clean_old_files():
    """Periodically clean files older than 30 minutes in upload and converted directories."""
    while True:
        try:
            current_time = time.time()
            for folder in [UPLOAD_FOLDER, CONVERTED_FOLDER]:
                for filename in os.listdir(folder):
                    file_path = os.path.join(folder, filename)
                    if os.path.isfile(file_path):
                        # If file is older than 30 minutes, delete it
                        if current_time - os.path.getmtime(file_path) > 1800:
                            os.remove(file_path)
        except Exception as e:
            print(f"Error cleaning files: {e}", file=sys.stderr)
        time.sleep(600)  # Run cleanup every 10 minutes

# Start the cleanup daemon thread
cleanup_thread = threading.Thread(target=clean_old_files, daemon=True)
cleanup_thread.start()

def convert_pdf_to_docx_task(pdf_path, docx_path):
    """Convert PDF to DOCX using pdf2docx library."""
    cv = Converter(pdf_path)
    cv.convert(docx_path)
    cv.close()

def convert_docx_to_pdf_task(docx_path, pdf_path):
    """Convert DOCX to PDF using win32com client and Microsoft Word."""
    # CoInitialize must be called when using COM in multi-threaded environments (like Flask request threads)
    pythoncom.CoInitialize()
    word = None
    try:
        # Launch Microsoft Word in a headless/background state
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        
        # Word requires absolute paths
        abs_docx = os.path.abspath(docx_path)
        abs_pdf = os.path.abspath(pdf_path)
        
        # Open and Export
        doc = word.Documents.Open(abs_docx)
        # wdFormatPDF = 17
        doc.SaveAs(abs_pdf, FileFormat=17)
        doc.Close()
    finally:
        if word:
            word.Quit()
        pythoncom.CoUninitialize()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({"error": "No file part in the request"}), 400
    
    file = request.files['file']
    direction = request.form.get('direction', '')  # 'pdf2docx' or 'docx2pdf'
    
    if file.filename == '':
        return jsonify({"error": "No file selected"}), 400
    
    if direction not in ['pdf2docx', 'docx2pdf']:
        return jsonify({"error": "Invalid conversion direction specified"}), 400
    
    filename = file.filename
    name, ext = os.path.splitext(filename)
    
    # Perform input validation on file extensions
    if direction == 'pdf2docx' and ext.lower() != '.pdf':
        return jsonify({"error": "Please upload a valid PDF file for PDF to DOCX conversion"}), 400
    elif direction == 'docx2pdf' and ext.lower() not in ['.docx', '.doc']:
        return jsonify({"error": "Please upload a valid Word document (.docx/.doc) for DOCX to PDF conversion"}), 400
    
    # Save input file with unique identifier
    unique_id = str(uuid.uuid4())
    input_filename = f"{unique_id}{ext}"
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
    file.save(input_path)
    
    # Setup output paths
    output_ext = '.docx' if direction == 'pdf2docx' else '.pdf'
    output_filename = f"{name}{output_ext}"
    safe_output_filename = f"{unique_id}{output_ext}"
    output_path = os.path.join(app.config['CONVERTED_FOLDER'], safe_output_filename)
    
    try:
        if direction == 'pdf2docx':
            convert_pdf_to_docx_task(input_path, output_path)
        else:
            convert_docx_to_pdf_task(input_path, output_path)
            
        # Register the successful conversion
        conversion_registry[unique_id] = {
            "filename": output_filename,
            "path": output_path
        }
        
        return jsonify({
            "success": True,
            "fileId": unique_id,
            "filename": output_filename,
            "message": "Conversion successful!"
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({
            "success": False,
            "error": f"Conversion failed: {str(e)}"
        }), 500

@app.route('/api/download/<file_id>', methods=['GET'])
def download(file_id):
    if file_id not in conversion_registry:
        return jsonify({"error": "File not found or link expired"}), 404
    
    meta = conversion_registry[file_id]
    file_path = meta['path']
    filename = meta['filename']
    
    if not os.path.exists(file_path):
        return jsonify({"error": "File does not exist on disk"}), 404
        
    return send_file(file_path, as_attachment=True, download_name=filename)

def open_browser():
    """Wait briefly for server spin-up and open the application page."""
    time.sleep(1.5)
    webbrowser.open("http://127.0.0.1:5000")

if __name__ == '__main__':
    # Start browser opener daemon thread
    threading.Thread(target=open_browser, daemon=True).start()
    
    # Run the local Flask server
    app.run(host='127.0.0.1', port=5000, debug=False)
