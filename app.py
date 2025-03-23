import os
import traceback
from flask import Flask, render_template, request, redirect, send_from_directory, jsonify, after_this_request
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from pptx.util import Pt
from werkzeug.utils import secure_filename
import gc
import time

UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted'
MAX_FILE_AGE = 300  # 5 minutes in seconds

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['CONVERTED_FOLDER'] = CONVERTED_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

def cleanup_old_files():
    """Clean up files older than MAX_FILE_AGE seconds"""
    current_time = time.time()
    
    for folder in [UPLOAD_FOLDER, CONVERTED_FOLDER]:
        for filename in os.listdir(folder):
            filepath = os.path.join(folder, filename)
            if os.path.isfile(filepath):
                file_age = current_time - os.path.getmtime(filepath)
                if file_age > MAX_FILE_AGE:
                    try:
                        os.remove(filepath)
                        print(f"[Cleanup] Removed old file: {filepath}")
                    except Exception as e:
                        print(f"[Cleanup] Error removing {filepath}: {e}")

def log_error(error, context=""):
    """Helper function to log errors with context"""
    print(f"[Error] {context}: {str(error)}")
    print(f"[Error] Stack trace: {traceback.format_exc()}")

def convert_number_to_arabic(text):
    """Helper function to convert numbers to Arabic and clean up formatting"""
    text = text.strip()
    if text.endswith('.'):
        text = text[:-1]
    if text.startswith('.'):
        text = text[1:]
        
    arabic_text = ""
    last_was_digit = False
    
    for char in text:
        if char.isdigit():
            arabic_text += chr(ord('٠') + int(char))
            last_was_digit = True
        else:
            if char == '.' and last_was_digit:
                continue
            arabic_text += char
            last_was_digit = False
            
    return arabic_text.strip()

def process_text_frame_format(text_frame, direction):
    """Process text frame formatting only"""
    try:
        is_placeholder = hasattr(text_frame, 'is_placeholder') and text_frame.is_placeholder
        
        if direction == 'en_to_ar':
            txBody = text_frame._element.get_or_add_txBody()
            bodyPr = txBody.get_or_add_bodyPr()
            bodyPr.set('rtl', '1')
            
            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.name = "Traditional Arabic"
                    if not run.font.size:
                        run.font.size = Pt(18)
                    
                    if any(char.isdigit() for char in run.text):
                        run.text = convert_number_to_arabic(run.text)

                if not is_placeholder:
                    paragraph.alignment = PP_ALIGN.RIGHT
                    
                    if hasattr(paragraph._pPr, 'numPr') and paragraph._pPr.numPr is not None:
                        pPr = paragraph._element.get_or_add_pPr()
                        if pPr.numPr is not None:
                            pPr.set('rtl', '1')
                            lvl = pPr.get_or_add_numPr().get_or_add_ilvl()
                            lvl.val = 0
                    
                    pPr = paragraph._element.get_or_add_pPr()
                    pPr.set('rtl', '1')
                    
                    rPr = run._r.get_or_add_rPr()
                    rPr.set(qn('w:lang'), 'ar-SA')
                    
                    rtl = parse_xml(r'<a:rtl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">1</a:rtl>')
                    paragraph._p.insert(0, rtl)
        else:
            txBody = text_frame._element.get_or_add_txBody()
            bodyPr = txBody.get_or_add_bodyPr()
            bodyPr.set('rtl', '0')
            
            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.name = "Arial"
                    if not run.font.size:
                        run.font.size = Pt(12)
                    
                if not is_placeholder:
                    paragraph.alignment = PP_ALIGN.LEFT
                    
                    if paragraph._pPr.numPr is not None:
                        pPr = paragraph._element.get_or_add_pPr()
                        if pPr.numPr is not None:
                            pPr.set('rtl', '0')
                    
                    pPr = paragraph._element.get_or_add_pPr()
                    pPr.set('rtl', '0')
                    
                    rPr = run._r.get_or_add_rPr()
                    rPr.set(qn('w:lang'), 'en-US')
                    
                    rtl_elements = paragraph._p.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}rtl')
                    for rtl_elem in rtl_elements:
                        paragraph._p.remove(rtl_elem)
    except Exception as e:
        log_error(e, "Error processing text frame format")

def process_shape_format(shape, slide_width, direction, in_group=False):
    """Process shape formatting only"""
    try:
        is_placeholder = hasattr(shape, 'is_placeholder') and shape.is_placeholder
        
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            if not in_group and not is_placeholder:
                try:
                    shape.left = slide_width - shape.left - shape.width
                except Exception as e:
                    log_error(e, "Error mirroring group container")
            for child in shape.shapes:
                process_shape_format(child, slide_width, direction, in_group=True)
        else:
            if not in_group and not is_placeholder:
                try:
                    shape.left = slide_width - shape.left - shape.width
                except Exception as e:
                    log_error(e, "Error mirroring shape")
            if shape.has_text_frame:
                process_text_frame_format(shape.text_frame, direction)
    except Exception as e:
        log_error(e, "Error processing shape format")

def convert_pptx(input_path, output_path, slide_indices=None, direction='en_to_ar'):
    try:
        print(f"[Conversion] Starting conversion from {input_path} to {output_path}")
        
        # Load presentation
        prs = Presentation(input_path)
        slide_width = prs.slide_width
        total_slides = len(prs.slides)
        
        if slide_indices:
            slide_indices = [i for i in slide_indices if 1 <= i <= total_slides]
        
        # Process slides
        for i, slide in enumerate(prs.slides, start=1):
            if slide_indices and i not in slide_indices:
                continue
            for shape in slide.shapes:
                process_shape_format(shape, slide_width, direction)
        
        # Save and cleanup
        prs.save(output_path)
        del prs
        gc.collect()  # Force garbage collection
        
        return 'completed'
    except Exception as e:
        log_error(e, "Error during PowerPoint conversion")
        raise
    finally:
        # Ensure we clean up the input file
        try:
            if os.path.exists(input_path):
                os.remove(input_path)
        except Exception as e:
            log_error(e, "Error cleaning up input file")

@app.route('/', methods=['GET'])
def index():
    cleanup_old_files()  # Clean up old files on page load
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    input_path = None
    output_path = None
    
    try:
        if 'file' not in request.files:
            return jsonify({'status': 'error', 'message': 'No file uploaded'}), 400

        file = request.files['file']
        if not file.filename:
            return jsonify({'status': 'error', 'message': 'No file selected'}), 400

        if not file.filename.endswith('.pptx'):
            return jsonify({'status': 'error', 'message': 'Invalid file type. Please upload a .pptx file'}), 400

        # Get parameters
        output_name = request.form.get('outputName')
        if not output_name:
            return jsonify({'status': 'error', 'message': 'No output name provided'}), 400

        slide_nums_raw = request.form.get('slideNumbers', '')
        conversion_direction = request.form.get('conversionDirection', 'en_to_ar')

        # Process slide numbers
        slide_indices = None
        if slide_nums_raw.strip():
            try:
                slide_indices = [int(num.strip()) for num in slide_nums_raw.split(',') if num.strip().isdigit()]
            except ValueError:
                slide_indices = None

        # Save and process file
        filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        output_filename = secure_filename(output_name) + '.pptx'
        output_path = os.path.join(app.config['CONVERTED_FOLDER'], output_filename)

        file.save(input_path)
        
        status = convert_pptx(
            input_path=input_path,
            output_path=output_path,
            slide_indices=slide_indices,
            direction=conversion_direction
        )

        # Set up automatic cleanup after sending
        @after_this_request
        def cleanup(response):
            try:
                if os.path.exists(output_path):
                    os.remove(output_path)
            except Exception as e:
                log_error(e, "Error in cleanup after request")
            return response

        return jsonify({
            'status': status,
            'download_url': f'/download/{output_filename}'
        })

    except Exception as e:
        # Clean up files in case of error
        try:
            if input_path and os.path.exists(input_path):
                os.remove(input_path)
            if output_path and os.path.exists(output_path):
                os.remove(output_path)
        except Exception as cleanup_error:
            log_error(cleanup_error, "Error cleaning up files after error")
            
        log_error(e, "Error during request processing")
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = os.path.join(app.config['CONVERTED_FOLDER'], filename)
        
        if not os.path.exists(file_path):
            return jsonify({'status': 'error', 'message': 'File not found'}), 404

        @after_this_request
        def cleanup(response):
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except Exception as e:
                log_error(e, "Error cleaning up after download")
            return response

        return send_from_directory(
            app.config['CONVERTED_FOLDER'],
            filename,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )

    except Exception as e:
        log_error(e, "Error during file download")
        return jsonify({'status': 'error', 'message': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5005))
    app.run(host='0.0.0.0', port=port, debug=False)
