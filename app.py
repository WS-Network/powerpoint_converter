import os
import traceback
from flask import Flask, render_template, request, redirect, send_from_directory, jsonify
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from pptx.util import Pt
from werkzeug.utils import secure_filename
import requests
import json
from collections import defaultdict
import time
from deep_translator import GoogleTranslator

UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

# Special markers for text separation
SLIDE_MARKER = "[[SLIDE]]"
SHAPE_MARKER = "[[SHAPE]]"
PARA_MARKER = "[[PARA]]"
RUN_MARKER = "[[RUN]]"

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['CONVERTED_FOLDER'] = CONVERTED_FOLDER

def log_error(error, context=""):
    """Helper function to log errors with context"""
    print(f"[Error] {context}: {str(error)}")
    print(f"[Error] Stack trace: {traceback.format_exc()}")

def batch_translate_text(text_blocks, direction):
    """
    Translate multiple text blocks using deep-translator library.
    """
    try:
        if not text_blocks:
            return []

        # Set source and target languages based on direction
        source_lang = 'en' if direction == 'en_to_ar' else 'ar'
        target_lang = 'ar' if direction == 'en_to_ar' else 'en'
        
        print(f"[Translation] Batch translating {len(text_blocks)} blocks from {source_lang} to {target_lang}")
        print(f"[Translation] Direction: {direction}")

        # Initialize translator
        translator = GoogleTranslator(source=source_lang, target=target_lang)
        translated_blocks = []
        
        # Process blocks in batches to avoid potential length limits
        batch_size = 10
        for i in range(0, len(text_blocks), batch_size):
            batch = text_blocks[i:i + batch_size]
            
            # Join the batch with special markers
            combined_text = "\n".join(batch)
            print(f"[Translation] Processing batch {i//batch_size + 1}, text: {combined_text}")
            
            try:
                # Translate the batch
                translated_text = translator.translate(combined_text)
                print(f"[Translation] Raw translated text: {translated_text}")
                
                # Split the translated text back into blocks
                batch_translations = translated_text.split("\n")
                translated_blocks.extend(batch_translations)
                print(f"[Translation] Successfully translated batch {i//batch_size + 1}")
            except Exception as e:
                print(f"[Translation] Error translating batch: {str(e)}")
                translated_blocks.extend(batch)  # Use original text if translation fails
            
            # Add a small delay between batches
            if i + batch_size < len(text_blocks):
                time.sleep(0.5)

        print(f"[Translation] Successfully translated all {len(translated_blocks)} blocks")
        return translated_blocks

    except Exception as e:
        log_error(e, "Batch translation error")
        return text_blocks

def extract_text_content(presentation, slide_indices=None):
    """
    Extract all text content from the presentation with position markers.
    Returns a dictionary mapping positions to text content and a list of all text blocks.
    """
    text_mapping = {}
    text_blocks = []
    current_block = 0

    print("[Extraction] Starting text content extraction")
    
    for slide_num, slide in enumerate(presentation.slides, start=1):
        if slide_indices and slide_num not in slide_indices:
            continue

        def process_shape(shape, shape_index):
            nonlocal current_block
            if shape.has_text_frame:
                for para_idx, paragraph in enumerate(shape.text_frame.paragraphs):
                    for run_idx, run in enumerate(paragraph.runs):
                        if run.text.strip():
                            position = (slide_num, shape_index, para_idx, run_idx)
                            text_mapping[current_block] = position
                            text_blocks.append(run.text)
                            current_block += 1

        for shape_idx, shape in enumerate(slide.shapes):
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                for child_idx, child in enumerate(shape.shapes):
                    process_shape(child, f"{shape_idx}.{child_idx}")
            else:
                process_shape(shape, shape_idx)

    print(f"[Extraction] Extracted {len(text_blocks)} text blocks from presentation")
    return text_mapping, text_blocks

def apply_translated_text(presentation, text_mapping, translated_blocks, direction, slide_indices=None):
    """
    Apply translated text back to the presentation while maintaining formatting.
    """
    print("[Application] Applying translated text to presentation")
    
    for block_idx, translated_text in enumerate(translated_blocks):
        if block_idx not in text_mapping:
            continue

        slide_num, shape_idx, para_idx, run_idx = text_mapping[block_idx]
        
        if slide_indices and slide_num not in slide_indices:
            continue

        try:
            slide = presentation.slides[slide_num - 1]
            
            # Handle nested shape indices
            if isinstance(shape_idx, str) and '.' in shape_idx:
                parent_idx, child_idx = map(int, shape_idx.split('.'))
                shape = slide.shapes[parent_idx].shapes[child_idx]
            else:
                shape = slide.shapes[shape_idx]

            if shape.has_text_frame:
                paragraph = shape.text_frame.paragraphs[para_idx]
                run = paragraph.runs[run_idx]
                
                # Apply text and formatting
                run.text = translated_text
                run.font.name = "Arial"
                rPr = run._r.get_or_add_rPr()
                
                # Set language and RTL formatting
                if direction == 'en_to_ar':
                    paragraph.alignment = PP_ALIGN.RIGHT
                    rtl = parse_xml(r'<a:rtl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">1</a:rtl>')
                    paragraph._p.insert(0, rtl)
                    rPr.set(qn('w:lang'), 'ar-LB')
                else:
                    paragraph.alignment = PP_ALIGN.LEFT
                    rtl_elements = paragraph._p.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}rtl')
                    for rtl_elem in rtl_elements:
                        paragraph._p.remove(rtl_elem)
                    rPr.set(qn('w:lang'), 'en-US')

        except Exception as e:
            log_error(e, f"Error applying translation to slide {slide_num}, shape {shape_idx}")

    print("[Application] Finished applying translated text")

def convert_number_to_arabic(text):
    """Helper function to convert numbers to Arabic and clean up formatting"""
    # Remove extra dots and spaces around numbers
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
            # Skip dots that follow numbers
            if char == '.' and last_was_digit:
                continue
            arabic_text += char
            last_was_digit = False
            
    return arabic_text.strip()

def process_text_frame_format(text_frame, direction):
    """Process text frame formatting only (without translation)"""
    try:
        # Check if this is a footer or header
        is_placeholder = hasattr(text_frame, 'is_placeholder') and text_frame.is_placeholder
        
        if direction == 'en_to_ar':
            # Set RTL property for the entire text frame
            txBody = text_frame._element.get_or_add_txBody()
            bodyPr = txBody.get_or_add_bodyPr()
            bodyPr.set('rtl', '1')
            
            for paragraph in text_frame.paragraphs:
                # Set Arabic-compatible font and RTL for each run
                for run in paragraph.runs:
                    run.font.name = "Traditional Arabic"
                    if not run.font.size:
                        run.font.size = Pt(18)  # Default size if none set
                    
                    # Convert numbers to Arabic numerals if they exist
                    text = run.text
                    if any(char.isdigit() for char in text):
                        run.text = convert_number_to_arabic(text)

                # Don't change alignment for footers/headers to preserve layout
                if not is_placeholder:
                    paragraph.alignment = PP_ALIGN.RIGHT
                    
                    # Handle bullet points and numbering
                    if hasattr(paragraph._pPr, 'numPr') and paragraph._pPr.numPr is not None:
                        pPr = paragraph._element.get_or_add_pPr()
                        if pPr.numPr is not None:
                            pPr.set('rtl', '1')
                            # Remove extra formatting from numbering
                            lvl = pPr.get_or_add_numPr().get_or_add_ilvl()
                            lvl.val = 0
                    
                    # Add RTL property to paragraph
                    pPr = paragraph._element.get_or_add_pPr()
                    pPr.set('rtl', '1')
                    
                    # Set Arabic language
                    rPr = run._r.get_or_add_rPr()
                    rPr.set(qn('w:lang'), 'ar-SA')
                    
                    # Add RTL for text direction
                    rtl = parse_xml(r'<a:rtl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">1</a:rtl>')
                    paragraph._p.insert(0, rtl)
        else:  # 'ar_to_en'
            # Remove RTL property from text frame
            txBody = text_frame._element.get_or_add_txBody()
            bodyPr = txBody.get_or_add_bodyPr()
            bodyPr.set('rtl', '0')
            
            for paragraph in text_frame.paragraphs:
                # Set English font
                for run in paragraph.runs:
                    run.font.name = "Arial"
                    if not run.font.size:
                        run.font.size = Pt(12)  # Default size for English
                    
                # Don't change alignment for footers/headers
                if not is_placeholder:
                    paragraph.alignment = PP_ALIGN.LEFT
                    
                    # Handle bullet points and numbering for LTR
                    if paragraph._pPr.numPr is not None:
                        pPr = paragraph._element.get_or_add_pPr()
                        if pPr.numPr is not None:
                            pPr.set('rtl', '0')
                    
                    # Remove RTL property from paragraph
                    pPr = paragraph._element.get_or_add_pPr()
                    pPr.set('rtl', '0')
                    
                    # Set English language
                    rPr = run._r.get_or_add_rPr()
                    rPr.set(qn('w:lang'), 'en-US')
                    
                    # Remove RTL elements
                    rtl_elements = paragraph._p.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}rtl')
                    for rtl_elem in rtl_elements:
                        paragraph._p.remove(rtl_elem)
    except Exception as e:
        log_error(e, "Error processing text frame format")

def process_shape_format(shape, slide_width, direction, in_group=False):
    """Process shape formatting only (without translation)"""
    try:
        # Check if shape is a placeholder (header/footer)
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
            # Don't mirror placeholders (headers/footers)
            if not in_group and not is_placeholder:
                try:
                    shape.left = slide_width - shape.left - shape.width
                except Exception as e:
                    log_error(e, "Error mirroring shape")
            if shape.has_text_frame:
                process_text_frame_format(shape.text_frame, direction)
    except Exception as e:
        log_error(e, "Error processing shape format")

def process_shape_translation(shape, direction, in_group=False):
    """Process shape translation only"""
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for child in shape.shapes:
                process_shape_translation(child, direction, in_group=True)
        else:
            if shape.has_text_frame:
                process_text_frame_translation(shape.text_frame, direction)
    except Exception as e:
        log_error(e, "Error processing shape translation")

def process_text_frame_translation(text_frame, direction):
    """Process text frame translation only"""
    try:
        text_blocks = []
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                if run.text.strip():
                    text_blocks.append(run.text)
        
        if text_blocks:
            translated_blocks = batch_translate_text(text_blocks, direction)
            block_index = 0
            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    if run.text.strip():
                        run.text = translated_blocks[block_index]
                        block_index += 1
    except Exception as e:
        log_error(e, "Error processing text frame translation")

def convert_pptx(input_path, output_path, slide_indices=None, direction='en_to_ar', enable_translation=True):
    try:
        print(f"[Conversion] Starting conversion from {input_path} to {output_path}")
        print(f"[Conversion] Translation enabled: {enable_translation}")
        print(f"[Conversion] Direction: {direction}")
        print(f"[Conversion] Slide indices: {slide_indices}")
        
        prs = Presentation(input_path)
        slide_width = prs.slide_width
        total_slides = len(prs.slides)
        print(f"[Conversion] Total slides: {total_slides}")

        # Validate slide indices
        if slide_indices:
            slide_indices = [i for i in slide_indices if 1 <= i <= total_slides]
            if not slide_indices:
                print("[Conversion] No valid slide indices provided, processing all slides")
                slide_indices = None
            else:
                print(f"[Conversion] Processing slides: {slide_indices}")

        # Step 1: Format Conversion
        print("[Conversion] Step 1: Format Conversion")
        for i, slide in enumerate(prs.slides, start=1):
            if slide_indices and i not in slide_indices:
                print(f"[Conversion] Skipping slide {i} (not in selected indices)")
                continue

            print(f"[Conversion] Processing slide {i}/{total_slides} format")
            for shape in slide.shapes:
                process_shape_format(shape, slide_width, direction)

        # Step 2: Translation (if enabled)
        if enable_translation:
            print("[Conversion] Step 2: Translation")
            
            # Extract all text content
            text_mapping, text_blocks = extract_text_content(prs, slide_indices)
            
            if text_blocks:
                # Translate all text in one batch
                print(f"[Conversion] Translating {len(text_blocks)} text blocks")
                print("[Translation] Text blocks to translate:", text_blocks)
                translated_blocks = batch_translate_text(text_blocks, direction)
                print("[Translation] Translated blocks:", translated_blocks)
                
                # Apply translations back to the presentation
                apply_translated_text(prs, text_mapping, translated_blocks, direction, slide_indices)
            else:
                print("[Conversion] No text content found to translate")
        else:
            print("[Conversion] Translation disabled, skipping translation step")

        print("[Conversion] Saving final presentation...")
        prs.save(output_path)
        print("[Conversion] Conversion completed successfully")
        return 'completed'
    except Exception as e:
        log_error(e, "Error during PowerPoint conversion")
        raise

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    try:
        print("[Request] Received POST request to /convert")
        if 'file' not in request.files:
            return jsonify({'status': 'error', 'message': 'No file uploaded'}), 400

        file = request.files['file']
        if not file.filename:
            return jsonify({'status': 'error', 'message': 'No file selected'}), 400

        output_name = request.form.get('outputName')
        if not output_name:
            return jsonify({'status': 'error', 'message': 'No output name provided'}), 400

        slide_nums_raw = request.form.get('slideNumbers', '')
        conversion_direction = request.form.get('conversionDirection', 'en_to_ar')
        enable_translation = request.form.get('translationToggle') == 'true'
        
        print(f"[Request] Processing file: {file.filename}")
        print(f"[Request] Output name: {output_name}")
        print(f"[Request] Conversion direction: {conversion_direction}")
        print(f"[Request] Translation enabled: {enable_translation}")
        print(f"[Request] Slide numbers raw: {slide_nums_raw}")

        slide_indices = None
        if slide_nums_raw.strip():
            try:
                slide_indices = [int(num.strip()) for num in slide_nums_raw.split(',') if num.strip().isdigit()]
                if not slide_indices:
                    print("[Request] No valid slide numbers found, processing all slides")
                else:
                    print(f"[Request] Processing slides: {slide_indices}")
            except ValueError as e:
                print(f"[Request] Error parsing slide numbers: {e}")
                print("[Request] Processing all slides")

        if not file.filename.endswith('.pptx'):
            return jsonify({'status': 'error', 'message': 'Invalid file type. Please upload a .pptx file'}), 400

        filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        print(f"[File] Saving uploaded file to: {input_path}")
        file.save(input_path)

        output_filename = secure_filename(output_name) + '.pptx'
        output_path = os.path.join(app.config['CONVERTED_FOLDER'], output_filename)
        print(f"[File] Output path: {output_path}")

        status = convert_pptx(
            input_path=input_path,
            output_path=output_path,
            slide_indices=slide_indices,
            direction=conversion_direction,
            enable_translation=enable_translation
        )

        if os.path.exists(input_path):
            print("[File] Cleaning up input file")
            os.remove(input_path)

        download_url = f'/download/{output_filename}'
        print(f"[Response] Returning download URL: {download_url}")
        return jsonify({
            'status': status,
            'download_url': download_url
        })

    except Exception as e:
        log_error(e, "Error during request processing")
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = os.path.join(app.config['CONVERTED_FOLDER'], filename)
        print(f"[Download] Requested file: {file_path}")
        
        if not os.path.exists(file_path):
            print(f"[Download] File not found: {file_path}")
            return jsonify({'status': 'error', 'message': 'File not found'}), 404
            
        print("[Download] Preparing download response")
        response = send_from_directory(
            app.config['CONVERTED_FOLDER'],
            filename,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        
        print("[Download] Setting response headers")
        response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
        response.headers['Pragma'] = 'no-cache'
        response.headers['Expires'] = '0'
        response.headers['Content-Disposition'] = f'attachment; filename="{filename}"'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        
        print("[Download] Download response prepared successfully")
        return response
        
    except Exception as e:
        log_error(e, "Error during file download")
        return jsonify({'status': 'error', 'message': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5005))
    app.run(host='0.0.0.0', port=port, debug=False)
