"""
Flask application for ODT Template Injection POC.
Supports three modes: direct DOCX upload, DOCX + editor, editor-only.
"""
import sys
import json
import os
from pathlib import Path
from io import BytesIO
import tempfile

sys.path.insert(0, 'modules')

from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

from odt_injector import ODTInjector
from docx_to_odt import DocxToODT
from html_to_tiptap import HTMLToTiptap
from tiptap_to_odt import TiptapToODT

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('output', exist_ok=True)

# Default placeholder values
DEFAULT_VALUES = {
    'numeroproposta': '2024/001',
    'dataproposta': '15/01/2024',
    'oggetto': 'Approvazione Piano Operativo Annuale 2024',
    'ufficioproponente': 'U.O.C. Programmazione e Controllo',
    'CentroDICosto': 'CC-001',
    'EstensoreNome': 'Dott. Mario Rossi',
    'RUPNome': 'Dott.ssa Anna Bianchi',
    'DirigenteNome': 'Dott. Giuseppe Verdi',
    'DirSanNome': 'Dott. Carlo Neri',
    'DirSanAzione': 'FAVOREVOLE',
    'DirSanData': '15/01/2024',
    'DirAmmNome': 'Dott.ssa Laura Gialli',
    'DirAmmAzione': 'FAVOREVOLE',
    'DirAmmData': '15/01/2024',
    'DirGenNome': 'Prof. Roberto Blu',
    'SostitutoDelDirettoreGenerale': '',
}


@app.route('/')
def index():
    """Serve main page."""
    return render_template('index.html', defaults=DEFAULT_VALUES)


@app.route('/editor')
def editor():
    """Serve advanced TipTap editor."""
    return render_template('editor.html')


@app.route('/api/docx-to-tiptap', methods=['POST'])
def docx_to_tiptap():
    """Modalita 2: Convert uploaded DOCX to TipTap JSON."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    if not file.filename.endswith('.docx'):
        return jsonify({'error': 'File must be .docx'}), 400

    try:
        # Save temp file
        temp_file = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
        file.save(temp_file)

        # Convert DOCX to ODT fragments
        converter = DocxToODT()
        fragments, styles = converter.convert_file(temp_file)

        # Convert first fragment to TipTap (simplistic approach)
        # In a real app, this would be more sophisticated
        tiptap_data = {
            "type": "doc",
            "content": [
                {
                    "type": "paragraph",
                    "content": [
                        {
                            "type": "text",
                            "text": f"Content from {len(fragments)} paragraphs in DOCX"
                        }
                    ]
                }
            ]
        }

        os.remove(temp_file)
        return jsonify({'tiptap': tiptap_data})

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/genera', methods=['POST'])
def genera_documento():
    """
    Generate ODT document.
    Accepts: simple placeholders + one of:
      - docx files (Modalita 1)
      - tiptap JSON (Modalita 2/3)
    """
    try:
        data = request.form.to_dict()

        # Extract simple placeholders
        simple_data = {}
        rich_content = {}

        for key in DEFAULT_VALUES.keys():
            if key in data:
                simple_data[key] = data[key]

        # Handle rich content - three modes
        has_tiptap_1 = 'tiptap_istruttoria' in data and data['tiptap_istruttoria']
        has_tiptap_2a = 'tiptap_proposta' in data and data['tiptap_proposta']
        has_tiptap_2b = 'tiptap_delibera' in data and data['tiptap_delibera']

        has_docx_1 = 'docx_istruttoria' in request.files and request.files['docx_istruttoria']
        has_docx_2a = 'docx_proposta' in request.files and request.files['docx_proposta']
        has_docx_2b = 'docx_delibera' in request.files and request.files['docx_delibera']

        # Modalita 1: Direct DOCX upload
        if has_docx_1 or has_docx_2a or has_docx_2b:
            docx_converter = DocxToODT()

            if has_docx_1:
                temp_file = _save_uploaded_file(request.files['docx_istruttoria'])
                fragments, styles = docx_converter.convert_file(temp_file)
                rich_content['I_testo_obj'] = (fragments, styles)
                os.remove(temp_file)

            if has_docx_2a:
                temp_file = _save_uploaded_file(request.files['docx_proposta'])
                fragments, styles = docx_converter.convert_file(temp_file)
                rich_content['P_testo_obj'] = (fragments, styles)
                os.remove(temp_file)

            if has_docx_2b:
                temp_file = _save_uploaded_file(request.files['docx_delibera'])
                fragments, styles = docx_converter.convert_file(temp_file)
                rich_content['P_testo_obj_2'] = (fragments, styles)
                os.remove(temp_file)

        # Modalita 2/3: TipTap JSON
        elif has_tiptap_1 or has_tiptap_2a or has_tiptap_2b:
            tiptap_converter = TiptapToODT()

            if has_tiptap_1:
                fragments, styles = tiptap_converter.convert(data['tiptap_istruttoria'])
                rich_content['I_testo_obj'] = (fragments, styles)

            if has_tiptap_2a:
                fragments, styles = tiptap_converter.convert(data['tiptap_proposta'])
                rich_content['P_testo_obj'] = (fragments, styles)

            if has_tiptap_2b:
                fragments, styles = tiptap_converter.convert(data['tiptap_delibera'])
                rich_content['P_testo_obj_2'] = (fragments, styles)

        # Generate ODT
        injector = ODTInjector('template/ASL_Template_Delibera.odt')
        odt_bytes = injector.inject_placeholders(simple_data, rich_content if rich_content else None)

        # Save and return
        output_file = 'output/delibera_output.odt'
        with open(output_file, 'wb') as f:
            f.write(odt_bytes)

        return send_file(
            BytesIO(odt_bytes),
            mimetype='application/vnd.oasis.opendocument.text',
            as_attachment=True,
            download_name='delibera.odt'
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/genera-pdf', methods=['POST'])
def genera_pdf():
    """Generate PDF from generated ODT (optional)."""
    try:
        # First generate ODT
        response = genera_documento()

        # Try to convert to PDF using soffice if available
        import subprocess

        temp_odt = 'output/temp_delibera.odt'
        temp_pdf = 'output/temp_delibera.pdf'

        # Save ODT
        data = request.form.to_dict()
        simple_data = {k: v for k, v in data.items() if k in DEFAULT_VALUES}
        injector = ODTInjector('template/ASL_Template_Delibera.odt')
        odt_bytes = injector.inject_placeholders(simple_data)

        with open(temp_odt, 'wb') as f:
            f.write(odt_bytes)

        # Convert to PDF
        result = subprocess.run(
            ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', 'output', temp_odt],
            capture_output=True,
            timeout=30
        )

        if os.path.exists(temp_pdf):
            with open(temp_pdf, 'rb') as f:
                pdf_bytes = f.read()

            os.remove(temp_odt)
            os.remove(temp_pdf)

            return send_file(
                BytesIO(pdf_bytes),
                mimetype='application/pdf',
                as_attachment=True,
                download_name='delibera.pdf'
            )
        else:
            return jsonify({'error': 'PDF conversion failed'}), 500

    except FileNotFoundError:
        return jsonify({'error': 'LibreOffice (soffice) not found. PDF generation unavailable.'}), 501
    except Exception as e:
        return jsonify({'error': str(e)}), 500


def _save_uploaded_file(file_obj):
    """Save uploaded file and return path."""
    temp_file = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file_obj.filename))
    file_obj.save(temp_file)
    return temp_file


if __name__ == '__main__':
    app.run(debug=True, port=5000)
