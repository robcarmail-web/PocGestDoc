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
            if has_docx_1:
                temp_file = _save_uploaded_file(request.files['docx_istruttoria'])
                rich_content['I_testo_obj'] = temp_file

            if has_docx_2a:
                temp_file = _save_uploaded_file(request.files['docx_proposta'])
                rich_content['P_testo_obj'] = temp_file

            if has_docx_2b:
                temp_file = _save_uploaded_file(request.files['docx_delibera'])
                rich_content['P_testo_obj_2'] = temp_file

        # Modalita 2/3: TipTap JSON
        # Per questo POC, manteniamo il fallback legacy a ODTInjector per l'editor web
        # (Idealmente l'editor TipTap dovrebbe inviare DOCX o HTML e il server lo converte a DOCX)
        elif has_tiptap_1 or has_tiptap_2a or has_tiptap_2b:
            return jsonify({'error': 'TipTap a DOCX non ancora supportato in questa iterazione. Usa DOCX Upload.'}), 400

        # Generate DOCX
        from docx_injector import DocxInjector
        injector = DocxInjector('template/ASL_Template_Delibera.docx')
        docx_bytes = injector.inject_placeholders(simple_data, rich_content if rich_content else None)

        # Cleanup temp files
        for var_path in rich_content.values():
            if os.path.exists(var_path):
                os.remove(var_path)

        # Save and return
        output_file = 'output/delibera_output.docx'
        with open(output_file, 'wb') as f:
            f.write(docx_bytes)

        return send_file(
            BytesIO(docx_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='delibera.docx'
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/genera-pdf', methods=['POST'])
def genera_pdf():
    """Generate PDF from generated DOCX using docx2pdf (Native Word)."""
    try:
        # First generate the DOCX document normally
        # We reuse the logic without returning the flask response yet
        data = request.form.to_dict()
        simple_data = {k: v for k, v in data.items() if k in DEFAULT_VALUES}
        rich_content = {}
        
        has_docx_1 = 'docx_istruttoria' in request.files and request.files['docx_istruttoria']
        has_docx_2a = 'docx_proposta' in request.files and request.files['docx_proposta']
        has_docx_2b = 'docx_delibera' in request.files and request.files['docx_delibera']

        if has_docx_1:
            rich_content['I_testo_obj'] = _save_uploaded_file(request.files['docx_istruttoria'])
        if has_docx_2a:
            rich_content['P_testo_obj'] = _save_uploaded_file(request.files['docx_proposta'])
        if has_docx_2b:
            rich_content['P_testo_obj_2'] = _save_uploaded_file(request.files['docx_delibera'])

        from docx_injector import DocxInjector
        injector = DocxInjector('template/ASL_Template_Delibera.docx')
        docx_bytes = injector.inject_placeholders(simple_data, rich_content if rich_content else None)

        # Cleanup temp upload files
        for var_path in rich_content.values():
            if os.path.exists(var_path):
                os.remove(var_path)

        # Save intermediate DOCX
        temp_docx = 'output/temp_delibera_for_pdf.docx'
        temp_pdf = 'output/temp_delibera_for_pdf.pdf'
        with open(temp_docx, 'wb') as f:
            f.write(docx_bytes)

        # Convert to PDF using docx2pdf
        from docx2pdf import convert
        import pythoncom
        # Initialize COM for the current thread (required for Flask/threaded apps)
        pythoncom.CoInitialize()
        try:
            convert(temp_docx, temp_pdf)
        finally:
             pythoncom.CoUninitialize()

        if os.path.exists(temp_pdf):
            with open(temp_pdf, 'rb') as f:
                pdf_bytes = f.read()

            os.remove(temp_docx)
            os.remove(temp_pdf)

            return send_file(
                BytesIO(pdf_bytes),
                mimetype='application/pdf',
                as_attachment=True,
                download_name='delibera.pdf'
            )
        else:
            return jsonify({'error': 'PDF conversion failed (docx2pdf)'}), 500

    except Exception as e:
        return jsonify({'error': str(e)}), 500


def _save_uploaded_file(file_obj):
    """Save uploaded file and return path."""
    temp_file = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file_obj.filename))
    file_obj.save(temp_file)
    return temp_file


if __name__ == '__main__':
    app.run(debug=True, port=5000)
