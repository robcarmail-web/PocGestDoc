"""
Simple POC - Una sola delibera con un solo placeholder.
Interfaccia minimale: 1 upload + 1 editor.
Output: DOCX con formattazione preservata.
"""
import sys
import json
import os
from pathlib import Path
from io import BytesIO

sys.path.insert(0, 'modules')

from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

from docx_injector import DocxInjector
from tiptap_to_docx import TiptapToDocx

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('output', exist_ok=True)


@app.route('/')
def index():
    """Serve main page."""
    return render_template('simple.html')


@app.route('/api/genera', methods=['POST'])
def genera():
    """Genera DOCX - supporta DOCX upload o TipTap editor."""
    try:
        data = request.form.to_dict()
        paragraphs = []

        injector = DocxInjector('template/ASL_Template_Simple.docx')

        # Modalità 1: DOCX upload - copia XML direttamente
        if 'docx_file' in request.files and request.files['docx_file']:
            file = request.files['docx_file']
            temp_file = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
            file.save(temp_file)
            try:
                docx_bytes = injector.inject_from_docx(temp_file, 'P_testo_obj')
            finally:
                os.remove(temp_file)

        # Modalità 2: Editor web
        elif 'editor_content' in data and data['editor_content']:
            tiptap_data = json.loads(data['editor_content'])
            converter = TiptapToDocx()
            paragraphs, _ = converter.convert(json.dumps(tiptap_data))
            docx_bytes = injector.inject_from_paragraphs(paragraphs, 'P_testo_obj')

        else:
            return jsonify({'error': 'Nessun contenuto fornito'}), 400

        output = BytesIO(docx_bytes)
        output.seek(0)

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='delibera.docx'
        )

    except Exception as e:
        print(f"[ERROR] {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


def _save_uploaded_file(file_obj):
    """Save uploaded file and return path."""
    temp_file = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file_obj.filename))
    file_obj.save(temp_file)
    return temp_file


if __name__ == '__main__':
    app.run(debug=True, port=5001)
