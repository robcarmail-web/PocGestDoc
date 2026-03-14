"""
Flask application for DOCX Template Injection POC.
Includes only the Atto implementation.
"""
import sys
import os
from io import BytesIO
import tempfile
import shutil

sys.path.insert(0, 'modules')

from flask import Flask, render_template, request, send_file, jsonify, redirect, url_for
from werkzeug.utils import secure_filename

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
    # Nuovi default per Atto
    'ResponsabileNome': 'Dott.ssa Laura Rosa',
    'DirettoreNome': 'Dott. Mario Rossi',
}

@app.route('/')
def index():
    """Redirect to Atto page."""
    return redirect(url_for('atto'))

@app.route('/atto')
def atto():
    """Serve Atto page e assicura TestoAtto.docx in output."""
    src = os.path.join('templates', 'TestoAtto.docx')
    dst = os.path.join('output', 'TestoAtto.docx')
    if os.path.exists(src) and not os.path.exists(dst):
        shutil.copy2(src, dst)
    return render_template('atto.html', defaults=DEFAULT_VALUES)

@app.route('/api/upload-testo-atto', methods=['POST'])
def upload_testo_atto():
    """Sovrascrive output/TestoAtto.docx con il file caricato."""
    if 'file' not in request.files:
        return jsonify({'error': 'Nessun file selezionato'}), 400
    
    file = request.files['file']
    if not file.filename.endswith('.docx'):
        return jsonify({'error': 'Il file deve essere .docx'}), 400
        
    try:
        dst = os.path.join('output', 'TestoAtto.docx')
        file.save(dst)
        return jsonify({'success': True, 'message': 'File caricato e sovrascritto con successo!'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/genera', methods=['POST'])
def genera_documento():
    """
    Generate DOCX document for Atto.
    """
    try:
        data = request.form.to_dict()

        # Extract simple placeholders
        simple_data = {}
        for key in DEFAULT_VALUES.keys():
            if key in data:
                simple_data[key] = data[key]

        rich_content = {}
        template_file = 'template/ASL_Template_Atto.docx'

        if os.path.exists(os.path.join('output', 'TestoAtto.docx')):
            # Creiamo un file temporaneo copiato da TestoAtto.docx
            # per farlo cancellare dalla logica standard senza rompere il file originale WebDAV
            temp_file = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_testo_atto_gen.docx')
            shutil.copy2(os.path.join('output', 'TestoAtto.docx'), temp_file)
            rich_content['P_testo_obj'] = temp_file
        else:
            return jsonify({'error': 'Il file output/TestoAtto.docx non esiste sul server. Ricarica la pagina iniziale.'}), 400

        # Generate DOCX
        from docx_injector import DocxInjector
        injector = DocxInjector(template_file)
        docx_bytes = injector.inject_placeholders(simple_data, rich_content)

        # Cleanup temp files
        for var_path in rich_content.values():
            if os.path.exists(var_path):
                os.remove(var_path)

        # Save and return
        output_file = 'output/atto_output.docx'
        with open(output_file, 'wb') as f:
            f.write(docx_bytes)

        return send_file(
            BytesIO(docx_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='atto.docx'
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/genera-pdf', methods=['POST'])
def genera_pdf():
    """Generate PDF from generated DOCX for Atto using docx2pdf (Native Word)."""
    try:
        data = request.form.to_dict()
        simple_data = {k: v for k, v in data.items() if k in DEFAULT_VALUES}
        rich_content = {}
        
        template_file = 'template/ASL_Template_Atto.docx'

        if os.path.exists(os.path.join('output', 'TestoAtto.docx')):
            temp_file = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_testo_atto_gen_pdf.docx')
            shutil.copy2(os.path.join('output', 'TestoAtto.docx'), temp_file)
            rich_content['P_testo_obj'] = temp_file
        else:
            return jsonify({'error': 'Il file output/TestoAtto.docx non esiste sul server.'}), 400

        from docx_injector import DocxInjector
        injector = DocxInjector(template_file)
        docx_bytes = injector.inject_placeholders(simple_data, rich_content)

        # Cleanup temp files
        for var_path in rich_content.values():
            if os.path.exists(var_path):
                os.remove(var_path)

        # Save intermediate DOCX
        temp_docx = 'output/temp_atto_for_pdf.docx'
        temp_pdf = 'output/temp_atto_for_pdf.pdf'
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
                download_name='atto.pdf'
            )
        else:
            return jsonify({'error': 'PDF conversion failed (docx2pdf)'}), 500

    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5000)
