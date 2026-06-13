"""
Flask application for DOCX Template Injection POC.
Includes only the Atto implementation.
"""
import sys
import os
import logging
import subprocess
from io import BytesIO
import shutil

sys.path.insert(0, 'modules')

from flask import Flask, render_template, request, send_file, jsonify, redirect, url_for

# Configurazione centralizzata (variabili d'ambiente con default)
def _env(key, default):
    return os.environ.get(key, default)

TEMPLATE_ATTO = _env('TEMPLATE_ATTO', 'template/ASL_Template_Atto.docx')
OUTPUT_DIR = _env('OUTPUT_DIR', 'output')
UPLOAD_DIR = _env('UPLOAD_FOLDER', 'uploads')
TESTO_ATTO_FILENAME = 'TestoAtto.docx'
ATTO_OUTPUT_FILENAME = 'atto_output.docx'
MAX_UPLOAD_MB = int(_env('MAX_UPLOAD_MB', '50'))
FLASK_DEBUG = _env('FLASK_DEBUG', '0').lower() in ('1', 'true', 'yes')
FLASK_PORT = int(_env('FLASK_PORT', '5000'))
WEBDAV_BASE_URL = _env('WEBDAV_BASE_URL', 'http://localhost:8080').rstrip('/')
LIBREOFFICE_BIN = _env('LIBREOFFICE_BIN', 'soffice')
PDF_CONVERT_TIMEOUT = int(_env('PDF_CONVERT_TIMEOUT', '120'))


def _pdf_backend():
    explicit = _env('PDF_BACKEND', '').lower()
    if explicit in ('libreoffice', 'word'):
        return explicit
    return 'word' if sys.platform == 'win32' else 'libreoffice'

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_DIR
app.config['MAX_CONTENT_LENGTH'] = MAX_UPLOAD_MB * 1024 * 1024
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

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
    'ResponsabileNome': 'Dott.ssa Laura Rosa',
    'DirettoreNome': 'Dott. Mario Rossi',
}


def _build_atto_docx(form_data):
    """
    Genera i byte del DOCX dell'Atto a partire dai dati del form.
    Solleva ValueError con messaggio utente in caso di configurazione mancante.
    Solleva altre eccezioni in caso di errore di generazione.
    """
    if not os.path.exists(TEMPLATE_ATTO):
        raise ValueError(
            f'Template non trovato: {TEMPLATE_ATTO}. Verificare la configurazione del server.'
        )
    testo_path = os.path.join(OUTPUT_DIR, TESTO_ATTO_FILENAME)
    if not os.path.exists(testo_path):
        raise ValueError(
            'Il file TestoAtto.docx non esiste sul server. Ricarica la pagina iniziale o carica un file.'
        )

    simple_data = {k: form_data[k] for k in DEFAULT_VALUES if k in form_data}
    temp_file = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_testo_atto_gen.docx')
    shutil.copy2(testo_path, temp_file)
    rich_content = {'P_testo_obj': temp_file}

    try:
        from docx_injector import DocxInjector
        injector = DocxInjector(TEMPLATE_ATTO)
        docx_bytes = injector.inject_placeholders(simple_data, rich_content)
        return docx_bytes
    finally:
        for path in rich_content.values():
            if os.path.exists(path):
                try:
                    os.remove(path)
                except OSError:
                    logger.warning('Impossibile rimuovere file temporaneo: %s', path)


def _convert_docx_to_pdf_libreoffice(docx_path, pdf_path):
    out_dir = os.path.dirname(os.path.abspath(pdf_path))
    cmd = [
        LIBREOFFICE_BIN,
        '--headless',
        '--nologo',
        '--nofirststartwizard',
        '--convert-to', 'pdf',
        '--outdir', out_dir,
        os.path.abspath(docx_path),
    ]
    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        timeout=PDF_CONVERT_TIMEOUT,
        check=False,
    )
    if result.returncode != 0:
        detail = (result.stderr or result.stdout or '').strip()
        raise RuntimeError(
            f'LibreOffice non ha convertito il documento ({LIBREOFFICE_BIN}). {detail}'.strip()
        )

    generated_pdf = os.path.join(
        out_dir,
        os.path.splitext(os.path.basename(docx_path))[0] + '.pdf',
    )
    if not os.path.exists(generated_pdf):
        raise RuntimeError(
            f'LibreOffice non ha prodotto il PDF atteso: {generated_pdf}'
        )
    if os.path.abspath(generated_pdf) != os.path.abspath(pdf_path):
        shutil.move(generated_pdf, pdf_path)


def _convert_docx_to_pdf_word(docx_path, pdf_path):
    import pythoncom
    from docx2pdf import convert

    pythoncom.CoInitialize()
    try:
        convert(docx_path, pdf_path)
    finally:
        pythoncom.CoUninitialize()


def _convert_docx_to_pdf(docx_path, pdf_path):
    backend = _pdf_backend()
    if backend == 'word':
        _convert_docx_to_pdf_word(docx_path, pdf_path)
        return

    _convert_docx_to_pdf_libreoffice(docx_path, pdf_path)


@app.route('/')
def index():
    """Redirect to Atto page."""
    return redirect(url_for('atto'))


@app.route('/atto')
def atto():
    """Serve Atto page e assicura TestoAtto.docx in output."""
    src = os.path.join('templates', TESTO_ATTO_FILENAME)
    dst = os.path.join(OUTPUT_DIR, TESTO_ATTO_FILENAME)
    if os.path.exists(src) and not os.path.exists(dst):
        shutil.copy2(src, dst)
    return render_template(
        'atto.html',
        defaults=DEFAULT_VALUES,
        webdav_base_url=WEBDAV_BASE_URL,
    )


@app.route('/api/upload-testo-atto', methods=['POST'])
def upload_testo_atto():
    """Sovrascrive output/TestoAtto.docx con il file caricato."""
    if 'file' not in request.files:
        return jsonify({'error': 'Nessun file selezionato'}), 400

    file = request.files['file']
    if not file or not file.filename or not file.filename.lower().endswith('.docx'):
        return jsonify({'error': 'Il file deve essere .docx'}), 400

    try:
        dst = os.path.join(OUTPUT_DIR, TESTO_ATTO_FILENAME)
        file.save(dst)
        return jsonify({'success': True, 'message': 'File caricato e sovrascritto con successo!'})
    except Exception as e:
        logger.exception('Upload TestoAtto fallito')
        return jsonify({'error': 'Errore durante il salvataggio del file.'}), 500


@app.route('/api/genera', methods=['POST'])
def genera_documento():
    """Genera il DOCX finale per l'Atto."""
    try:
        data = request.form.to_dict()
        docx_bytes = _build_atto_docx(data)

        output_file = os.path.join(OUTPUT_DIR, ATTO_OUTPUT_FILENAME)
        with open(output_file, 'wb') as f:
            f.write(docx_bytes)

        return send_file(
            BytesIO(docx_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='atto.docx'
        )
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    except Exception:
        logger.exception('Generazione DOCX fallita')
        return jsonify({'error': 'Errore durante la generazione del documento.'}), 500


@app.route('/api/genera-pdf', methods=['POST'])
def genera_pdf():
    """Genera PDF dall'Atto (LibreOffice su Linux, Word/COM su Windows)."""
    temp_docx = None
    temp_pdf = None
    try:
        data = request.form.to_dict()
        docx_bytes = _build_atto_docx(data)

        temp_docx = os.path.join(OUTPUT_DIR, 'temp_atto_for_pdf.docx')
        temp_pdf = os.path.join(OUTPUT_DIR, 'temp_atto_for_pdf.pdf')
        with open(temp_docx, 'wb') as f:
            f.write(docx_bytes)

        _convert_docx_to_pdf(temp_docx, temp_pdf)

        if not os.path.exists(temp_pdf):
            backend = _pdf_backend()
            if backend == 'word':
                message = 'Conversione in PDF non riuscita. Verificare che Word sia installato e disponibile.'
            else:
                message = (
                    f'Conversione in PDF non riuscita. Verificare che LibreOffice sia installato '
                    f'({LIBREOFFICE_BIN}).'
                )
            return jsonify({'error': message}), 500

        with open(temp_pdf, 'rb') as f:
            pdf_bytes = f.read()

        return send_file(
            BytesIO(pdf_bytes),
            mimetype='application/pdf',
            as_attachment=True,
            download_name='atto.pdf'
        )
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    except Exception:
        logger.exception('Generazione PDF fallita')
        return jsonify({'error': 'Errore durante la generazione del PDF.'}), 500
    finally:
        for path in (temp_docx, temp_pdf):
            if path and os.path.exists(path):
                try:
                    os.remove(path)
                except OSError:
                    logger.warning('Impossibile rimuovere file temporaneo: %s', path)


if __name__ == '__main__':
    app.run(debug=FLASK_DEBUG, port=FLASK_PORT)
