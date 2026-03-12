# POC - Injection Template Delibera ASL

Proof of Concept per la gestione dei **placeholder di testo ricco** in un template ODT istituzionale. Il sistema supporta tre modalità diverse di alimentazione dei dati.

## Panoramica

Questo POC dimostra come iniettare contenuto formattato (paragrafi, bold/italic, liste, tabelle) in un documento ODT template, gestendo:

1. **Placeholder semplici** (`$$nome$$`): sostituzione scalare di testo
2. **Placeholder ricchi** (`I_testo_obj`, `P_testo_obj`): iniezione di frammenti XML ODT nativi

## Tre Modalità di Utilizzo

### Modalità 1: Upload DOCX Diretto
- Carica un DOCX già formattato
- Il server estrae il contenuto e lo converte in XML ODT nativo
- Inietta direttamente nel template
- **Output**: ODT + PDF (opzionale)

### Modalità 2: DOCX + Editor Web
- Carica un DOCX come punto di partenza
- Visualizza e modifica il contenuto nell'editor web
- Converte le modifiche in frammenti ODT
- **Output**: ODT + PDF

### Modalità 3: Editor Web Puro
- Scrivi/formatta direttamente nell'editor web
- Toolbar per bold, italic, underline, liste
- Converte il contenuto in ODT
- **Output**: ODT + PDF

## Quick Start

### 1. Installazione
```bash
# Clone o naviga nella directory
cd /c/ClaudePrj/POCGestDoc

# Installa dipendenze
pip install -r requirements.txt

# Verifica il template
ls -la template/ASL_Template_Delibera.odt
```

### 2. Esecuzione
```bash
# Avvia il server Flask
python app.py

# Apri il browser
http://localhost:5000
```

### 3. Test
```bash
# Test del singolo modulo (odt_injector)
python test_injector.py

# Test TipTap converter
python test_tiptap.py

# Test end-to-end di tutte e tre le modalità
python test_e2e.py
```

## Architettura

```
flask app (app.py)
    ↓
Tre endpoint API:
├── /api/genera → genera ODT
├── /api/genera-pdf → genera PDF
└── /api/docx-to-tiptap → converte DOCX per editor
    ↓
Moduli di conversione:
├── odt_injector.py      (core: ZIP + placeholder substitution)
├── docx_to_odt.py       (DOCX paragraphs → ODT XML)
├── tiptap_to_odt.py     (TipTap JSON → ODT XML)
└── html_to_tiptap.py    (HTML → TipTap JSON)
    ↓
Template ODT (ZIP format):
└── content.xml (con placeholder: $$nome$$, I_testo_obj, P_testo_obj)
```

## Struttura Dati - Placeholder

### Placeholder Semplici (Scalar)
```
$$numeroproposta$$     → 2024/001
$$dataproposta$$       → 15/01/2024
$$oggetto$$            → Approvazione Piano
$$EstensoreNome$$      → Dott. Mario Rossi
... (altri)
```

### Placeholder Ricchi (XML Fragments)
```
I_testo_obj     → RELAZIONE ISTRUTTORIA (tabella)
P_testo_obj     → PROPOSTA (tabella, 1° occorrenza)
P_testo_obj_2   → DELIBERA (tabella, 2° occorrenza)
```

Ogni placeholder ricco riceve una **lista di frammenti XML ODT** che contengono paragrafi, liste, tabelle formattate.

## Formati Supportati

### DOCX → ODT
- Paragrafi semplici
- Testo con mark (bold, italic, underline, combinazioni)
- Liste puntate e numerate
- Tabelle semplici

### TipTap JSON → ODT
- Paragrafi
- Testo con marks (bold, italic, underline)
- Bullet list e ordered list
- Tabelle

### HTML → TipTap (via mammoth)
- Converte HTML (da DOCX via mammoth) in TipTap JSON
- Preserva formattazione base

## Output

### ODT
- File valido apribile con LibreOffice, Word, ecc.
- Mantiene la struttura del template
- Formattazione corretta in tutte le sezioni

### PDF
- Opzionale (richiede `soffice` installato)
- Generato da ODT via LibreOffice headless
- Fallback (messaggio di errore) se LibreOffice non disponibile

## Flussi di Test Completati

✓ **Test 1 (DOCX upload)**: DOCX → ODT fragments → Template injection
✓ **Test 2 (TipTap editor)**: TipTap JSON → ODT fragments → Template injection
✓ **Test 3 (Simple text)**: Plain text → TipTap JSON → ODT fragments → Template injection

## File Generati nel Test E2E

```
output/test_mode1.odt    # DOCX direct upload result
output/test_mode2.odt    # TipTap editor result
output/test_mode3.odt    # Simple editor result
```

## Dettagli Tecnici Importanti

### ODT = ZIP Archive
```
ASL_Template_Delibera.odt
├── mimetype (text/plain, uncompressed)
├── content.xml (main XML with placeholders)
├── styles.xml (style definitions)
├── meta.xml (metadata)
└── META-INF/manifest.xml (archive manifest)
```

### Iniezione dei Placeholder Ricchi
1. Leggi content.xml dall'ODT ZIP
2. Trova il `<text:p>` che contiene il placeholder (es. "I_testo_obj")
3. Sostituisci l'intero paragrafo con i frammenti XML generati dal converter
4. Mantieni la cella della tabella parent (`<table:table-cell>`)
5. Richiudi il ZIP

### Namespace ODT
```python
NS = {
    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
    'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
    ...
}
# Registrati con lxml.etree.register_namespace() PRIMA di qualsiasi operazione
```

## Limitazioni Attuali & Future Improvements

### Limitazioni
- Editor web usa `contenteditable` DIV (non TipTap completo via JS)
- PDF generation richiede LibreOffice (no Python fallback)
- Conversione HTML→TipTap è basica
- No support per tracked changes / comments

### Possibili Miglioramenti
1. Implementare TipTap editor JavaScript completo (non contenteditable)
2. Aggiungere mammoth per conversione DOCX→HTML migliore
3. Auto-generate styles.xml per formattazione avanzata
4. Implementare fallback PDF con reportlab
5. Database per versionamento documenti
6. Auditlog per tracking delle modifiche

## Riferimenti & Skill Utilizzate

- **docx skill**: `/skills-reference/document-skills/docx/`
  - `ooxml.md` - Document Library API e pattern XML
  - `SKILL.md` - Workflows per manipolazione DOCX

- **pdf skill**: `/skills-reference/document-skills/pdf/`
  - Pattern per generazione PDF via LibreOffice

## Troubleshooting

### "No module named 'lxml'"
```bash
pip install lxml
```

### "ModuleNotFoundError: No module named 'flask'"
```bash
pip install flask
```

### ODT file is not valid ZIP
- Verificare che `zipfile.ZipFile()` stia usando `ZIP_DEFLATED`
- Verificare che mimetype sia prima, uncompressed (`ZIP_STORED`)

### PDF generation fails
- Verificare che LibreOffice sia installato: `soffice --version`
- Su Linux: `sudo apt-get install libreoffice`
- Se non disponibile, il POC funziona comunque (solo ODT)

### Placeholder not replaced
- Verificare che il placeholder sia nel `content.xml` (unzip e ispezionare)
- Placeholder frammentati? Normalizzare il testo prima di cercare
- Verificare che la sostituzione avvenga nel nodo corretto (`<text:p>`)

## Contatto & Supporto

Questo è un Proof of Concept educativo. Per uso in produzione:
- Aggiungi error handling robusto
- Implementa authentication/authorization
- Aggiungi logging e monitoring
- Test con documenti reali complessi
- Valida tutti gli input utente

---

**Status**: POC Completo ✓
**Ultimo aggiornamento**: Marzo 2026
**Tecnologie**: Python 3.10, Flask, lxml, python-docx, LibreOffice
