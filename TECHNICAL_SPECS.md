# Requisiti Tecnici e Stack Tecnologico del POC

Questo documento elenca tutte le tecnologie, i moduli e le configurazioni necessarie per il corretto funzionamento del sistema di iniezione DOCX, integrazione WebDAV e conversione PDF.

## 1. Linguaggio e Ambiente
- **Python 3.10+**: Il linguaggio di programmazione principale utilizzato per il backend.
- **Ambiente Windows**: Necessario per l'integrazione nativa con Microsoft Word (via WebDAV e per la conversione PDF ad alta fedeltà).

## 2. Framework e Server Web
- **Flask (v3.0+)**: Framework web per gestire le API di generazione e l'interfaccia utente.
- **WsgiDAV (v4.3+)**: Implementazione del server WebDAV in Python per permettere l'apertura e il salvataggio diretto dei file in Word.
- **Cheroot**: Server WSGI ad alte prestazioni utilizzato per ospitare l'applicazione WsgiDAV.

## 3. Moduli Python (Dipendenze)
Le seguenti librerie devono essere installate tramite `pip`:

| Modulo | Scopo |
| :--- | :--- |
| `flask` | Gestione dell'applicazione web e degli endpoint API. |
| `python-docx` | Creazione e manipolazione programmatica dei file .docx (utilizzato per il template base). |
| `lxml` | Parsing e manipolazione dell'XML (OpenXML) per l'iniezione ad alta fedeltà. |
| `wsgidav` | Fornisce il protocollo WebDAV per la modifica remota. |
| `cheroot` | Server per WsgiDAV. |
| `docx2pdf` | Conversione da DOCX a PDF utilizzando l'API di Microsoft Word. |
| `pywin32` | Necessario per `docx2pdf` per comunicare con Word tramite COM. |
| `pythoncom` | Integrato in `pywin32`, gestisce l'inizializzazione dei thread per l'uso dell'API di Windows. |

## 4. Requisiti Lato Client / Sistema
- **Microsoft Word (Desktop)**: Installato sul PC locale per permettere l'apertura tramite `ms-word:` protocol e per la generazione dei PDF.
- **Porte di Comunicazione**:
  - `5000`: Utilizzata dall'applicazione Flask (Backend/UI).
  - `8080`: Utilizzata dal server WebDAV (Open/Save Word).

## 5. Installazione Rapida
Per installare tutti i moduli necessari in un colpo solo, eseguire:

```bash
pip install flask python-docx lxml wsgidav cheroot docx2pdf pywin32
```

## 6. Struttura dei File Critici nel POC
- `app.py`: Logica principale e API Flask.
- `webdav_server.py`: Server dedicato alla modifica in tempo reale.
- `modules/docx_injector.py`: Motore XML per l'unione dei file DOCX.
- `output/`: Cartella condivisa via WebDAV dove vengono memorizzati i documenti generati e modificati.
