#!/usr/bin/env python3
"""
Server WebDAV per permettere la modifica diretta dei documenti DOCX in Microsoft Word.
Espone la cartella 'output' sulla porta 8080.
"""
from wsgidav.wsgidav_app import WsgiDAVApp
from cheroot import wsgi
import os

def run_webdav():
    # Assicurati che la cartella esista
    root_path = os.path.abspath("output")
    os.makedirs(root_path, exist_ok=True)
    
    print(f"Avvio Server WebDAV su porto 8080...")
    print(f"Radice cartella: {root_path}")
    print(f"URL WebDAV: http://localhost:8080/")

    config = {
        "host": "0.0.0.0",
        "port": 8080,
        "provider_mapping": {
            "/": root_path,
        },
        "simple_dc": {
            "user_mapping": {
                "*": True, # Accesso anonimo per il POC
            },
        },
        "verbose": 1,
        "logging": {
            "enable_loggers": [],
        },
        "property_manager": True, # Necessario per Word (LOCK)
        "lock_storage": True,     # Necessario per Word (precentemente lock_manager)
    }

    app = WsgiDAVApp(config)

    server = wsgi.Server(("0.0.0.0", 8080), app)
    
    try:
        server.start()
    except KeyboardInterrupt:
        server.stop()

if __name__ == "__main__":
    run_webdav()
