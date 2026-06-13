#!/usr/bin/env python3
"""
Server WebDAV per permettere la modifica diretta dei documenti DOCX in Microsoft Word.
Espone la cartella output sulla porta configurabile (default 8080).
"""
from wsgidav.wsgidav_app import WsgiDAVApp
from cheroot import wsgi
import os


def _env(key, default):
    return os.environ.get(key, default)


def run_webdav():
    output_dir = _env('OUTPUT_DIR', 'output')
    host = _env('WEBDAV_HOST', '0.0.0.0')
    port = int(_env('WEBDAV_PORT', '8080'))
    root_path = os.path.abspath(output_dir)
    os.makedirs(root_path, exist_ok=True)

    print(f'Avvio Server WebDAV su {host}:{port}...')
    print(f'Radice cartella: {root_path}')
    print(f'URL WebDAV: http://{host if host != "0.0.0.0" else "localhost"}:{port}/')

    config = {
        'host': host,
        'port': port,
        'provider_mapping': {
            '/': root_path,
        },
        'simple_dc': {
            'user_mapping': {
                '*': True,
            },
        },
        'verbose': 1,
        'logging': {
            'enable_loggers': [],
        },
        'property_manager': True,
        'lock_storage': True,
    }

    app = WsgiDAVApp(config)
    server = wsgi.Server((host, port), app)

    try:
        server.start()
    except KeyboardInterrupt:
        server.stop()


if __name__ == '__main__':
    run_webdav()
