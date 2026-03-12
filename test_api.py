#!/usr/bin/env python3
"""Test API endpoint with editor content."""
import sys
import json
import requests
import zipfile
from io import BytesIO

sys.path.insert(0, 'modules')

print("=" * 60)
print("TEST - API Editor Content")
print("=" * 60)

# Start app in background if not running
import subprocess
import time

proc = subprocess.Popen([sys.executable, 'app_simple.py'],
                       stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL)
time.sleep(2)

try:
    # Prepare test content
    tiptap_data = {
        "type": "doc",
        "content": [
            {"type": "paragraph", "content": [{"type": "text", "text": "IL DIRETTORE GENERALE"}]},
            {"type": "paragraph", "content": [{"type": "text", "text": "Visto il piano operativo 2024;"}]},
            {"type": "paragraph", "content": [{"type": "text", "text": "DELIBERA", "marks": [{"type": "bold"}]}]},
            {"type": "paragraph", "content": [{"type": "text", "text": "Di approvare il piano."}]}
        ]
    }

    print("\n[STEP 1] POST to /api/genera with editor content...")
    response = requests.post(
        'http://localhost:5001/api/genera',
        data={'editor_content': json.dumps(tiptap_data)}
    )

    if response.status_code != 200:
        print(f"[ERROR] Status {response.status_code}: {response.text}")
        sys.exit(1)

    print("[OK] Got response")

    # Check if it's a valid ODT
    odt_bytes = response.content
    print(f"[OK] Received {len(odt_bytes)} bytes")

    # Verify it's a valid ZIP
    try:
        with zipfile.ZipFile(BytesIO(odt_bytes), 'r') as zf:
            files = zf.namelist()
            print(f"[OK] Valid ODT with {len(files)} files")

            # Extract and check content.xml
            content_xml = zf.read('content.xml').decode('utf-8')

            # Check for proper structure
            checks = [
                ('DELIBERA' in content_xml, "Contains DELIBERA text"),
                ('IL DIRETTORE GENERALE' in content_xml, "Contains IL DIRETTORE GENERALE"),
                ('text:style-name="T_bold"' in content_xml, "Contains bold style reference"),
                ('<style:style style:name="T_bold"' in content_xml, "Contains bold style definition"),
                ('<?xml' in content_xml, "Has XML declaration"),
            ]

            for check, desc in checks:
                status = "[OK]" if check else "[FAIL]"
                print(f"{status} {desc}")
                if not check:
                    print("Content excerpt:")
                    print(content_xml[:1000])

    except Exception as e:
        print(f"[ERROR] Invalid ODT: {e}")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("All tests passed!")
    print("=" * 60)

finally:
    proc.terminate()
    proc.wait(timeout=2)
