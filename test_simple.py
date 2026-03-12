#!/usr/bin/env python3
"""Test versione semplificata."""
import sys
import json
sys.path.insert(0, 'modules')

from odt_injector_v2 import ODTInjectorV2
from tiptap_to_odt import TiptapToODT

print("=" * 60)
print("TEST - Versione Semplificata")
print("=" * 60)

tiptap_data = {
    "type": "doc",
    "content": [
        {"type": "paragraph", "content": [{"type": "text", "text": "IL DIRETTORE GENERALE"}]},
        {"type": "paragraph", "content": [{"type": "text", "text": "Visto il piano operativo 2024;"}]},
        {"type": "paragraph", "content": [{"type": "text", "text": "DELIBERA", "marks": [{"type": "bold"}]}]},
        {"type": "paragraph", "content": [{"type": "text", "text": "Di approvare il piano."}]}
    ]
}

print("\n[STEP 1] Convert to ODT...")
converter = TiptapToODT()
fragments, styles = converter.convert(json.dumps(tiptap_data))
print(f"[OK] {len(fragments)} fragments")

print("\n[STEP 2] Inject content...")
injector = ODTInjectorV2('template/ASL_Template_Simple.odt')
injector.save('output/test_simple.odt', fragments, 'P_testo_obj', styles)
print("[OK] output/test_simple.odt")

print("\n" + "=" * 60)
print("Ready: http://localhost:5000")
print("=" * 60)
