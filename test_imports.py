#!/usr/bin/env python3
"""Test that all modules can be imported."""
import sys
sys.path.insert(0, 'modules')

try:
    from odt_injector import ODTInjector
    print("[OK] odt_injector imported")
except Exception as e:
    print(f"[FAIL] odt_injector: {e}")

try:
    from tiptap_to_odt import TiptapToODT
    print("[OK] tiptap_to_odt imported")
except Exception as e:
    print(f"[FAIL] tiptap_to_odt: {e}")

try:
    from html_to_tiptap import HTMLToTiptap
    print("[OK] html_to_tiptap imported")
except Exception as e:
    print(f"[FAIL] html_to_tiptap: {e}")

try:
    from docx_to_odt import DocxToODT
    print("[OK] docx_to_odt imported")
except Exception as e:
    print(f"[FAIL] docx_to_odt: {e}")

try:
    import flask
    print("[OK] flask imported")
except Exception as e:
    print(f"[FAIL] flask: {e}")

print("\nAll modules imported successfully!")
