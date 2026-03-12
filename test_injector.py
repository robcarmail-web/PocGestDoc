#!/usr/bin/env python3
"""Test odt_injector module."""
import sys
sys.path.insert(0, 'modules')

from odt_injector import ODTInjector

# Test data
test_data = {
    'numeroproposta': '2024/001',
    'dataproposta': '15/01/2024',
    'oggetto': 'Test Proposal',
    'EstensoreNome': 'Dott. Mario Rossi',
    'DirGenNome': 'Prof. Roberto Blu',
}

# Create injector
injector = ODTInjector('template/ASL_Template_Delibera.odt')

# Generate output
output_path = 'output/test_output.odt'
injector.save(output_path, test_data)

print(f"[OK] Generated: {output_path}")

# Verify output is a valid ZIP
import zipfile
try:
    with zipfile.ZipFile(output_path, 'r') as zf:
        print(f"[OK] Output is valid ZIP with {len(zf.namelist())} files")

        # Check content.xml
        content = zf.read('content.xml')
        print(f"[OK] content.xml size: {len(content)} bytes")

        # Verify placeholder replacement
        if b'2024/001' in content:
            print("[OK] Placeholder 'numeroproposta' correctly replaced")
        else:
            print("[FAIL] Placeholder 'numeroproposta' NOT found")

except Exception as e:
    print(f"[FAIL] {e}")
