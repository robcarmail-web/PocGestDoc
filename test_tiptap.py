#!/usr/bin/env python3
"""Test tiptap_to_odt module."""
import sys
import json
sys.path.insert(0, 'modules')

from tiptap_to_odt import TiptapToODT

# Test data - simple TipTap JSON
tiptap_data = {
    "type": "doc",
    "content": [
        {
            "type": "paragraph",
            "content": [
                {
                    "type": "text",
                    "text": "This is ",
                    "marks": []
                },
                {
                    "type": "text",
                    "text": "bold",
                    "marks": [{"type": "bold"}]
                },
                {
                    "type": "text",
                    "text": " and ",
                    "marks": []
                },
                {
                    "type": "text",
                    "text": "italic",
                    "marks": [{"type": "italic"}]
                },
                {
                    "type": "text",
                    "text": " text."
                }
            ]
        },
        {
            "type": "bulletList",
            "content": [
                {
                    "type": "listItem",
                    "content": [
                        {
                            "type": "paragraph",
                            "content": [
                                {"type": "text", "text": "First item"}
                            ]
                        }
                    ]
                },
                {
                    "type": "listItem",
                    "content": [
                        {
                            "type": "paragraph",
                            "content": [
                                {"type": "text", "text": "Second item"}
                            ]
                        }
                    ]
                }
            ]
        }
    ]
}

# Convert
converter = TiptapToODT()
tiptap_json = json.dumps(tiptap_data)
fragments, styles = converter.convert(tiptap_json)

print(f"[OK] Converted {len(fragments)} fragments")
print(f"\nFragment 1 (paragraph):")
print(fragments[0][:150] + "...")

print(f"\nFragment 2 (list):")
print(fragments[1][:100] + "...")

print(f"\nStyles generated:")
print(styles if styles else "(none)")
