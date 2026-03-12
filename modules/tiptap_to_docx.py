#!/usr/bin/env python3
"""Converte TipTap JSON in paragrafi DOCX con formattazione preservata."""
import json
from typing import List, Tuple, Optional
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt, RGBColor


class TiptapToDocx:
    """Converte TipTap JSON a paragrafi DOCX."""

    def convert(self, tiptap_json: str) -> Tuple[List[dict], str]:
        """
        Converte TipTap JSON a lista di paragrafi DOCX.

        Args:
            tiptap_json: JSON string dal TipTap editor

        Returns:
            (list of paragraph dicts, styles string - vuoto per DOCX)
        """
        data = json.loads(tiptap_json)

        if not isinstance(data, dict) or 'content' not in data:
            raise ValueError("Invalid TipTap format: missing 'content' field")

        paragraphs = []
        for node in data.get('content', []):
            para_info = self._convert_node(node)
            if para_info:
                paragraphs.append(para_info)

        return paragraphs, ""  # DOCX non richiede stili separati

    def _convert_node(self, node: dict) -> Optional[dict]:
        """Converte un nodo TipTap a paragrafo DOCX."""
        node_type = node.get('type')

        if node_type == 'paragraph':
            return self._convert_paragraph(node)
        elif node_type == 'bulletList':
            return self._convert_bullet_list(node)
        elif node_type == 'orderedList':
            return self._convert_ordered_list(node)
        else:
            return None

    def _convert_paragraph(self, node: dict) -> dict:
        """Converte paragrafo con formattazione inline."""
        runs = []

        for child in node.get('content', []):
            if child.get('type') == 'text':
                run_info = self._create_run_info(child)
                runs.append(run_info)

        return {
            'type': 'paragraph',
            'runs': runs,
            'style': 'Normal',
            'list_level': None
        }

    def _create_run_info(self, text_node: dict) -> dict:
        """Crea info su run (testo con formattazione)."""
        text = text_node.get('text', '')
        marks = text_node.get('marks', [])

        # Estrai proprietà da marks
        bold = any(m.get('type') == 'bold' for m in marks)
        italic = any(m.get('type') == 'italic' for m in marks)
        underline = any(m.get('type') == 'underline' for m in marks)

        return {
            'text': text,
            'bold': bold,
            'italic': italic,
            'underline': underline
        }

    def _convert_bullet_list(self, node: dict) -> dict:
        """Converte lista puntata."""
        return self._convert_list(node, list_type='bullet')

    def _convert_ordered_list(self, node: dict) -> dict:
        """Converte lista numerata."""
        return self._convert_list(node, list_type='ordered')

    def _convert_list(self, node: dict, list_type: str) -> dict:
        """Converte lista generica."""
        # Restituisce una lista di paragrafi con list_level
        result = {
            'type': 'list',
            'list_type': list_type,
            'items': []
        }

        for child in node.get('content', []):
            if child.get('type') == 'listItem':
                item = self._convert_list_item(child)
                result['items'].append(item)

        return result

    def _convert_list_item(self, node: dict) -> dict:
        """Converte item di lista."""
        runs = []

        for child in node.get('content', []):
            if child.get('type') == 'paragraph':
                for text_child in child.get('content', []):
                    if text_child.get('type') == 'text':
                        run_info = self._create_run_info(text_child)
                        runs.append(run_info)

        return {
            'type': 'paragraph',
            'runs': runs,
            'style': 'List Bullet' if 'bullet' in str(node) else 'List Number'
        }
