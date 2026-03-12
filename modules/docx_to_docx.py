#!/usr/bin/env python3
"""Legge DOCX e converte paragrafi in formato per injector."""
from typing import Tuple, List
from docx import Document


class DocxToDocx:
    """Estrae paragrafi formattati da DOCX."""

    def convert_file(self, docx_path: str) -> Tuple[List[dict], str]:
        """
        Estrae paragrafi da DOCX preservando formattazione.

        Args:
            docx_path: Percorso al file DOCX

        Returns:
            (list of paragraph dicts, empty string for styles)
        """
        doc = Document(docx_path)
        paragraphs = []

        for para in doc.paragraphs:
            if not para.text.strip():
                continue

            # Estrai runs con formattazione
            runs = []
            for run in para.runs:
                if run.text.strip():
                    runs.append({
                        'text': run.text,
                        'bold': run.bold if run.bold is not None else False,
                        'italic': run.italic if run.italic is not None else False,
                        'underline': run.underline if run.underline is not None else False
                    })

            if runs:
                # Determina lo stile dalla lista
                style_name = para.style.name if para.style else 'Normal'

                paragraphs.append({
                    'type': 'paragraph',
                    'runs': runs,
                    'style': style_name,
                    'list_level': None
                })

        return paragraphs, ""
