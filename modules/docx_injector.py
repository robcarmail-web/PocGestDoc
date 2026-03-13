#!/usr/bin/env python3
"""
Inietta contenuto in template DOCX.
Approccio: copia XML direttamente dal sorgente al template usando lxml,
preservando liste, tabelle, formattazione completa.
"""
import zipfile
from io import BytesIO
from pathlib import Path
from copy import deepcopy
from lxml import etree
from typing import List, Optional, Dict, Tuple

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML_SPACE = 'http://www.w3.org/XML/1998/namespace'


def _w(tag):
    return '{%s}%s' % (W, tag)


class DocxInjector:
    """Sostituisce placeholder semplici e ricchi in un template DOCX."""

    def __init__(self, template_path: str):
        self.template_path = Path(template_path)

    # ------------------------------------------------------------------ #
    # Metodi di iniezione multipla
    # ------------------------------------------------------------------ #

    def inject_placeholders(self,
                            simple_data: Dict[str, str],
                            rich_content: Optional[Dict[str, str]] = None) -> bytes:
        """
        Sostituisce nel template:
        1. simple_data: dizionario { 'oggetto': 'Valore Oggetto', ... } -> sostituisce il testo inline
        2. rich_content: dizionario { 'I_testo_obj': '/path/to/fragment.docx', ... } -> inietta intero body
        """
        rich_content = rich_content or {}

        # Leggi il template ZIP
        template_files = self._read_zip(self.template_path)
        tpl_doc = etree.fromstring(template_files['word/document.xml'])
        tpl_body = tpl_doc.find('.//{%s}body' % W)

        # 1. Simple Replacements
        self._replace_simple_placeholders(tpl_body, simple_data)

        # 2. Rich Replacements (DOCX upload)
        for placeholder_key, source_docx_path in rich_content.items():
            try:
                source_files = self._read_zip(source_docx_path)
                src_doc = etree.fromstring(source_files['word/document.xml'])

                placeholder_elem = self._find_placeholder(tpl_body, placeholder_key)
                if placeholder_elem is None:
                    print(f'[WARNING] Rich placeholder "{placeholder_key}" non trovato nel template.')
                    continue

                # Estrai dal fragment
                src_body = src_doc.find('.//{%s}body' % W)
                src_elements = [deepcopy(child) for child in src_body
                                if child.tag != _w('sectPr')]

                # Inserisci nel posizionamento salvato
                for elem in reversed(src_elements):
                    placeholder_elem.addnext(elem)

                # Rimuovi l'elemento placeholder dal DOM (prendendo il parent esatto)
                placeholder_elem.getparent().remove(placeholder_elem)

                # Unisci gli asset collegati dal sorgente nel template (Numerazioni e Stili)
                if 'word/numbering.xml' in source_files:
                    template_files['word/numbering.xml'] = source_files['word/numbering.xml']
                    template_files = self._update_manifest(template_files, 'word/numbering.xml',
                                                           'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml')
                    template_files = self._update_rels(template_files, 'word/numbering.xml')

                if 'word/styles.xml' in source_files:
                    template_files['word/styles.xml'] = self._merge_styles(
                        template_files.get('word/styles.xml'),
                        source_files['word/styles.xml']
                    )

            except Exception as e:
                print(f"[ERROR] Iniezione fallita per {placeholder_key}: {str(e)}")

        # Salva back nel template virtuale
        template_files['word/document.xml'] = etree.tostring(
            tpl_doc, xml_declaration=True, encoding='UTF-8', standalone=True
        )

        return self._rebuild_zip(template_files)

    # ------------------------------------------------------------------ #
    # Gestione nodi
    # ------------------------------------------------------------------ #

    def _replace_simple_placeholders(self, body_element, simple_data: dict):
        """Sostituisce i placeholder nel testo scalar mantenendo la formattazione esistente."""
        if not simple_data:
            return

        for paragraph in body_element.iter(_w('p')):
            for key, val in simple_data.items():
                target_str = f'$${key}$$'
                # Metodo semplice: concatena il testo logico
                p_text = ''.join(t.text or '' for t in paragraph.iter(_w('t')))
                if target_str in p_text:
                    # Abbiamo trovato il target nel paragrafo.
                    # Per evitare split complessi di XML runs, sostituiamo brutale
                    # raccogliendo il testo. Un approccio più robusto farebbe chunking nel run.
                    self._replace_text_in_paragraph(paragraph, target_str, str(val))

    def _replace_text_in_paragraph(self, paragraph, search_str: str, replace_str: str):
        """Versione basica e sicura per ricerca/sostituzione nei text run Ooxml."""
        # Se un paragrafo contiene una run con frammentazione, lo puliamo e uniamo tutto in un solo run,
        # per semplicità di sostituzione in questo POC.
        all_text = ""
        runs = list(paragraph.findall('.//{%s}r' % W))
        if not runs:
            return

        first_run_props = None
        for r in runs:
            r_pr = r.find('{%s}rPr' % W)
            if r_pr is not None and first_run_props is None:
                first_run_props = deepcopy(r_pr)

            for t in r.findall('{%s}t' % W):
                all_text += t.text or ""
            # Pulisce i vecchi attributi
            paragraph.remove(r)

        new_text = all_text.replace(search_str, replace_str)

        new_run = etree.SubElement(paragraph, _w('r'))
        if first_run_props is not None:
            new_run.append(first_run_props)
        new_t = etree.SubElement(new_run, _w('t'))
        new_t.set('{%s}space' % XML_SPACE, 'preserve')
        new_t.text = new_text

    def _find_placeholder(self, body, placeholder: str):
        """Cerca il paragrafo che contiene esattamente il placeholder desiderato."""
        for p in body.iter(_w('p')):
            text = ''.join(t.text or '' for t in p.iter(_w('t')))
            if placeholder in text:
                return p
        return None

    # ------------------------------------------------------------------ #
    # Utility Zip e XML
    # ------------------------------------------------------------------ #

    def _read_zip(self, path) -> dict:
        files = {}
        with zipfile.ZipFile(path, 'r') as zf:
            for name in zf.namelist():
                files[name] = zf.read(name)
        return files

    def _rebuild_zip(self, files: dict) -> bytes:
        output = BytesIO()
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zf:
            if '[Content_Types].xml' in files:
                zf.writestr('[Content_Types].xml', files['[Content_Types].xml'])
            for name, data in files.items():
                if name != '[Content_Types].xml':
                    zf.writestr(name, data)
        return output.getvalue()

    def _merge_styles(self, tpl_styles_bytes: Optional[bytes], src_styles_bytes: bytes) -> bytes:
        """Aggiunge stili del sorgente che non esistono nel template."""
        if not tpl_styles_bytes:
            return src_styles_bytes

        tpl_root = etree.fromstring(tpl_styles_bytes)
        src_root = etree.fromstring(src_styles_bytes)

        S = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

        existing_ids = set()
        for style in tpl_root.findall('.//{%s}style' % S):
            sid = style.get('{%s}styleId' % S)
            if sid:
                existing_ids.add(sid)

        for style in src_root.findall('.//{%s}style' % S):
            sid = style.get('{%s}styleId' % S)
            if sid and sid not in existing_ids:
                tpl_root.append(deepcopy(style))

        return etree.tostring(tpl_root, xml_declaration=True, encoding='UTF-8', standalone=True)

    def _update_manifest(self, files: dict, part_name: str, content_type: str) -> dict:
        """Aggiunge una entry al [Content_Types].xml se non esiste."""
        if '[Content_Types].xml' not in files:
            return files

        root = etree.fromstring(files['[Content_Types].xml'])
        CT = 'http://schemas.openxmlformats.org/package/2006/content-types'

        part_uri = '/' + part_name
        for override in root.findall('{%s}Override' % CT):
            if override.get('PartName') == part_uri:
                return files

        override = etree.SubElement(root, '{%s}Override' % CT)
        override.set('PartName', part_uri)
        override.set('ContentType', content_type)

        files['[Content_Types].xml'] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        return files

    def _update_rels(self, files: dict, part_name: str) -> dict:
        """Aggiunge la relazione al word/_rels/document.xml.rels."""
        rels_path = 'word/_rels/document.xml.rels'
        if rels_path not in files:
            return files

        root = etree.fromstring(files[rels_path])
        R = 'http://schemas.openxmlformats.org/package/2006/relationships'

        target = part_name.replace('word/', '')
        for rel in root.findall('{%s}Relationship' % R):
            if rel.get('Target') == target:
                return files

        existing_ids = {rel.get('Id') for rel in root.findall('{%s}Relationship' % R)}
        rid = 'rId' + str(len(existing_ids) + 20)

        rel = etree.SubElement(root, '{%s}Relationship' % R)
        rel.set('Id', rid)
        rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering')
        rel.set('Target', target)

        files[rels_path] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        return files

