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

                # Normalizza il contenuto iniettato per adattarlo al template
                src_default_tab = self._get_default_tab_stop(source_files)
                tpl_default_tab = self._get_default_tab_stop(template_files)
                tpl_text_width = self._get_text_area_width(template_files)
                for elem in src_elements:
                    self._normalize_content(elem, src_default_tab, tpl_default_tab, tpl_text_width)

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

    def _normalize_content(self, element, src_default_tab: int, tpl_default_tab: int, tpl_text_width: int):
        """Normalizza un elemento iniettato per adattarlo al layout del template.

        Applica in sequenza:
        1. Tab stop espliciti (se defaultTabStop diverso)
        2. Pulizia spazi usati come padding
        3. Rimozione paragrafi vuoti ridondanti in coda
        """
        if src_default_tab != tpl_default_tab:
            self._make_tabs_explicit(element, src_default_tab, tpl_text_width)
        self._clean_space_padding(element)

    def _clean_space_padding(self, element):
        """Rimuove run con solo spazi usati come padding tra contenuto reale.

        Pattern tipico: run con 20+ spazi inseriti manualmente per simulare
        allineamento. Li rimuove e lascia che il tab stop faccia il lavoro.
        """
        for p in element.iter(_w('p')):
            runs = list(p.findall(_w('r')))
            has_tab = any(r.findall(_w('tab')) for r in runs)
            if not has_tab:
                continue

            for r in runs:
                texts = r.findall(_w('t'))
                tabs = r.findall(_w('tab'))
                if not tabs and len(texts) == 1:
                    text = texts[0].text or ''
                    # Run con solo spazi (>= 5) tra run con tab e run con testo
                    if len(text) >= 5 and not text.strip():
                        p.remove(r)

    def _get_default_tab_stop(self, docx_files: dict) -> int:
        """Legge il defaultTabStop da settings.xml. Default OOXML: 720 (1.27cm)."""
        if 'word/settings.xml' not in docx_files:
            return 720
        settings = etree.fromstring(docx_files['word/settings.xml'])
        dts = settings.find('.//{%s}defaultTabStop' % W)
        if dts is not None:
            try:
                return int(dts.get(_w('val')))
            except (ValueError, TypeError):
                pass
        return 720

    def _get_text_area_width(self, docx_files: dict) -> int:
        """Calcola la larghezza dell'area testo (page width - margini L/R)."""
        if 'word/document.xml' not in docx_files:
            return 9072  # ~16cm fallback
        doc = etree.fromstring(docx_files['word/document.xml'])
        sectPr = doc.find('.//{%s}sectPr' % W)
        if sectPr is None:
            return 9072
        pgSz = sectPr.find('{%s}pgSz' % W)
        pgMar = sectPr.find('{%s}pgMar' % W)
        if pgSz is None or pgMar is None:
            return 9072
        try:
            w = int(pgSz.get(_w('w')))
            l = int(pgMar.get(_w('left')))
            r = int(pgMar.get(_w('right')))
            return w - l - r
        except (ValueError, TypeError):
            return 9072

    def _make_tabs_explicit(self, element, src_default_tab: int, tpl_text_width: int):
        """Normalizza i tab nei paragrafi iniettati per adattarli al template.

        - Pochi tab (<=3): aggiunge tab stop espliciti con l'intervallo del sorgente.
        - Molti tab (>3): pattern di 'fake right alignment', convertito in un singolo
          tab stop right-aligned alla larghezza del template.
        """
        for p in element.iter(_w('p')):
            tab_count = sum(1 for r in p.iter(_w('r'))
                           for _ in r.findall(_w('tab')))
            if tab_count == 0:
                continue

            pPr = p.find(_w('pPr'))
            if pPr is not None and pPr.find(_w('tabs')) is not None:
                continue

            if pPr is None:
                pPr = etree.SubElement(p, _w('pPr'))
                p.insert(0, pPr)

            tabs_elem = etree.SubElement(pPr, _w('tabs'))

            if tab_count > 3:
                # Molti tab = fake right alignment.
                # Sostituiamo i 18+ tab con un unico tab right-aligned,
                # e comprimiamo i run eliminando i tab ridondanti.
                tab = etree.SubElement(tabs_elem, _w('tab'))
                tab.set(_w('val'), 'right')
                tab.set(_w('pos'), str(tpl_text_width))

                # Comprimi: tieni solo il primo tab e rimuovi quelli extra
                self._collapse_tabs(p)
            else:
                # Pochi tab: mantieni l'intervallo originale del sorgente
                for i in range(1, tab_count + 2):
                    tab = etree.SubElement(tabs_elem, _w('tab'))
                    tab.set(_w('val'), 'left')
                    tab.set(_w('pos'), str(src_default_tab * i))

    def _collapse_tabs(self, paragraph):
        """Comprime run con molti tab consecutivi in un singolo tab."""
        runs = list(paragraph.findall(_w('r')))
        found_first_tab = False
        runs_to_remove = []

        for r in runs:
            tabs_in_run = r.findall(_w('tab'))
            texts_in_run = r.findall(_w('t'))
            has_real_text = any((t.text or '').strip() for t in texts_in_run)

            if tabs_in_run and not has_real_text:
                if found_first_tab:
                    # Tab ridondante - rimuovi intero run
                    runs_to_remove.append(r)
                else:
                    # Primo tab - tieni solo un <w:tab/>, rimuovi extra
                    found_first_tab = True
                    for extra_tab in tabs_in_run[1:]:
                        r.remove(extra_tab)
                    # Rimuovi anche gli spazi bianchi nei <w:t> dello stesso run
                    for t in texts_in_run:
                        if not (t.text or '').strip():
                            r.remove(t)
            elif has_real_text and found_first_tab:
                # Run con testo dopo i tab - tieni così com'è
                break

        for r in runs_to_remove:
            paragraph.remove(r)

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

