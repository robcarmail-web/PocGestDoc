"""
ODT Injector v2 - Versione corretta.
Sostituisce correttamente il placeholder P_testo_obj.
"""
import zipfile
import io
from pathlib import Path
from typing import List
from lxml import etree

NS = {
    'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
}

for prefix, uri in NS.items():
    etree.register_namespace(prefix, uri)


class ODTInjectorV2:
    """Sostituisce P_testo_obj con contenuto formattato come elementi XML propri."""

    def __init__(self, template_path: str):
        self.template_path = Path(template_path)

    def inject_content(self, content_list: List[str], placeholder: str = 'P_testo_obj', styles_str: str = '') -> bytes:
        """Inietta contenuto nel placeholder sostituendo con elementi XML strutturati."""
        # Estrai ODT
        files = {}
        with zipfile.ZipFile(self.template_path, 'r') as zf:
            for name in zf.namelist():
                files[name] = zf.read(name)

        # Parsa content.xml
        root = etree.fromstring(files['content.xml'])

        # Inietta gli stili se forniti
        if styles_str:
            self._inject_styles(root, styles_str)

        # Trova e sostituisci il placeholder
        replaced = False
        for para in root.findall('.//text:p', NS):
            # Estrai tutto il testo del paragrafo
            full_text = ''.join([node.text or '' for node in para.iter() if node.text])

            if placeholder in full_text:
                print(f"[DEBUG] Found placeholder in para")

                # Rimuovi tutti gli elementi figli
                for child in list(para):
                    para.remove(child)
                para.text = None
                para.tail = None

                # Processa ogni frammento come elemento XML
                for fragment_xml_str in content_list:
                    try:
                        # Parsa il frammento come elemento XML
                        fragment_elem = etree.fromstring(fragment_xml_str)

                        # Se e' un paragrafo, copia il suo contenuto al placeholder
                        if fragment_elem.tag == '{' + NS['text'] + '}p':
                            # Copia tutti i children del paragrafo
                            for child in fragment_elem:
                                # Crea una copia profonda dell'elemento
                                child_copy = etree.Element(child.tag, attrib=dict(child.attrib))
                                child_copy.text = child.text
                                child_copy.tail = child.tail

                                # Copia ricorsivamente i sotto-elementi
                                self._copy_children(child, child_copy)

                                para.append(child_copy)

                            # Copia il testo diretto del paragrafo se presente
                            if fragment_elem.text and fragment_elem.text.strip():
                                if len(para) > 0:
                                    # Aggiungi come tail del primo elemento
                                    if para[0].tail:
                                        para[0].tail += fragment_elem.text
                                    else:
                                        para[0].tail = fragment_elem.text
                                else:
                                    para.text = fragment_elem.text
                        else:
                            # Per elementi non paragrafo (liste, tabelle), appendi direttamente
                            para.addnext(fragment_elem)

                    except etree.XMLSyntaxError as e:
                        print(f"[WARNING] Failed to parse fragment: {e}")
                        continue

                replaced = True
                print(f"[DEBUG] Placeholder replaced successfully")
                break

        if not replaced:
            print(f"[WARNING] Placeholder '{placeholder}' not found in document")

        # Serializza
        files['content.xml'] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', pretty_print=False)

        # Ricrea ZIP
        output = io.BytesIO()
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zf:
            if 'mimetype' in files:
                zf.writestr('mimetype', files['mimetype'], compress_type=zipfile.ZIP_STORED)
            for name, data in files.items():
                if name != 'mimetype':
                    zf.writestr(name, data)

        return output.getvalue()

    def _copy_children(self, source: etree._Element, target: etree._Element) -> None:
        """Copia ricorsivamente tutti i figli di source in target."""
        for child in source:
            child_copy = etree.Element(child.tag, attrib=dict(child.attrib))
            child_copy.text = child.text
            child_copy.tail = child.tail
            self._copy_children(child, child_copy)
            target.append(child_copy)

    def save(self, output_path: str, content_list: List[str], placeholder: str = 'P_testo_obj', styles_str: str = '') -> None:
        """Salva ODT su disco."""
        odt_bytes = self.inject_content(content_list, placeholder, styles_str)
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, 'wb') as f:
            f.write(odt_bytes)

    def _inject_styles(self, root: etree._Element, styles_str: str) -> None:
        """Inietta le definizioni di stile nel documento."""
        # Trova l'elemento office:automatic-styles
        auto_styles = root.find('.//{' + NS['office'] + '}automatic-styles')
        if auto_styles is None:
            return

        # Parsa e aggiungi ogni stile
        try:
            # Avvolgi gli stili in un elemento temporaneo per parserl
            wrapped_styles = f'<root xmlns:style="{NS["style"]}" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">{styles_str}</root>'
            styles_root = etree.fromstring(wrapped_styles.encode('utf-8'))

            for style_elem in styles_root.findall('.//'):
                if style_elem.tag != '{' + NS['style'] + '}style':
                    continue
                # Crea una copia dello stile
                style_copy = etree.Element('{' + NS['style'] + '}style', attrib=dict(style_elem.attrib))
                style_copy.text = style_elem.text
                style_copy.tail = style_elem.tail
                self._copy_children(style_elem, style_copy)
                auto_styles.append(style_copy)
        except Exception as e:
            print(f"[WARNING] Failed to inject styles: {e}")
