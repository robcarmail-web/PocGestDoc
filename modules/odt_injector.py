"""
Core module for ODT template injection.
Handles placeholder replacement in ODT files.
"""
import zipfile
import io
from lxml import etree
from pathlib import Path
from typing import Dict, Optional, List

# ODT XML namespaces
NS = {
    'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
    'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
    'fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
    'svg': 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
    'draw': 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
}

# Register namespaces globally for lxml
for prefix, uri in NS.items():
    etree.register_namespace(prefix, uri)


class ODTInjector:
    """Manages ODT template manipulation and placeholder injection."""

    def __init__(self, template_path: str):
        """Initialize with ODT template path."""
        self.template_path = Path(template_path)
        self.root = None
        self.content_tree = None

    def _extract_odt(self) -> Dict[str, bytes]:
        """Extract all files from ODT ZIP archive."""
        files = {}
        with zipfile.ZipFile(self.template_path, 'r') as zf:
            for name in zf.namelist():
                files[name] = zf.read(name)
        return files

    def _pack_odt(self, files: Dict[str, bytes]) -> bytes:
        """Pack files back into ODT ZIP archive."""
        output = io.BytesIO()
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zf:
            # Add mimetype first, uncompressed
            if 'mimetype' in files:
                zf.writestr('mimetype', files['mimetype'], compress_type=zipfile.ZIP_STORED)

            # Add other files
            for name, data in files.items():
                if name != 'mimetype':
                    zf.writestr(name, data)

        return output.getvalue()

    def _parse_content_xml(self, content_xml: bytes) -> etree._Element:
        """Parse content.xml as XML tree."""
        return etree.fromstring(content_xml)

    def _serialize_content_xml(self, root: etree._Element) -> bytes:
        """Serialize XML tree back to bytes."""
        return etree.tostring(root, xml_declaration=True, encoding='UTF-8', pretty_print=False)

    def _normalize_text(self, elem: etree._Element) -> str:
        """Extract text from element accounting for fragmenting across spans."""
        text_parts = []
        for span in elem.findall('.//text:span', NS):
            if span.text:
                text_parts.append(span.text)
        # Also check direct text in paragraphs
        if elem.text:
            text_parts.insert(0, elem.text)
        return ''.join(text_parts)

    def _replace_simple_placeholder(self, root: etree._Element, placeholder: str, value: str) -> bool:
        """Replace simple placeholder ($$name$$) with scalar value."""
        replaced = False
        pattern = f'$${placeholder}$$'

        # Find all paragraphs and text spans
        for para in root.findall('.//text:p', NS):
            para_text = self._normalize_text(para)

            if pattern in para_text:
                # Strategy: rebuild paragraph preserving styles
                # Remove all existing spans and runs
                for span in list(para.findall('text:span', NS)):
                    para.remove(span)

                # Add new span with the value
                new_span = etree.Element('{' + NS['text'] + '}span')
                new_span.text = value
                para.append(new_span)
                replaced = True

        return replaced

    def _inject_rich_content(self, root: etree._Element, placeholder: str, content_nodes: List[str]) -> bool:
        """
        Replace placeholder_obj with formatted content (XML fragments).
        Replaces the <text:p> node containing placeholder with content nodes.
        """
        replaced = False

        for para in root.findall('.//text:p', NS):
            para_text = self._normalize_text(para)

            if placeholder in para_text:
                # Find the parent table cell
                parent_cell = para.getparent()
                if parent_cell is not None and parent_cell.tag == '{' + NS['table'] + '}table-cell':
                    # Replace the paragraph with new content
                    parent_idx = list(parent_cell).index(para)
                    parent_cell.remove(para)

                    # Insert content nodes
                    for i, content_xml_str in enumerate(content_nodes):
                        try:
                            content_elem = etree.fromstring(content_xml_str)
                            parent_cell.insert(parent_idx + i, content_elem)
                        except etree.XMLSyntaxError as e:
                            raise ValueError(f"Invalid XML fragment: {e}")

                    replaced = True
                    break

        return replaced

    def _add_styles(self, root: etree._Element, styles_xml: str) -> None:
        """Add style definitions to automatic-styles section."""
        if not styles_xml or not styles_xml.strip():
            return

        # Find or create office:automatic-styles
        auto_styles = root.find('.//office:automatic-styles', NS)
        if auto_styles is None:
            # Create if doesn't exist
            office_body = root.find('.//office:body', NS)
            auto_styles = etree.Element('{' + NS['office'] + '}automatic-styles')
            if office_body is not None:
                office_body.addprevious(auto_styles)
            else:
                root.insert(0, auto_styles)

        # Parse and add style elements
        for style_line in styles_xml.split('\n'):
            style_line = style_line.strip()
            if not style_line:
                continue
            try:
                style_elem = etree.fromstring(style_line)
                # Check if style already exists
                style_name = style_elem.get('{' + NS['style'] + '}name')
                existing = auto_styles.find(f".//style:style[@style:name='{style_name}']", NS)
                if existing is None:
                    auto_styles.append(style_elem)
            except etree.XMLSyntaxError:
                # Skip invalid style lines
                pass

    def inject_placeholders(self, data: Dict[str, str], rich_content: Optional[Dict[str, tuple]] = None) -> bytes:
        """
        Inject placeholders into ODT template.

        Args:
            data: Dict of simple placeholders (name -> value)
            rich_content: Dict of rich placeholders (name -> (fragments_list, styles_string))

        Returns:
            Modified ODT file as bytes
        """
        # Extract and parse
        files = self._extract_odt()
        content_xml = files['content.xml']
        root = self._parse_content_xml(content_xml)

        # Replace simple placeholders
        for placeholder, value in data.items():
            self._replace_simple_placeholder(root, placeholder, str(value))

        # Replace rich content placeholders
        if rich_content:
            for placeholder, content_data in rich_content.items():
                if isinstance(content_data, tuple):
                    xml_fragments, styles_str = content_data
                else:
                    # Backward compatibility: just fragments
                    xml_fragments = content_data
                    styles_str = ""

                self._inject_rich_content(root, placeholder, xml_fragments)
                self._add_styles(root, styles_str)

        # Serialize back
        files['content.xml'] = self._serialize_content_xml(root)

        # Pack and return
        return self._pack_odt(files)

    def save(self, output_path: str, data: Dict[str, str], rich_content: Optional[Dict[str, List[str]]] = None) -> None:
        """Generate modified ODT and save to file."""
        odt_bytes = self.inject_placeholders(data, rich_content)
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, 'wb') as f:
            f.write(odt_bytes)
