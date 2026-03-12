"""
Converts TipTap JSON editor format to ODT XML fragments.
Handles paragraphs, text formatting, lists, tables, and hard breaks.
"""
import json
from typing import Dict, List, Tuple, Optional
from lxml import etree

NS = {
    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
    'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
    'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
}

for prefix, uri in NS.items():
    etree.register_namespace(prefix, uri)


class TiptapToODT:
    """Converts TipTap editor JSON to ODT XML."""

    def __init__(self):
        self.styles_needed = set()  # Track which styles are needed

    def convert(self, tiptap_json: str) -> Tuple[List[str], str]:
        """
        Convert TipTap JSON to list of ODT XML fragments.

        Args:
            tiptap_json: JSON string from TipTap editor

        Returns:
            (list of XML fragment strings, style definitions CSS-like)
        """
        data = json.loads(tiptap_json)

        if not isinstance(data, dict) or 'content' not in data:
            raise ValueError("Invalid TipTap format: missing 'content' field")

        fragments = []
        for node in data.get('content', []):
            frag = self._convert_node(node)
            if frag:
                fragments.append(frag)

        return fragments, self._generate_styles()

    def _convert_node(self, node: Dict) -> Optional[str]:
        """Convert a single TipTap node to ODT XML."""
        node_type = node.get('type')

        if node_type == 'paragraph':
            return self._convert_paragraph(node)
        elif node_type == 'bulletList':
            return self._convert_bullet_list(node)
        elif node_type == 'orderedList':
            return self._convert_ordered_list(node)
        elif node_type == 'table':
            return self._convert_table(node)
        elif node_type == 'hardBreak':
            return f'<text:line-break xmlns:text="{NS["text"]}" />'
        else:
            return None

    def _convert_paragraph(self, node: Dict) -> str:
        """Convert paragraph with inline formatting."""
        para = etree.Element('{' + NS['text'] + '}p')
        para.set('{' + NS['text'] + '}style-name', 'Textbody')

        for child in node.get('content', []):
            if child.get('type') == 'text':
                span = self._create_text_span(child)
                para.append(span)
            elif child.get('type') == 'hardBreak':
                br = etree.Element('{' + NS['text'] + '}line-break')
                para.append(br)

        # If empty, add placeholder
        if len(para) == 0:
            span = etree.Element('{' + NS['text'] + '}span')
            para.append(span)

        return etree.tostring(para, encoding='unicode')

    def _create_text_span(self, text_node: Dict) -> etree._Element:
        """Create text span with marks (bold, italic, underline)."""
        text = text_node.get('text', '')
        marks = text_node.get('marks', [])

        # Determine style name based on marks
        style_name = self._get_style_name(marks)

        if style_name:
            self.styles_needed.add(style_name)

        span = etree.Element('{' + NS['text'] + '}span')
        if style_name:
            span.set('{' + NS['text'] + '}style-name', style_name)

        text_elem = etree.Element('{' + NS['text'] + '}t')
        text_elem.text = text
        span.append(text_elem)

        return span

    def _get_style_name(self, marks: List[Dict]) -> Optional[str]:
        """Generate style name from marks."""
        mark_types = sorted([m.get('type') for m in marks if m.get('type')])

        if not mark_types:
            return None

        return 'T_' + '_'.join(mark_types)

    def _convert_bullet_list(self, node: Dict) -> str:
        """Convert bullet list."""
        return self._convert_list(node, list_type='bullet')

    def _convert_ordered_list(self, node: Dict) -> str:
        """Convert ordered list."""
        return self._convert_list(node, list_type='ordered')

    def _convert_list(self, node: Dict, list_type: str) -> str:
        """Convert generic list (bullet or ordered)."""
        list_elem = etree.Element('{' + NS['text'] + '}list')

        # Set list style
        style_name = 'BulletList' if list_type == 'bullet' else 'OrderedList'
        list_elem.set('{' + NS['text'] + '}style-name', style_name)

        for child in node.get('content', []):
            if child.get('type') == 'listItem':
                list_item = self._convert_list_item(child)
                list_elem.append(list_item)

        return etree.tostring(list_elem, encoding='unicode')

    def _convert_list_item(self, node: Dict) -> etree._Element:
        """Convert list item."""
        list_item = etree.Element('{' + NS['text'] + '}list-item')

        for child in node.get('content', []):
            if child.get('type') == 'paragraph':
                para_xml = self._convert_paragraph(child)
                para = etree.fromstring(para_xml)
                list_item.append(para)

        return list_item

    def _convert_table(self, node: Dict) -> str:
        """Convert table."""
        table = etree.Element('{' + NS['table'] + '}table')
        table.set('{' + NS['table'] + '}name', 'Table1')
        table.set('{' + NS['table'] + '}style-name', 'Table1')

        rows = node.get('content', [])
        if not rows:
            return etree.tostring(table, encoding='unicode')

        # Get number of columns from first row
        first_row = rows[0] if rows else {}
        num_cols = len(first_row.get('content', []))

        # Add column definitions
        col_width = str(9144 // max(1, num_cols))  # Proportional width
        for _ in range(num_cols):
            col = etree.Element('{' + NS['table'] + '}table-column')
            col.set('{' + NS['table'] + '}style-name', 'Table1.A')
            table.append(col)

        # Add rows
        for row_idx, row in enumerate(rows):
            if row.get('type') != 'tableRow':
                continue
            table_row = etree.Element('{' + NS['table'] + '}table-row')
            table_row.set('{' + NS['table'] + '}style-name', 'Table1.1')

            for cell_idx, cell in enumerate(row.get('content', [])):
                if cell.get('type') != 'tableCell':
                    continue

                table_cell = etree.Element('{' + NS['table'] + '}table-cell')
                table_cell.set('{' + NS['table'] + '}style-name', 'Table1.A1')

                for child in cell.get('content', []):
                    if child.get('type') == 'paragraph':
                        para_xml = self._convert_paragraph(child)
                        para = etree.fromstring(para_xml)
                        table_cell.append(para)

                table_row.append(table_cell)

            table.append(table_row)

        return etree.tostring(table, encoding='unicode')

    def _generate_styles(self) -> str:
        """Generate style definitions for automatic styles."""
        styles = []

        for style_name in sorted(self.styles_needed):
            parts = style_name[2:].split('_')  # Remove 'T_' prefix

            properties = []
            if 'bold' in parts:
                properties.append('<style:text-properties fo:font-weight="bold"/>')
            if 'italic' in parts:
                properties.append('<style:text-properties fo:font-style="italic"/>')
            if 'underline' in parts:
                properties.append('<style:text-properties style:text-underline-style="solid"/>')

            props_xml = ''.join(properties) if properties else ''
            styles.append(f'''<style:style style:name="{style_name}" style:family="text">
  {props_xml}
</style:style>''')

        return '\n'.join(styles)
