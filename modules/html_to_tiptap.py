"""
Converts HTML (from mammoth) to TipTap JSON format.
Handles paragraphs, text formatting, lists, tables, and headings.
"""
import json
from html.parser import HTMLParser
from typing import List, Dict, Optional, Tuple


class HTMLToTiptapParser(HTMLParser):
    """Parses HTML and builds TipTap node structure."""

    def __init__(self):
        super().__init__()
        self.root = {"type": "doc", "content": []}
        self.current_block = None
        self.current_para = None
        self.text_marks = []
        self.list_stack = []
        self.table_state = None
        self.heading_level = 0

    def handle_starttag(self, tag: str, attrs: List[Tuple]):
        """Handle opening tags."""
        attrs_dict = dict(attrs)

        if tag in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            self.heading_level = int(tag[1])

        elif tag == 'p':
            self._flush_current()
            self.current_para = {"type": "paragraph", "content": []}

        elif tag == 'br':
            if self.current_para is None:
                self.current_para = {"type": "paragraph", "content": []}
            self.current_para["content"].append({"type": "hardBreak"})

        elif tag == 'ul':
            self._flush_current()
            self.list_stack.append({"type": "bulletList", "content": []})

        elif tag == 'ol':
            self._flush_current()
            self.list_stack.append({"type": "orderedList", "content": []})

        elif tag == 'li':
            if self.list_stack:
                self.current_para = {"type": "paragraph", "content": []}

        elif tag == 'strong' or tag == 'b':
            self.text_marks.append({"type": "bold"})

        elif tag == 'em' or tag == 'i':
            self.text_marks.append({"type": "italic"})

        elif tag == 'u':
            self.text_marks.append({"type": "underline"})

        elif tag == 'table':
            self._flush_current()
            self.table_state = {"rows": []}

        elif tag == 'tr':
            if self.table_state is not None:
                self.table_state["current_row"] = {"type": "tableRow", "content": []}

        elif tag in ['td', 'th']:
            if self.table_state is not None:
                self.current_para = {"type": "paragraph", "content": []}

    def handle_endtag(self, tag: str):
        """Handle closing tags."""
        if tag in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            if self.current_para:
                self.root["content"].append(self.current_para)
            self.current_para = None
            self.heading_level = 0

        elif tag == 'p':
            if self.current_para:
                self.root["content"].append(self.current_para)
                self.current_para = None

        elif tag == 'ul' or tag == 'ol':
            if self.list_stack:
                list_node = self.list_stack.pop()
                if self.list_stack:
                    # Nested list
                    self.list_stack[-1]["content"].append(list_node)
                else:
                    # Top-level list
                    self.root["content"].append(list_node)

        elif tag == 'li':
            if self.current_para and self.list_stack:
                list_item = {"type": "listItem", "content": [self.current_para]}
                self.list_stack[-1]["content"].append(list_item)
                self.current_para = None

        elif tag in ['td', 'th']:
            if self.current_para is not None and self.table_state is not None:
                cell = {"type": "tableCell", "content": [self.current_para]}
                self.table_state["current_row"]["content"].append(cell)
                self.current_para = None

        elif tag == 'tr':
            if self.table_state is not None and "current_row" in self.table_state:
                self.table_state["rows"].append(self.table_state["current_row"])
                del self.table_state["current_row"]

        elif tag == 'table':
            if self.table_state is not None:
                table_node = {
                    "type": "table",
                    "content": self.table_state.get("rows", [])
                }
                self.root["content"].append(table_node)
                self.table_state = None

        elif tag in ['strong', 'b']:
            if self.text_marks and self.text_marks[-1].get('type') == 'bold':
                self.text_marks.pop()

        elif tag in ['em', 'i']:
            if self.text_marks and self.text_marks[-1].get('type') == 'italic':
                self.text_marks.pop()

        elif tag == 'u':
            if self.text_marks and self.text_marks[-1].get('type') == 'underline':
                self.text_marks.pop()

    def handle_data(self, data: str):
        """Handle text content."""
        if not data.strip():
            return

        # Ensure we have a paragraph to add text to
        if self.current_para is None and not self.table_state:
            self.current_para = {"type": "paragraph", "content": []}

        if self.current_para is not None:
            text_node = {
                "type": "text",
                "text": data,
                "marks": list(self.text_marks)  # Copy current marks
            }
            self.current_para["content"].append(text_node)

    def _flush_current(self):
        """Flush current paragraph if any."""
        if self.current_para and self.current_para["content"]:
            self.root["content"].append(self.current_para)
        self.current_para = None

    def get_result(self) -> dict:
        """Get final TipTap document."""
        self._flush_current()
        return self.root


class HTMLToTiptap:
    """Converts HTML to TipTap JSON."""

    @staticmethod
    def convert(html: str) -> str:
        """
        Convert HTML to TipTap JSON.

        Args:
            html: HTML string (typically from mammoth)

        Returns:
            JSON string of TipTap document
        """
        parser = HTMLToTiptapParser()
        parser.feed(html)
        result = parser.get_result()
        return json.dumps(result)

    @staticmethod
    def convert_to_dict(html: str) -> dict:
        """Convert HTML to TipTap dict."""
        return json.loads(HTMLToTiptap.convert(html))
