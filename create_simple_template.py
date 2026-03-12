#!/usr/bin/env python3
"""Crea template ODT minimale con solo P_testo_obj."""
import zipfile
from pathlib import Path

def create_simple_odt():
    output_path = 'template/ASL_Template_Simple.odt'
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    mimetype = 'application/vnd.oasis.opendocument.text'

    meta_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                      xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
                      xmlns:dc="http://purl.org/dc/elements/1.1/">
  <office:meta>
    <meta:creation-date>2024-01-15T10:00:00</meta:creation-date>
    <dc:creator>System</dc:creator>
    <dc:title>ASL Template Delibera Simple</dc:title>
  </office:meta>
</office:document-meta>'''

    styles_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                        xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
                        xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
                        xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
  <office:styles>
    <style:style style:name="Textbody" style:family="paragraph">
      <style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.3cm" fo:line-height="1.5"/>
    </style:style>
    <style:style style:name="Heading1" style:family="paragraph">
      <style:paragraph-properties fo:margin-top="0.42cm" fo:margin-bottom="0.21cm"/>
      <style:text-properties fo:font-size="200%" fo:font-weight="bold"/>
    </style:style>
  </office:styles>
  <office:automatic-styles/>
</office:document-styles>'''

    content_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                         xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
                         xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
                         xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:scripts/>
  <office:automatic-styles/>
  <office:body>
    <office:text>
      <text:h text:outline-level="1" text:style-name="Heading1">DELIBERA</text:h>
      <text:p text:style-name="Textbody">P_testo_obj</text:p>
    </office:text>
  </office:body>
</office:document-content>'''

    manifest_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.text" manifest:full-path="/"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml"/>
</manifest:manifest>'''

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('mimetype', mimetype, compress_type=zipfile.ZIP_STORED)
        zf.writestr('meta.xml', meta_xml)
        zf.writestr('styles.xml', styles_xml)
        zf.writestr('content.xml', content_xml)
        zf.writestr('META-INF/manifest.xml', manifest_xml)

    print(f"[OK] Template creato: {output_path}")

if __name__ == '__main__':
    create_simple_odt()
