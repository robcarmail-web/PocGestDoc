#!/usr/bin/env python3
"""
Crea un template ODT minimale per il POC.
Il template contiene placeholder semplici e ricchi per testare l'iniezione.
"""
import zipfile
import os
from pathlib import Path

def create_minimal_odt_template(output_path='template/ASL_Template_Delibera.odt'):
    """Crea un template ODT minimale con placeholder."""

    # Ensure output directory exists
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    # Contenuto XML minimale per ODT
    mimetype = 'application/vnd.oasis.opendocument.text'

    # Meta.xml
    meta_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                      xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
                      xmlns:dc="http://purl.org/dc/elements/1.1/">
  <office:meta>
    <meta:creation-date>2024-01-15T10:00:00</meta:creation-date>
    <dc:creator>System</dc:creator>
    <dc:title>ASL Template Delibera</dc:title>
  </office:meta>
</office:document-meta>'''

    # Styles.xml - stili di base
    styles_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                        xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
                        xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
                        xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
                        xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
                        xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">
  <office:styles>
    <style:style style:name="Textbody" style:family="paragraph">
      <style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.3cm"/>
    </style:style>
    <style:style style:name="Heading1" style:family="paragraph">
      <style:paragraph-properties fo:margin-top="0.42cm" fo:margin-bottom="0.21cm"/>
      <style:text-properties fo:font-size="200%"/>
    </style:style>
    <style:style style:name="Table1" style:family="table">
      <style:table-properties/>
    </style:style>
    <style:style style:name="Table1.A" style:family="table-column"/>
    <style:style style:name="Table1.1" style:family="table-row"/>
    <style:style style:name="Table1.A1" style:family="table-cell"/>
  </office:styles>
  <office:automatic-styles>
    <style:style style:name="P1" style:family="paragraph" style:parent-style-name="Textbody"/>
  </office:automatic-styles>
</office:document-styles>'''

    # Content.xml - il contenuto principale con placeholder
    content_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                         xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
                         xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
                         xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
                         xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
  <office:scripts/>
  <office:automatic-styles/>
  <office:body>
    <office:text>
      <text:p text:style-name="Textbody">
        <text:span>DELIBERA N. </text:span>
        <text:span>$$numeroproposta$$</text:span>
      </text:p>

      <text:p text:style-name="Textbody">
        <text:span>Data: </text:span>
        <text:span>$$dataproposta$$</text:span>
      </text:p>

      <text:p text:style-name="Heading1">$$oggetto$$</text:p>

      <text:h text:outline-level="1">RELAZIONE ISTRUTTORIA</text:h>
      <table:table table:name="IstruttoriaTable" table:style-name="Table1">
        <table:table-column table:style-name="Table1.A" table:number-columns-repeated="1"/>
        <table:table-row table:style-name="Table1.1">
          <table:table-cell table:style-name="Table1.A1">
            <text:p>I_testo_obj</text:p>
          </table:table-cell>
        </table:table-row>
      </table:table>

      <text:h text:outline-level="1">PROPOSTA</text:h>
      <table:table table:name="PropostaTable" table:style-name="Table1">
        <table:table-column table:style-name="Table1.A" table:number-columns-repeated="1"/>
        <table:table-row table:style-name="Table1.1">
          <table:table-cell table:style-name="Table1.A1">
            <text:p>P_testo_obj</text:p>
          </table:table-cell>
        </table:table-row>
      </table:table>

      <text:h text:outline-level="1">DELIBERA</text:h>
      <table:table table:name="DeliberaTable" table:style-name="Table1">
        <table:table-column table:style-name="Table1.A" table:number-columns-repeated="1"/>
        <table:table-row table:style-name="Table1.1">
          <table:table-cell table:style-name="Table1.A1">
            <text:p>P_testo_obj</text:p>
          </table:table-cell>
        </table:table-row>
      </table:table>

      <text:p text:style-name="Textbody">
        <text:span>Estensore: $$EstensoreNome$$</text:span>
      </text:p>

      <text:p text:style-name="Textbody">
        <text:span>Direttore Generale: $$DirGenNome$$</text:span>
      </text:p>
    </office:text>
  </office:body>
</office:document-content>'''

    # Manifest.xml
    manifest_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.text" manifest:full-path="/"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml"/>
</manifest:manifest>'''

    # Crea il file ZIP
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        # Mimetype must be first and uncompressed
        zf.writestr('mimetype', mimetype, compress_type=zipfile.ZIP_STORED)

        # Aggiungi gli XML
        zf.writestr('meta.xml', meta_xml)
        zf.writestr('styles.xml', styles_xml)
        zf.writestr('content.xml', content_xml)
        zf.writestr('META-INF/manifest.xml', manifest_xml)

    print(f"[OK] Template created: {output_path}")
    return output_path

if __name__ == '__main__':
    create_minimal_odt_template()
