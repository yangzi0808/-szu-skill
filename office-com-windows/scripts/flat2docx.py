# -*- coding: utf-8 -*-
"""Rebuild a real .docx package from Word "Flat OPC" XML.

Why this exists: a .doc file cannot reliably be converted by asking Word to
SaveAs - on this machine that call hangs whenever the target path is inside the
user profile, and even outside it the conversion is unnecessary. Word's
Document.WordOpenXML property returns the complete document (all parts,
including media) as Flat OPC XML, and this script turns that XML back into a
normal .docx that python-docx can open. Word's Flat OPC omits
[Content_Types].xml, so it is synthesised from each part's contentType.

Usage:
    python flat2docx.py input_flat.xml output.docx
"""
import base64
import os
import sys
import zipfile

from lxml import etree

PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def convert(src, dst):
    root = etree.parse(src).getroot()
    parts = root.findall("{%s}part" % PKG)
    if not parts:
        raise SystemExit("no pkg:part elements found - is this Flat OPC XML?")

    entries = []
    for part in parts:
        name = part.get("{%s}name" % PKG).lstrip("/")
        ct = part.get("{%s}contentType" % PKG)
        compression = part.get("{%s}compression" % PKG)
        binary = part.find("{%s}binaryData" % PKG)
        xmldata = part.find("{%s}xmlData" % PKG)
        if binary is not None:
            data = base64.b64decode("".join((binary.text or "").split()))
        elif xmldata is not None:
            kids = [k for k in xmldata if isinstance(k.tag, str)]
            payload = etree.tostring(kids[0])
            data = (
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                + payload
            )
        else:
            data = b""
        entries.append((name, ct, compression, data))

    ct_root = etree.Element("{%s}Types" % CT_NS, nsmap={None: CT_NS})
    for ext, ctype in (
        ("rels", "application/vnd.openxmlformats-package.relationships+xml"),
        ("xml", "application/xml"),
    ):
        etree.SubElement(
            ct_root, "{%s}Default" % CT_NS, Extension=ext, ContentType=ctype
        )
    for name, ct, _c, _d in entries:
        if ct:
            etree.SubElement(
                ct_root, "{%s}Override" % CT_NS, PartName="/" + name, ContentType=ct
            )
    ct_xml = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8")

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct_xml)
        for name, _ct, compression, data in entries:
            zi = zipfile.ZipInfo(name)
            if compression == "store":
                zi.compress_type = zipfile.ZIP_STORED
            z.writestr(zi, data)
    print("wrote", dst, os.path.getsize(dst), "parts", len(entries) + 1)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit("usage: flat2docx.py input_flat.xml output.docx")
    convert(sys.argv[1], sys.argv[2])
