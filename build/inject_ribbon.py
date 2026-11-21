# -*- coding: utf-8 -*-
"""
Injects customUI/customUI14.xml into an Excel .xlam / .xlsm OOXML archive cleanly.
Updates _rels/.rels and [Content_Types].xml if needed without disturbing vbaProject.bin.
"""
import os
import sys
import zipfile
import xml.etree.ElementTree as ET

RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RIBBON_REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
RIBBON_TARGET = "customUI/customUI14.xml"


def inject_ribbon(xlam_path: str, ribbon_xml_path: str):
    if not os.path.exists(xlam_path):
        raise FileNotFoundError(xlam_path)
    if not os.path.exists(ribbon_xml_path):
        raise FileNotFoundError(ribbon_xml_path)

    with open(ribbon_xml_path, "rb") as f:
        ribbon_bytes = f.read()

    # Validate XML well-formedness before injection
    ET.fromstring(ribbon_bytes)

    tmp_path = xlam_path + ".tmp"
    with zipfile.ZipFile(xlam_path, "r") as zin:
        rels_bytes = zin.read("_rels/.rels")
        ct_bytes = zin.read("[Content_Types].xml")

        # Update _rels/.rels
        ET.register_namespace("", RELS_NS)
        rels_root = ET.fromstring(rels_bytes)
        exists = False
        max_id = 0
        for rel in rels_root.findall(f"{{{RELS_NS}}}Relationship"):
            rid = rel.attrib.get("Id", "")
            if rid.startswith("rId") and rid[3:].isdigit():
                max_id = max(max_id, int(rid[3:]))
            if rel.attrib.get("Type") == RIBBON_REL_TYPE:
                rel.attrib["Target"] = RIBBON_TARGET
                exists = True
        if not exists:
            ET.SubElement(
                rels_root,
                f"{{{RELS_NS}}}Relationship",
                {
                    "Id": f"rId{max_id + 1}",
                    "Type": RIBBON_REL_TYPE,
                    "Target": RIBBON_TARGET,
                },
            )
        new_rels_bytes = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)

        # Ensure xml default content type exists in [Content_Types].xml
        ET.register_namespace("", CT_NS)
        ct_root = ET.fromstring(ct_bytes)
        has_xml_ext = any(
            el.attrib.get("Extension", "").lower() == "xml"
            for el in ct_root.findall(f"{{{CT_NS}}}Default")
        )
        if not has_xml_ext:
            ET.SubElement(
                ct_root,
                f"{{{CT_NS}}}Default",
                {"Extension": "xml", "ContentType": "application/xml"},
            )
        new_ct_bytes = ET.tostring(ct_root, encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename in ("_rels/.rels", "[Content_Types].xml", RIBBON_TARGET):
                    continue
                zout.writestr(item, zin.read(item.filename))
            zout.writestr("_rels/.rels", new_rels_bytes)
            zout.writestr("[Content_Types].xml", new_ct_bytes)
            zout.writestr(RIBBON_TARGET, ribbon_bytes)

    os.replace(tmp_path, xlam_path)
    print(f"Injected {RIBBON_TARGET} cleanly into {xlam_path}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python inject_ribbon.py <xlam_path> <customUI14.xml>")
        sys.exit(1)
    inject_ribbon(sys.argv[1], sys.argv[2])
