from __future__ import annotations

import json
import sys
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ET.register_namespace("a", A_NS)
ET.register_namespace("p", P_NS)
ET.register_namespace("r", R_NS)

NS = {"p": P_NS, "a": A_NS, "r": R_NS}


def qn(namespace: str, tag: str) -> str:
    return f"{{{namespace}}}{tag}"


def extract_shape_text(shape: ET.Element) -> str:
    parts = []
    for node in shape.findall(".//a:t", NS):
        if node.text:
            parts.append(node.text)
    return " ".join(parts).strip()


def extract_shape_id(shape: ET.Element) -> str | None:
    node = shape.find("./p:nvSpPr/p:cNvPr", NS)
    if node is None:
        return None
    return node.attrib.get("id")


def build_transition(spec: dict) -> ET.Element:
    transition = ET.Element(qn(P_NS, "transition"), {"spd": "slow"})
    ET.SubElement(transition, qn(P_NS, spec["type"]))
    return transition


def build_timing(shape_ids: list[str]) -> ET.Element:
    timing = ET.Element(qn(P_NS, "timing"))
    build_list = ET.SubElement(timing, qn(P_NS, "bldLst"))
    for idx, shape_id in enumerate(shape_ids):
        ET.SubElement(
            build_list,
            qn(P_NS, "bldP"),
            {
                "spid": str(shape_id),
                "grpId": str(idx),
                "build": "p",
                "advAuto": "250",
            },
        )
    return timing


def update_slide_xml(slide_xml: bytes, slide_meta: dict) -> bytes:
    root = ET.fromstring(slide_xml)

    for child in list(root):
        if child.tag in {qn(P_NS, "transition"), qn(P_NS, "timing")}:
            root.remove(child)

    transition = build_transition(slide_meta["transition"])

    shape_ids: list[str] = []
    for shape in root.findall(".//p:sp", NS):
        shape_text = extract_shape_text(shape)
        if not shape_text:
            continue
        if any(target in shape_text for target in slide_meta["animationTargets"]):
            shape_id = extract_shape_id(shape)
            if shape_id and shape_id not in shape_ids:
                shape_ids.append(shape_id)

    insert_index = len(root)
    for idx, child in enumerate(list(root)):
        if child.tag == qn(P_NS, "clrMapOvr"):
            insert_index = idx + 1
            break

    root.insert(insert_index, transition)
    if shape_ids:
        root.insert(insert_index + 1, build_timing(shape_ids))

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def patch_pptx(pptx_path: Path, metadata_path: Path) -> None:
    slide_metadata = {entry["number"]: entry for entry in json.loads(metadata_path.read_text("utf-8"))}

    with zipfile.ZipFile(pptx_path, "r") as src_zip, tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        tmp_path = Path(tmp.name)
        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as dst_zip:
            for info in src_zip.infolist():
                data = src_zip.read(info.filename)
                if info.filename.startswith("ppt/slides/slide") and info.filename.endswith(".xml"):
                    slide_number = int(info.filename.replace("ppt/slides/slide", "").replace(".xml", ""))
                    if slide_number in slide_metadata:
                        data = update_slide_xml(data, slide_metadata[slide_number])
                dst_zip.writestr(info, data)

    tmp_path.replace(pptx_path)


def main(argv: list[str]) -> int:
    if len(argv) != 3:
        print("Usage: python3 src/postprocess_pptx.py <pptx_path> <metadata_path>")
        return 1

    pptx_path = Path(argv[1])
    metadata_path = Path(argv[2])
    patch_pptx(pptx_path, metadata_path)
    print(f"Patched transitions and build animations in {pptx_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
