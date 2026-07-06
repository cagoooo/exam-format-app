from __future__ import annotations

import argparse
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from app import GENERATED_DIR, PROFILES, build_exam_docx  # noqa: E402


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


EXAMPLES = [
    ("國卷範例(橫式)a4版.docx", "a4-horizontal"),
    ("國卷範例(橫式)b4版.docx", "b4-horizontal"),
    ("國字檢測考試格式(橫式)b4版.docx", "b4-horizontal"),
    ("試卷格式(直式)b4版.docx", "b4-vertical"),
    ("試卷範例(直式)a4版.docx", "a4-vertical"),
]


def attrs(el):
    if el is None:
        return {}
    return {key.split("}")[-1]: value for key, value in el.attrib.items()}


def inspect_docx(path: Path) -> dict:
    with zipfile.ZipFile(path) as zf:
        document = ET.fromstring(zf.read("word/document.xml"))
    sect = document.findall(".//w:sectPr", NS)[-1]
    return {
        "pgSz": attrs(sect.find("w:pgSz", NS)),
        "pgMar": attrs(sect.find("w:pgMar", NS)),
        "cols": attrs(sect.find("w:cols", NS)),
        "tables": len(document.findall(".//w:tbl", NS)),
        "paragraphs": len(document.findall(".//w:p", NS)),
    }


def verify_output(output: Path, profile_key: str) -> list[str]:
    profile = PROFILES[profile_key]
    info = inspect_docx(output)
    errors: list[str] = []
    if info["pgSz"].get("w") != str(profile.width) or info["pgSz"].get("h") != str(profile.height):
        errors.append(f"頁面尺寸不符：{info['pgSz']}")
    for key, expected in profile.margins.items():
        if info["pgMar"].get(key) != str(expected):
            errors.append(f"邊界 {key} 不符：{info['pgMar'].get(key)} != {expected}")
    if info["cols"].get("num") != str(profile.columns):
        errors.append(f"欄數不符：{info['cols']}")
    if info["cols"].get("space") != str(profile.column_space):
        errors.append(f"欄距不符：{info['cols']}")
    return errors


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-dir", default="H:/Word")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    source_dir = Path(args.source_dir)
    GENERATED_DIR.mkdir(exist_ok=True)
    results = []
    exit_code = 0

    for filename, profile_key in EXAMPLES:
        source = source_dir / filename
        output = GENERATED_DIR / f"verify-{profile_key}-{source.stem}.docx"
        if not source.exists():
            results.append({"file": filename, "profile": profile_key, "ok": False, "errors": ["找不到範例檔"]})
            exit_code = 1
            continue
        try:
            stats = build_exam_docx(source, output, PROFILES[profile_key], {})
            errors = verify_output(output, profile_key)
            ok = not errors
            if not ok:
                exit_code = 1
            results.append({"file": filename, "profile": profile_key, "ok": ok, "errors": errors, "stats": stats})
        except Exception as exc:
            results.append({"file": filename, "profile": profile_key, "ok": False, "errors": [str(exc)]})
            exit_code = 1

    if args.json:
        print(json.dumps(results, ensure_ascii=False, indent=2))
    else:
        for item in results:
            mark = "PASS" if item["ok"] else "FAIL"
            print(f"{mark} {item['file']} -> {item['profile']}")
            for error in item["errors"]:
                print(f"  - {error}")
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
