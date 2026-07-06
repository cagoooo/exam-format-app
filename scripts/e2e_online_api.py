from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import requests


DEFAULT_API = "https://exam-format-api-142975838924.asia-east1.run.app"


def pick_sample(source_dir: Path) -> Path:
    files = sorted(source_dir.glob("*.docx"))
    if not files:
        raise FileNotFoundError(f"找不到測試 DOCX：{source_dir}")
    for file in files:
        if "b4" in file.name.lower():
            return file
    return files[0]


def main() -> int:
    parser = argparse.ArgumentParser(description="線上 API 端到端測試")
    parser.add_argument("--api", default=DEFAULT_API)
    parser.add_argument("--source-dir", default="H:/Word")
    parser.add_argument("--sample")
    args = parser.parse_args()

    api = args.api.rstrip("/")
    sample = Path(args.sample) if args.sample else pick_sample(Path(args.source_dir))

    health = requests.get(f"{api}/version.json", timeout=30)
    health.raise_for_status()

    with sample.open("rb") as file:
        analyze = requests.post(
            f"{api}/api/analyze",
            files={"file": (sample.name, file, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
            timeout=120,
        )
    analyze.raise_for_status()
    analyze_json = analyze.json()
    if not analyze_json.get("ok"):
        raise RuntimeError(f"分析失敗：{analyze_json}")

    with sample.open("rb") as file:
        convert = requests.post(
            f"{api}/api/convert",
            data={"profile": "b4-horizontal", "target_pages": "2", "subject": "國語", "teacher": "阿凱老師"},
            files={"file": (sample.name, file, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
            timeout=300,
        )
    convert.raise_for_status()
    convert_json = convert.json()
    if not convert_json.get("ok"):
        raise RuntimeError(f"轉換失敗：{convert_json}")
    for key in ("download", "pdf"):
        url = convert_json.get(key)
        if not url:
            raise RuntimeError(f"缺少 {key} 下載連結：{convert_json}")
        res = requests.get(f"{api}{url}", timeout=120)
        res.raise_for_status()
        if len(res.content) < 1024:
            raise RuntimeError(f"{key} 檔案太小，可能輸出失敗")

    print(json.dumps({
        "ok": True,
        "sample": sample.name,
        "blocks": analyze_json.get("blocks"),
        "pdf_pages": convert_json.get("pdf_pages"),
        "job_id": convert_json.get("job_id"),
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
