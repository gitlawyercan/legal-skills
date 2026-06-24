#!/usr/bin/env python3
"""Lightweight QA for contract redline DOCX files."""

from __future__ import annotations

import argparse
import json
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def qn(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def inspect_docx(docx_path: Path) -> dict:
    with tempfile.TemporaryDirectory(prefix="contract-redline-qa-") as tmp:
        root = Path(tmp)
        with zipfile.ZipFile(docx_path) as zf:
            zf.extractall(root)
        document_path = root / "word" / "document.xml"
        settings_path = root / "word" / "settings.xml"
        comments_path = root / "word" / "comments.xml"
        rels_path = root / "word" / "_rels" / "document.xml.rels"

        if not document_path.exists():
            raise FileNotFoundError("word/document.xml 不存在")
        doc_root = ET.parse(document_path).getroot()
        ins_count = len(doc_root.findall(".//w:ins", NS))
        del_count = len(doc_root.findall(".//w:del", NS))

        track_revisions = False
        if settings_path.exists():
            settings_root = ET.parse(settings_path).getroot()
            track_revisions = settings_root.find("w:trackRevisions", NS) is not None

        comment_count = 0
        if comments_path.exists():
            comments_root = ET.parse(comments_path).getroot()
            comment_count = len(comments_root.findall(".//w:comment", NS))

        has_comment_rel = False
        if rels_path.exists():
            rels_text = rels_path.read_text(encoding="utf-8", errors="ignore")
            has_comment_rel = "comments.xml" in rels_text

        return {
            "docx": str(docx_path),
            "track_revisions": track_revisions,
            "insertions": ins_count,
            "deletions": del_count,
            "comments": comment_count,
            "has_comment_relationship": has_comment_rel,
        }


def main() -> int:
    parser = argparse.ArgumentParser(description="检查合同红线稿 DOCX 结构")
    parser.add_argument("--docx", required=True, help="待检查 DOCX")
    parser.add_argument("--expect-ins", type=int, default=0, help="至少应有的 w:ins 数量")
    parser.add_argument("--expect-del", type=int, default=0, help="至少应有的 w:del 数量")
    parser.add_argument("--expect-comments", type=int, default=0, help="至少应有的批注数量")
    parser.add_argument("--json", action="store_true", help="输出 JSON")
    args = parser.parse_args()

    result = inspect_docx(Path(args.docx).expanduser().resolve())
    failures = []
    if not result["track_revisions"]:
        failures.append("word/settings.xml 未启用 w:trackRevisions")
    if result["insertions"] < args.expect_ins:
        failures.append(f"w:ins 数量不足：{result['insertions']} < {args.expect_ins}")
    if result["deletions"] < args.expect_del:
        failures.append(f"w:del 数量不足：{result['deletions']} < {args.expect_del}")
    if result["comments"] < args.expect_comments:
        failures.append(f"批注数量不足：{result['comments']} < {args.expect_comments}")
    if result["comments"] and not result["has_comment_relationship"]:
        failures.append("存在 comments.xml 但 document.xml.rels 未包含 comments.xml 关系")

    result["status"] = "PASS" if not failures else "FAIL"
    result["failures"] = failures
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"redline QA: {result['status']}")
        for key in ("track_revisions", "insertions", "deletions", "comments"):
            print(f"- {key}: {result[key]}")
        for failure in failures:
            print(f"- failure: {failure}")
    return 0 if not failures else 1


if __name__ == "__main__":
    raise SystemExit(main())
