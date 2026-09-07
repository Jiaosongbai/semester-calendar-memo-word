#!/usr/bin/env python3
"""Run the three release-gate fixtures before publishing the skill."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
GENERATOR = ROOT / "scripts/generate_notebook.py"
OUT = ROOT / "tests/output/final-acceptance"
W = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


CASES = (
    {
        "name": "pdf-input",
        "spec": ROOT / "references/spec-example.json",
        "tables": 40,
        "footnotes": 9,
        "required": ("第一学期", "第二学期", "暑期学期", "寒假：2027年1月18日至2月21日"),
    },
    {
        "name": "image-input",
        "spec": ROOT / "regression/zhejiang-image-spec.json",
        "tables": 18,
        "footnotes": 4,
        "required": ("秋学期", "冬学期", "国庆节（法定放假）", "元旦放假（调休安排另行通知）（校历标记）"),
    },
    {
        "name": "text-input",
        "spec": ROOT / "regression/xiaohongshu-text-spec.json",
        "tables": 17,
        "footnotes": 0,
        "required": ("第十四周", "第十五周（考试周）", "生日｜愿新岁常安，所愿皆如意"),
    },
)


def inspect_docx(path: Path, case: dict) -> None:
    with ZipFile(path) as archive:
        assert archive.testzip() is None
        names = set(archive.namelist())
        document = etree.fromstring(archive.read("word/document.xml"))
        text = "".join(document.xpath("//w:t/text()", namespaces=W))
        assert len(document.xpath("//w:tbl", namespaces=W)) == case["tables"]
        assert not document.xpath('//w:trHeight[@w:hRule="exact"]', namespaces=W)
        assert not document.xpath("//w:txbxContent", namespaces=W)
        for required in case["required"]:
            assert required in text, (case["name"], required)
        refs = document.xpath("//w:footnoteReference/@w:id", namespaces=W)
        assert len(refs) == case["footnotes"]
        if refs:
            footnotes = etree.fromstring(archive.read("word/footnotes.xml"))
            ids = footnotes.xpath('//w:footnote[not(@w:type)]/@w:id', namespaces=W)
            bottom_refs = footnotes.xpath('//w:footnote[not(@w:type)]//w:footnoteRef', namespaces=W)
            assert refs == ids
            assert len(bottom_refs) == len(refs)
        else:
            assert "word/footnotes.xml" not in names
        if case["name"] == "text-input":
            runs = document.xpath('//w:r[w:t[contains(.,"生日｜愿新岁常安，所愿皆如意")]]', namespaces=W)
            assert len(runs) == 1
            assert runs[0].xpath('./w:rPr/w:i[@w:val="0"]', namespaces=W)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for case in CASES:
        json.loads(case["spec"].read_text(encoding="utf-8"))
        output = OUT / f'{case["name"]}.docx'
        subprocess.run([sys.executable, str(GENERATOR), str(case["spec"]), "--output", str(output)], check=True)
        inspect_docx(output, case)
        print(f'PASS {case["name"]}: {output}')
    print("FINAL ACCEPTANCE PASS")


if __name__ == "__main__":
    main()
