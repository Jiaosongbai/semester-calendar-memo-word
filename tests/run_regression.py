#!/usr/bin/env python3
"""Deterministic regression checks for the calendar DOCX generator."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
GENERATOR = ROOT / "scripts/generate_notebook.py"
OUT = ROOT / "tests/output"
W = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def run_case(name, spec, expected_tables, expected_footnotes, required_text=()):
    spec_path = OUT / f"{name}.json"
    docx_path = OUT / f"{name}.docx"
    spec_path.write_text(json.dumps(spec, ensure_ascii=False, indent=2), encoding="utf-8")
    subprocess.run([sys.executable, str(GENERATOR), str(spec_path), "-o", str(docx_path)], check=True)
    with ZipFile(docx_path) as archive:
        document = etree.fromstring(archive.read("word/document.xml"))
        assert len(document.xpath("//w:tbl", namespaces=W)) == expected_tables
        assert not document.xpath('//w:trHeight[@w:hRule="exact"]', namespaces=W)
        assert set(document.xpath("//w:trHeight/@w:hRule", namespaces=W)) == {"atLeast"}
        refs = document.xpath("//w:footnoteReference/@w:id", namespaces=W)
        assert len(refs) == expected_footnotes
        body_text = "".join(document.xpath("//w:t/text()", namespaces=W))
        for text in required_text:
            assert text in body_text, text
        if expected_footnotes:
            footnotes = etree.fromstring(archive.read("word/footnotes.xml"))
            ids = footnotes.xpath('//w:footnote[not(@w:type)]/@w:id', namespaces=W)
            bottom_refs = footnotes.xpath('//w:footnote[not(@w:type)]//w:footnoteRef', namespaces=W)
            assert refs == ids and len(bottom_refs) == expected_footnotes
        else:
            assert "word/footnotes.xml" not in archive.namelist()
    return docx_path


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    single = {
        "title": "单学期稳定性测试",
        "subtitle": "仅生成已提供的一个学期",
        "sections": [{"name": "秋季学期", "start": "2026-09-07", "weeks": 6, "exam_from_week": 6}],
        "statutory_holidays": [{"name": "国庆节", "start": "2026-10-01", "end": "2026-10-03"}],
        "calendar_day_markers": [{"name": "校庆日", "date": "2026-09-25"}],
        "school_breaks": [{"name": "校庆假期", "start": "2026-09-26", "end": "2026-09-27"}],
        "activities": [{"text": "讲座：10月举办人工智能专题讲座（日期待定）", "precision": "month", "anchor_date": "2026-10-08", "duration_rank": 1}]
    }
    no_events = {
        "title": "无活动稳定性测试",
        "sections": [{"name": "春季学期", "start": "2027-02-22", "weeks": 4}],
        "activities": []
    }
    day_event = {
        "title": "明确日期活动稳定性测试",
        "sections": [{"name": "短学期", "start": "2027-07-05", "weeks": 2}],
        "activities": [{"text": "毕业典礼：2027年7月9日", "precision": "day", "anchor_date": "2027-07-09", "duration_rank": 1}]
    }
    paths = [
        run_case("single-semester", single, 6, 1, ("国庆节（法定放假）", "校庆日（校历标记）", "校庆假期（学校假期）")),
        run_case("no-activities", no_events, 4, 0),
        run_case("confirmed-day-activity", day_event, 2, 1),
    ]
    invalid = {"title": "信息不足测试", "sections": [{"name": "秋季学期", "weeks": 18}]}
    bad_path = OUT / "insufficient.json"
    bad_path.write_text(json.dumps(invalid, ensure_ascii=False), encoding="utf-8")
    result = subprocess.run([sys.executable, str(GENERATOR), str(bad_path), "-o", str(OUT / "must-not-exist.docx")], capture_output=True, text=True)
    assert result.returncode == 2 and "缺少字段：start" in result.stderr
    assert not (OUT / "must-not-exist.docx").exists()
    print("PASS")
    for path in paths:
        print(path)
    print("信息不足测试：已正确拒绝生成")


if __name__ == "__main__":
    main()
