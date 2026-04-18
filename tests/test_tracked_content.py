"""Bug 5: `_with_track_changes` variants for content-insertion tools.

The plain `add_paragraph` / `add_heading` / `insert_line_or_paragraph_near_text`
tools insert content without wrapping in `<w:ins>`. Legal/medical/editorial
review workflows need the new tracked variants so every author change
shows up as a tracked insertion.
"""
import asyncio
import json
import zipfile

import pytest

from word_document_server.tools.track_changes_tools import (
    add_paragraph_with_track_changes,
    add_heading_with_track_changes,
    insert_line_or_paragraph_near_text_with_track_changes,
)


def _run(coro):
    return asyncio.run(coro)


def _parse(result):
    return json.loads(result)


def _document_xml(path: str) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8")


class TestTrackedAddParagraph:
    def test_add_paragraph_with_track_changes_wraps_in_w_ins(self, make_docx):
        path = make_docx(paragraphs=["Existing body"])
        data = _parse(_run(add_paragraph_with_track_changes(
            path, "New tracked paragraph", author="Tester"
        )))
        assert data["success"] is True, data
        xml = _document_xml(path)
        assert "New tracked paragraph" in xml
        assert "<w:ins " in xml
        # The author must live inside the w:ins attributes.
        ins_tag = xml[xml.find("<w:ins "):xml.find("<w:ins ") + 200]
        assert 'w:author="Tester"' in ins_tag, ins_tag


class TestTrackedAddHeading:
    def test_add_heading_with_track_changes_wraps_in_w_ins(self, make_docx):
        path = make_docx(paragraphs=["Body 1", "Body 2"])
        data = _parse(_run(add_heading_with_track_changes(
            path, "Tracked Heading", level=2, author="Tester"
        )))
        assert data["success"] is True, data
        xml = _document_xml(path)
        assert "Tracked Heading" in xml
        assert "<w:ins " in xml


class TestTrackedInsertLine:
    def test_insert_line_after_anchor_wraps_in_w_ins(self, make_docx):
        path = make_docx(paragraphs=["anchor line", "tail"])
        data = _parse(_run(insert_line_or_paragraph_near_text_with_track_changes(
            path, anchor_text="anchor line",
            text_to_insert="new tracked line",
            position="after", author="Tester"
        )))
        assert data["success"] is True, data
        xml = _document_xml(path)
        assert "new tracked line" in xml
        assert "<w:ins " in xml

    def test_insert_line_missing_anchor_reports_not_found(self, make_docx):
        path = make_docx(paragraphs=["only line"])
        data = _parse(_run(insert_line_or_paragraph_near_text_with_track_changes(
            path, anchor_text="nowhere",
            text_to_insert="x",
            position="after", author="Tester"
        )))
        assert data["success"] is False
        assert "not found" in data["error"].lower()
