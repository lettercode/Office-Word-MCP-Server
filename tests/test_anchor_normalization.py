"""Tests for cross-run anchor matching.

docx-editor matches anchor text within a single run. `find_text_in_document`
and `count_tracked_matches` match across runs (paragraph-text). The gap
means an anchor that appears in the paragraph text — e.g. "Hello World"
split across two runs — silently fails in add_comment,
replace_with_track_changes, and the insert_{before,after}_with_track_changes
tools. These tests assert those tools succeed on cross-run anchors.
"""
import asyncio
import json

import pytest

from word_document_server.tools.comment_management_tools import add_comment
from word_document_server.tools.comment_tools import get_all_comments
from word_document_server.tools.track_changes_tools import (
    replace_with_track_changes,
    insert_before_with_track_changes,
    insert_after_with_track_changes,
)


def _run(coro):
    return asyncio.run(coro)


def _parse(result):
    return json.loads(result)


class TestCrossRunAnchors:
    def test_add_comment_cross_run_anchor(self, cross_run_docx):
        data = _parse(_run(add_comment(cross_run_docx, "Hello World", "x", author="T")))
        assert data["success"] is True, data
        # And the comment round-trips out correctly.
        all_data = _parse(_run(get_all_comments(cross_run_docx)))
        assert all_data["total_comments"] == 1
        assert all_data["comments"][0]["text"] == "x"

    def test_replace_with_track_changes_cross_run_anchor(self, cross_run_docx):
        data = _parse(_run(replace_with_track_changes(
            cross_run_docx, "Hello World", "Goodbye Moon", author="T"
        )))
        assert data["success"] is True, data

    def test_insert_before_with_track_changes_cross_run_anchor(self, cross_run_docx):
        data = _parse(_run(insert_before_with_track_changes(
            cross_run_docx, "Hello World", "Prefix ", author="T"
        )))
        assert data["success"] is True, data

    def test_insert_after_with_track_changes_cross_run_anchor(self, cross_run_docx):
        data = _parse(_run(insert_after_with_track_changes(
            cross_run_docx, "Hello World", " suffix", author="T"
        )))
        assert data["success"] is True, data

    def test_anchor_truly_absent_still_reports_not_found(self, cross_run_docx):
        data = _parse(_run(add_comment(
            cross_run_docx, "Definitely Not Present", "x", author="T"
        )))
        assert data["success"] is False
        assert "not found" in data["error"].lower()
