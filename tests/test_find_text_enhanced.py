"""Tests for find_text with include_paragraph_text enhancement."""
import pytest
from docx import Document
from word_document_server.utils.extended_document_utils import find_text


class TestFindTextDefault:
    """Backward-compatible default behavior."""

    def test_default_returns_context_field(self, make_docx):
        """Default behavior returns context field, not text field."""
        path = make_docx(paragraphs=["Hello world"])
        result = find_text(path, "Hello")
        occ = result["occurrences"][0]
        assert "context" in occ
        assert "text" not in occ

    def test_context_truncated_at_100(self, make_docx):
        """Long paragraph context is truncated to 100 chars + '...'."""
        long_text = "A" * 200
        path = make_docx(paragraphs=[long_text])
        result = find_text(path, "AAAA")
        occ = result["occurrences"][0]
        assert occ["context"].endswith("...")
        assert len(occ["context"]) == 103  # 100 chars + "..."

    def test_short_paragraph_no_truncation(self, make_docx):
        """Short paragraph has no '...' suffix."""
        path = make_docx(paragraphs=["Short text"])
        result = find_text(path, "Short")
        occ = result["occurrences"][0]
        assert not occ["context"].endswith("...")


class TestFindTextIncludeFullParagraph:
    """New include_paragraph_text=True behavior."""

    def test_returns_full_text(self, make_docx):
        """include_paragraph_text=True returns full text, no truncation."""
        long_text = "A" * 200
        path = make_docx(paragraphs=[long_text])
        result = find_text(path, "AAAA", include_paragraph_text=True)
        occ = result["occurrences"][0]
        assert "text" in occ
        assert len(occ["text"]) == 200

    def test_includes_style_field(self, make_docx):
        """Each occurrence includes style when include_paragraph_text=True."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "My Heading"}]},
        ])
        result = find_text(path, "Heading", include_paragraph_text=True)
        occ = result["occurrences"][0]
        assert occ["style"] == "Heading 1"

    def test_no_context_field(self, make_docx):
        """context field is absent when include_paragraph_text=True."""
        path = make_docx(paragraphs=["Hello world"])
        result = find_text(path, "Hello", include_paragraph_text=True)
        occ = result["occurrences"][0]
        assert "context" not in occ
        assert "text" in occ

    def test_false_preserves_backward_compat(self, make_docx):
        """include_paragraph_text=False returns context, no text field."""
        path = make_docx(paragraphs=["Hello world"])
        result = find_text(path, "Hello", include_paragraph_text=False)
        occ = result["occurrences"][0]
        assert "context" in occ
        assert "text" not in occ

    def test_paragraph_index_still_present(self, make_docx):
        """paragraph_index is still returned with include_paragraph_text."""
        path = make_docx(paragraphs=["A", "B", "C target D"])
        result = find_text(path, "target", include_paragraph_text=True)
        occ = result["occurrences"][0]
        assert occ["paragraph_index"] == 2

    def test_multiple_occurrences_each_have_text(self, make_docx):
        """Multiple occurrences each include their full paragraph text."""
        path = make_docx(paragraphs=["First match here", "Second match here"])
        result = find_text(path, "match", include_paragraph_text=True)
        assert result["total_count"] == 2
        assert result["occurrences"][0]["text"] == "First match here"
        assert result["occurrences"][1]["text"] == "Second match here"

    def test_case_insensitive_with_full_text(self, make_docx):
        """Case-insensitive search with full text works."""
        path = make_docx(paragraphs=["Hello World"])
        result = find_text(path, "hello", match_case=False, include_paragraph_text=True)
        assert result["total_count"] == 1
        assert result["occurrences"][0]["text"] == "Hello World"


class TestFindTextHyperlinkAware:
    """Bug B Symptom 3: find_text must locate hyperlink-embedded text.

    The hyperlink_docx fixture seeds 5 occurrences of 'FOO' across plain
    paragraphs, single-run hyperlinks, multi-run hyperlinks, plain trailing
    paragraphs, and a table cell. Before the fix, find_text only saw 3.
    """

    def test_finds_all_occurrences_including_hyperlinks(self, hyperlink_docx):
        result = find_text(hyperlink_docx, "FOO")
        assert result["total_count"] == 5

    def test_full_text_view_includes_hyperlink_display(self, hyperlink_docx):
        result = find_text(hyperlink_docx, "FOO", include_paragraph_text=True)
        full_texts = [occ.get("text") for occ in result["occurrences"]]
        # The hyperlink-bearing paragraph's full text must include "FOO bar".
        assert any(t and "FOO bar" in t for t in full_texts)
        # The cross-run hyperlink paragraph's full text must include "FOO baz".
        assert any(t and "FOO baz" in t for t in full_texts)

    def test_can_locate_hyperlink_only_substring(self, hyperlink_docx):
        # 'bar' only appears inside a hyperlink. Pre-fix this returned 0.
        result = find_text(hyperlink_docx, "bar")
        assert result["total_count"] == 1
