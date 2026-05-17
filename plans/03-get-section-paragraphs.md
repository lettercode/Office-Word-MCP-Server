# Plan 03: `get_section_paragraphs` (Issues 3 + 12)

**Status: COMPLETED** — Merged to main. See changes.md entry #9.

**Branch:** `feat/get-section-paragraphs`
**Issues:** 3 (feature), 12 (workaround — finding section boundaries)

---

## Context

There is no way to get all paragraphs under a specific heading (up to the next same-or-higher-level heading) in a single call. This is the most common operation when editing a paper. Currently it requires finding the heading, walking forward one paragraph at a time, and manually checking styles. Issue 12 documents this workaround. This plan eliminates it.

---

## Repository Info

- **Repository:** `<REPO_ROOT>`
- **Run commands with:** `uv run pytest` (or `python -m pytest` if uv unavailable)

## Architecture

- **Utility functions (sync):** `word_document_server/utils/extended_document_utils.py`
- **Tool wrappers (async):** `word_document_server/tools/extended_document_tools.py`
- **MCP registration:** `word_document_server/main.py` inside `register_tools()`
- **Test fixtures:** `tests/conftest.py` — `make_docx(tmp_path)` factory fixture

### Key existing function

`_normalize_text(s)` in `document_utils.py` — NFKC normalization + whitespace collapse + strip. Import this for heading text matching.

---

## Files to Modify

1. `word_document_server/utils/extended_document_utils.py` — add `get_section_paragraphs()`
2. `word_document_server/tools/extended_document_tools.py` — add async wrapper
3. `word_document_server/main.py` — register tool in `register_tools()`
4. `changes.md` — append entry after all tests pass

## Files to Create

1. `tests/test_get_section_paragraphs.py`

---

## TDD Workflow

### Step 1: Create branch

```bash
git checkout -b feat/get-section-paragraphs
```

### Step 2: Write failing tests

Create `tests/test_get_section_paragraphs.py`:

```python
"""Tests for get_section_paragraphs utility."""
import pytest
from docx import Document
from word_document_server.utils.extended_document_utils import get_section_paragraphs


class TestGetSectionParagraphs:
    """Core section extraction behavior."""

    def test_basic_h1_section(self, make_docx):
        """H1 section returns content until next H1."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Section A"}]},
            "Content 1",
            "Content 2",
            {"style": "Heading 1", "runs": [{"text": "Section B"}]},
            "Content 3",
        ])
        result = get_section_paragraphs(path, "Section A")
        assert "error" not in result
        assert result["heading_index"] == 0
        assert result["heading_text"] == "Section A"
        assert result["heading_style"] == "Heading 1"
        assert result["heading_level"] == 1
        assert result["next_heading_index"] == 3
        # Content paragraphs (excluding heading itself)
        content = [p for p in result["paragraphs"] if p["index"] != 0]
        content_texts = [p["text"] for p in content]
        assert "Content 1" in content_texts
        assert "Content 2" in content_texts
        assert "Content 3" not in content_texts

    def test_h2_section_stops_at_same_level(self, make_docx):
        """H2 section stops at next H2."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Chapter"}]},
            {"style": "Heading 2", "runs": [{"text": "Part A"}]},
            "Content A",
            {"style": "Heading 2", "runs": [{"text": "Part B"}]},
            "Content B",
        ])
        result = get_section_paragraphs(path, "Part A")
        assert result["heading_level"] == 2
        assert result["next_heading_index"] == 3
        content = [p for p in result["paragraphs"] if p["index"] > result["heading_index"]]
        assert len(content) == 1
        assert content[0]["text"] == "Content A"

    def test_h2_section_stops_at_higher_level(self, make_docx):
        """H2 section stops at H1 (higher level = lower number)."""
        path = make_docx(paragraphs=[
            {"style": "Heading 2", "runs": [{"text": "Subsection"}]},
            "Content",
            {"style": "Heading 1", "runs": [{"text": "Next Chapter"}]},
        ])
        result = get_section_paragraphs(path, "Subsection")
        assert result["next_heading_index"] == 2

    def test_last_section_no_next_heading(self, make_docx):
        """Last section returns content to end, next_heading_index is null."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Only Section"}]},
            "Content 1",
            "Content 2",
        ])
        result = get_section_paragraphs(path, "Only Section")
        assert result["next_heading_index"] is None
        content = [p for p in result["paragraphs"] if p["index"] > 0]
        assert len(content) == 2

    def test_empty_section(self, make_docx):
        """Section with no content between headings returns only heading in paragraphs."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Empty"}]},
            {"style": "Heading 1", "runs": [{"text": "Next"}]},
        ])
        result = get_section_paragraphs(path, "Empty")
        assert result["next_heading_index"] == 1
        content = [p for p in result["paragraphs"] if p["index"] > 0]
        assert len(content) == 0

    def test_include_heading_true(self, make_docx):
        """Default includes heading paragraph in results."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Title"}]},
            "Content",
        ])
        result = get_section_paragraphs(path, "Title")
        assert any(p["text"] == "Title" for p in result["paragraphs"])

    def test_include_heading_false(self, make_docx):
        """include_heading=False omits heading from paragraphs list."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Title"}]},
            "Content",
        ])
        result = get_section_paragraphs(path, "Title", include_heading=False)
        assert not any(p["text"] == "Title" for p in result["paragraphs"])

    def test_with_empty_spacer_paragraphs(self, make_docx):
        """Empty spacer paragraphs between content are included."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Section"}]},
            "",
            "Content 1",
            "",
            "Content 2",
            "",
            {"style": "Heading 1", "runs": [{"text": "Next"}]},
        ])
        result = get_section_paragraphs(path, "Section", include_heading=False)
        assert len(result["paragraphs"]) == 5  # 3 spacers + 2 content


class TestSectionHeadingMatching:
    """Heading text matching behavior."""

    def test_heading_not_found(self, make_docx):
        """Non-existent heading returns error."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Actual Title"}]},
        ])
        result = get_section_paragraphs(path, "Nonexistent")
        assert "error" in result

    def test_partial_text_match(self, make_docx):
        """Substring of heading text finds the heading."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Chapter 1: Introduction"}]},
            "Content",
        ])
        result = get_section_paragraphs(path, "Introduction")
        assert result["heading_index"] == 0

    def test_normalized_whitespace_match(self, make_docx):
        """Extra whitespace in search text still matches."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "My  Section"}]},
            "Content",
        ])
        result = get_section_paragraphs(path, "My Section")
        assert result["heading_index"] == 0

    def test_file_not_found(self, tmp_path):
        """Non-existent file returns error dict."""
        result = get_section_paragraphs(str(tmp_path / "missing.docx"), "Heading")
        assert "error" in result
```

### Step 3: Run tests — verify they FAIL

```bash
uv run pytest tests/test_get_section_paragraphs.py -v
```

### Step 4: Implement

#### 4a. Add utility function to `extended_document_utils.py`

Add import at top:
```python
import unicodedata
import re
```

Add helper (or import from document_utils if accessible):
```python
def _normalize_text(s: str) -> str:
    """Normalize text for reliable matching: NFKC normalize, collapse whitespace, strip."""
    s = unicodedata.normalize("NFKC", s)
    s = re.sub(r'\s+', ' ', s)
    return s.strip()
```

Note: If `_normalize_text` is already importable from `document_utils`, import it instead of duplicating:
```python
from word_document_server.utils.document_utils import _normalize_text
```

Add function:
```python
def get_section_paragraphs(doc_path: str, heading_text: str, include_heading: bool = True) -> Dict[str, Any]:
    """Get all paragraphs under a heading until the next same-or-higher-level heading.

    Uses normalized text matching (NFKC + whitespace collapse) to find the heading.
    Falls back to substring matching if exact match fails.

    Args:
        doc_path: Path to the Word document
        heading_text: Text of the heading to find
        include_heading: Whether to include the heading paragraph itself (default True)

    Returns:
        Dict with heading metadata and paragraphs list.
        Or dict with "error" key on failure.
    """
    import os
    if not os.path.exists(doc_path):
        return {"error": f"Document {doc_path} does not exist"}

    try:
        doc = Document(doc_path)
        normalized_search = _normalize_text(heading_text)

        # Find the heading paragraph
        heading_idx = None
        for i, para in enumerate(doc.paragraphs):
            if para.style and para.style.name.startswith("Heading"):
                normalized_para = _normalize_text(para.text)
                if normalized_para == normalized_search:
                    heading_idx = i
                    break

        # Fallback: substring match on heading paragraphs
        if heading_idx is None:
            for i, para in enumerate(doc.paragraphs):
                if para.style and para.style.name.startswith("Heading"):
                    if normalized_search in _normalize_text(para.text):
                        heading_idx = i
                        break

        if heading_idx is None:
            return {"error": f"Heading '{heading_text}' not found in document"}

        heading_para = doc.paragraphs[heading_idx]
        heading_style = heading_para.style.name if heading_para.style else "Heading 1"

        # Extract heading level number
        try:
            heading_level = int(heading_style.split(" ")[1])
        except (ValueError, IndexError):
            heading_level = 1

        # Walk forward collecting paragraphs until next same-or-higher-level heading
        next_heading_idx = None
        content_end_idx = heading_idx  # will update as we find content

        for i in range(heading_idx + 1, len(doc.paragraphs)):
            para = doc.paragraphs[i]
            if para.style and para.style.name.startswith("Heading"):
                try:
                    para_level = int(para.style.name.split(" ")[1])
                except (ValueError, IndexError):
                    para_level = 1
                if para_level <= heading_level:
                    next_heading_idx = i
                    break
            content_end_idx = i

        # If no next heading found, content goes to end of doc
        if next_heading_idx is None:
            content_end_idx = len(doc.paragraphs) - 1
        else:
            content_end_idx = next_heading_idx - 1

        # Build paragraphs list
        paragraphs = []
        start = heading_idx if include_heading else heading_idx + 1
        end = content_end_idx

        for i in range(start, end + 1):
            para = doc.paragraphs[i]
            paragraphs.append({
                "index": i,
                "text": para.text,
                "style": para.style.name if para.style else "Normal",
                "is_heading": para.style.name.startswith("Heading") if para.style else False
            })

        return {
            "heading_index": heading_idx,
            "heading_text": heading_para.text,
            "heading_style": heading_style,
            "heading_level": heading_level,
            "content_start_index": heading_idx + 1 if heading_idx + 1 <= content_end_idx else None,
            "content_end_index": content_end_idx if content_end_idx > heading_idx else None,
            "next_heading_index": next_heading_idx,
            "paragraphs": paragraphs
        }
    except Exception as e:
        return {"error": f"Failed to get section paragraphs: {str(e)}"}
```

#### 4b. Add async wrapper to `extended_document_tools.py`

Add import:
```python
from word_document_server.utils.extended_document_utils import get_paragraph_text, find_text, get_section_paragraphs
```

Add function:
```python
async def get_section_paragraphs_from_document(filename: str, heading_text: str, include_heading: bool = True) -> str:
    """Get all paragraphs under a heading until the next same-or-higher-level heading."""
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return json.dumps({"error": f"Document {filename} does not exist"})

    try:
        result = get_section_paragraphs(filename, heading_text, include_heading)
        return json.dumps(result, indent=2)
    except Exception as e:
        return json.dumps({"error": f"Failed to get section paragraphs: {str(e)}"})
```

#### 4c. Register tool in `main.py`

```python
@mcp.tool(
    annotations=ToolAnnotations(
        title="Get Section Paragraphs",
        readOnlyHint=True,
    ),
)
def get_section_paragraphs(filename: str, heading_text: str, include_heading: bool = True):
    """Get all paragraphs under a heading until the next same-or-higher-level heading.

    Uses normalized text matching to find the heading. Returns heading metadata,
    section boundaries, and all paragraphs in the section.

    Eliminates the need for multiple get_paragraph_text calls to find section boundaries.
    """
    return extended_document_tools.get_section_paragraphs_from_document(filename, heading_text, include_heading)
```

### Step 5: Run tests — verify they PASS

```bash
uv run pytest tests/test_get_section_paragraphs.py -v
```

### Step 6: Run full test suite

```bash
uv run pytest tests/ -v
```

### Step 7: Update `changes.md`

Append to `changes.md`:

```markdown
### 9. `get_section_paragraphs` tool — section content extraction
**Branch:** `feat/get-section-paragraphs`
**Files:** `extended_document_utils.py`, `extended_document_tools.py`, `main.py`
**Issues resolved:** 3, 12

New MCP tool that returns all paragraphs under a heading (up to the next same-or-higher-level
heading) in a single call. Uses normalized text matching with substring fallback. Returns
heading metadata (index, text, style, level), section boundaries, and paragraph list.
Eliminates the multi-call workaround documented in Issue 12.
```

### Step 8: Commit and push

```bash
git add tests/test_get_section_paragraphs.py word_document_server/utils/extended_document_utils.py word_document_server/tools/extended_document_tools.py word_document_server/main.py changes.md
git commit -m "feat: add get_section_paragraphs tool (Issues 3, 12)"
git push -u origin feat/get-section-paragraphs
```
