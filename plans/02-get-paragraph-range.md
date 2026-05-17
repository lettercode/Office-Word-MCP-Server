# Plan 02: `get_paragraph_range` (Issue 2)

**Status: COMPLETED** — Merged to main. See changes.md entry #8.

**Branch:** `feat/get-paragraph-range`
**Issue:** 2

---

## Context

There is no way to read multiple paragraphs in a single call. `get_paragraph_text_from_document` reads exactly one paragraph per call. When mapping out the structure of a section, many sequential calls are needed. This plan adds a batch read operation.

---

## Repository Info

- **Repository:** `<REPO_ROOT>`
- **Run commands with:** `uv run pytest` (or `python -m pytest` if uv unavailable)

## Architecture

- **Utility functions (sync):** `word_document_server/utils/extended_document_utils.py`
- **Tool wrappers (async):** `word_document_server/tools/extended_document_tools.py`
- **MCP registration:** `word_document_server/main.py` inside `register_tools()`
- **Test fixtures:** `tests/conftest.py` — `make_docx(tmp_path)` factory fixture

---

## Files to Modify

1. `word_document_server/utils/extended_document_utils.py` — add `get_paragraph_range()`
2. `word_document_server/tools/extended_document_tools.py` — add `get_paragraph_range_from_document()` async wrapper
3. `word_document_server/main.py` — register tool in `register_tools()`
4. `changes.md` — append entry after all tests pass

## Files to Create

1. `tests/test_get_paragraph_range.py`

---

## TDD Workflow

### Step 1: Create branch

```bash
git checkout -b feat/get-paragraph-range
```

### Step 2: Write failing tests

Create `tests/test_get_paragraph_range.py`:

```python
"""Tests for get_paragraph_range utility."""
import pytest
from docx import Document
from word_document_server.utils.extended_document_utils import get_paragraph_range


class TestGetParagraphRange:
    """Core read behavior."""

    def test_read_full_range(self, make_docx):
        """Read paragraphs 1-3 from 5-paragraph doc, verify 3 results."""
        path = make_docx(paragraphs=["A", "B", "C", "D", "E"])
        result = get_paragraph_range(path, 1, 3)
        assert "error" not in result
        assert len(result["paragraphs"]) == 3
        assert result["paragraphs"][0]["text"] == "B"
        assert result["paragraphs"][1]["text"] == "C"
        assert result["paragraphs"][2]["text"] == "D"

    def test_single_paragraph_range(self, make_docx):
        """start == end returns exactly one paragraph."""
        path = make_docx(paragraphs=["A", "B", "C"])
        result = get_paragraph_range(path, 1, 1)
        assert len(result["paragraphs"]) == 1
        assert result["paragraphs"][0]["text"] == "B"

    def test_includes_correct_fields(self, make_docx):
        """Each result has index, text, style, is_heading fields."""
        path = make_docx(paragraphs=["Hello world"])
        result = get_paragraph_range(path, 0, 0)
        para = result["paragraphs"][0]
        assert "index" in para
        assert "text" in para
        assert "style" in para
        assert "is_heading" in para

    def test_heading_detected(self, make_docx):
        """Heading paragraphs have is_heading=True and correct style."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Title"}]},
            "Content",
            {"style": "Heading 2", "runs": [{"text": "Subtitle"}]},
        ])
        result = get_paragraph_range(path, 0, 2)
        assert result["paragraphs"][0]["is_heading"] is True
        assert result["paragraphs"][0]["style"] == "Heading 1"
        assert result["paragraphs"][1]["is_heading"] is False
        assert result["paragraphs"][2]["is_heading"] is True
        assert result["paragraphs"][2]["style"] == "Heading 2"

    def test_empty_paragraphs_included(self, make_docx):
        """Empty paragraphs are returned with empty text."""
        path = make_docx(paragraphs=["Content", "", "More content"])
        result = get_paragraph_range(path, 0, 2)
        assert result["paragraphs"][1]["text"] == ""

    def test_indices_are_correct(self, make_docx):
        """Returned index values match the actual document indices."""
        path = make_docx(paragraphs=["A", "B", "C", "D", "E"])
        result = get_paragraph_range(path, 2, 4)
        assert result["paragraphs"][0]["index"] == 2
        assert result["paragraphs"][1]["index"] == 3
        assert result["paragraphs"][2]["index"] == 4

    def test_count_field(self, make_docx):
        """Result includes count field matching paragraphs length."""
        path = make_docx(paragraphs=["A", "B", "C", "D", "E"])
        result = get_paragraph_range(path, 1, 3)
        assert result["count"] == 3


class TestGetParagraphRangeValidation:
    """Input validation."""

    def test_start_greater_than_end(self, make_docx):
        """start > end returns error."""
        path = make_docx(paragraphs=["A", "B", "C"])
        result = get_paragraph_range(path, 2, 1)
        assert "error" in result

    def test_end_out_of_bounds(self, make_docx):
        """end beyond doc length returns error."""
        path = make_docx(paragraphs=["A", "B", "C"])
        result = get_paragraph_range(path, 0, 10)
        assert "error" in result

    def test_negative_start(self, make_docx):
        """Negative start returns error."""
        path = make_docx(paragraphs=["A", "B", "C"])
        result = get_paragraph_range(path, -1, 2)
        assert "error" in result

    def test_file_not_found(self, tmp_path):
        """Non-existent file returns error dict."""
        result = get_paragraph_range(str(tmp_path / "missing.docx"), 0, 1)
        assert "error" in result
```

### Step 3: Run tests — verify they FAIL

```bash
uv run pytest tests/test_get_paragraph_range.py -v
```

### Step 4: Implement

#### 4a. Add utility function to `extended_document_utils.py`

Add after the existing `find_text()` function:

```python
def get_paragraph_range(doc_path: str, start_index: int, end_index: int) -> Dict[str, Any]:
    """Get text from a range of paragraphs (start to end index inclusive).

    Args:
        doc_path: Path to the Word document
        start_index: First paragraph index (inclusive, 0-based)
        end_index: Last paragraph index (inclusive, 0-based)

    Returns:
        Dict with "paragraphs" list (each having index, text, style, is_heading)
        and "count" field. Or dict with "error" key on failure.
    """
    import os
    if not os.path.exists(doc_path):
        return {"error": f"Document {doc_path} does not exist"}

    try:
        doc = Document(doc_path)
        total = len(doc.paragraphs)

        if start_index < 0:
            return {"error": f"start_index ({start_index}) must be >= 0"}
        if end_index >= total:
            return {"error": f"end_index ({end_index}) exceeds paragraph count ({total})"}
        if start_index > end_index:
            return {"error": f"start_index ({start_index}) > end_index ({end_index})"}

        paragraphs = []
        for i in range(start_index, end_index + 1):
            para = doc.paragraphs[i]
            paragraphs.append({
                "index": i,
                "text": para.text,
                "style": para.style.name if para.style else "Normal",
                "is_heading": para.style.name.startswith("Heading") if para.style else False
            })

        return {
            "paragraphs": paragraphs,
            "count": len(paragraphs)
        }
    except Exception as e:
        return {"error": f"Failed to get paragraph range: {str(e)}"}
```

#### 4b. Add async wrapper to `extended_document_tools.py`

Add import at top:
```python
from word_document_server.utils.extended_document_utils import get_paragraph_text, find_text, get_paragraph_range
```

Add function:
```python
async def get_paragraph_range_from_document(filename: str, start_index: int, end_index: int) -> str:
    """Get text from a range of paragraphs in a Word document."""
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return json.dumps({"error": f"Document {filename} does not exist"})

    try:
        result = get_paragraph_range(filename, start_index, end_index)
        return json.dumps(result, indent=2)
    except Exception as e:
        return json.dumps({"error": f"Failed to get paragraph range: {str(e)}"})
```

#### 4c. Register tool in `main.py`

Add inside `register_tools()`, near the existing `get_paragraph_text_from_document` registration:

```python
@mcp.tool(
    annotations=ToolAnnotations(
        title="Get Paragraph Range",
        readOnlyHint=True,
    ),
)
def get_paragraph_range(filename: str, start_index: int, end_index: int):
    """Get text from a range of paragraphs (start to end index inclusive).

    Returns a list of paragraph objects, each with index, text, style, and is_heading.
    More efficient than multiple get_paragraph_text_from_document calls.
    """
    return extended_document_tools.get_paragraph_range_from_document(filename, start_index, end_index)
```

### Step 5: Run tests — verify they PASS

```bash
uv run pytest tests/test_get_paragraph_range.py -v
```

### Step 6: Run full test suite

```bash
uv run pytest tests/ -v
```

### Step 7: Update `changes.md`

Append to `changes.md`:

```markdown
### 8. `get_paragraph_range` tool — batch paragraph read
**Branch:** `feat/get-paragraph-range`
**Files:** `extended_document_utils.py`, `extended_document_tools.py`, `main.py`
**Issue resolved:** 2

New MCP tool that reads a contiguous range of paragraphs in a single call. Returns a list
of paragraph objects (index, text, style, is_heading) and a count field. Replaces the pattern
of making many sequential get_paragraph_text_from_document calls to map section structure.
```

### Step 8: Commit and push

```bash
git add tests/test_get_paragraph_range.py word_document_server/utils/extended_document_utils.py word_document_server/tools/extended_document_tools.py word_document_server/main.py changes.md
git commit -m "feat: add get_paragraph_range tool (Issue 2)"
git push -u origin feat/get-paragraph-range
```
