# Plan 01: `delete_paragraph_range` (Issues 1 + 10 + 11)

**Status: COMPLETED** — Merged to main. See changes.md entry #7.

**Branch:** `feat/delete-paragraph-range`
**Issues:** 1 (feature), 10 (workaround — working backward), 11 (workaround — bottom-up deletion)

---

## Context

There is no way to delete multiple consecutive paragraphs in a single call. The only option is `delete_paragraph(filename, paragraph_index)`, which deletes one paragraph at a time and must be called repeatedly from the highest index downward. Issues 10 and 11 document the workaround patterns this creates. This plan eliminates those workarounds.

---

## Repository Info

- **Repository:** `<REPO_ROOT>`
- **Run commands with:** `uv run pytest` (or `python -m pytest` if uv unavailable)

## Architecture

- **Utility functions (sync):** `word_document_server/utils/document_utils.py`
- **Tool wrappers (async):** `word_document_server/tools/content_tools.py`
- **MCP registration:** `word_document_server/main.py` inside `register_tools()`
- **Test fixtures:** `tests/conftest.py` — `make_docx(tmp_path)` factory fixture

---

## Files to Modify

1. `word_document_server/utils/document_utils.py` — add `delete_paragraph_range()` utility
2. `word_document_server/tools/content_tools.py` — add `delete_paragraph_range_tool()` async wrapper
3. `word_document_server/main.py` — register `delete_paragraph_range` tool in `register_tools()`
4. `changes.md` — append entry after all tests pass

## Files to Create

1. `tests/test_delete_paragraph_range.py`

---

## TDD Workflow

### Step 1: Create branch

```bash
git checkout -b feat/delete-paragraph-range
```

### Step 2: Write failing tests

Create `tests/test_delete_paragraph_range.py`:

```python
"""Tests for delete_paragraph_range utility."""
import pytest
from docx import Document
from word_document_server.utils.document_utils import delete_paragraph_range


class TestDeleteRange:
    """Core deletion behavior."""

    def test_delete_middle_range(self, make_docx):
        """Delete paragraphs 1-3 from 5-paragraph doc, verify A and E remain."""
        path = make_docx(paragraphs=["A", "B", "C", "D", "E"])
        result = delete_paragraph_range(path, 1, 3)
        assert "error" not in result.lower()
        doc = Document(path)
        texts = [p.text for p in doc.paragraphs]
        assert texts == ["A", "E"]

    def test_delete_single_paragraph_range(self, make_docx):
        """start_index == end_index deletes exactly one paragraph."""
        path = make_docx(paragraphs=["A", "B", "C"])
        delete_paragraph_range(path, 1, 1)
        doc = Document(path)
        texts = [p.text for p in doc.paragraphs]
        assert texts == ["A", "C"]

    def test_delete_from_start(self, make_docx):
        """Delete paragraphs 0-2 from 5-paragraph doc."""
        path = make_docx(paragraphs=["A", "B", "C", "D", "E"])
        delete_paragraph_range(path, 0, 2)
        doc = Document(path)
        texts = [p.text for p in doc.paragraphs]
        assert texts == ["D", "E"]

    def test_delete_to_end(self, make_docx):
        """Delete last 3 paragraphs of 5-paragraph doc."""
        path = make_docx(paragraphs=["A", "B", "C", "D", "E"])
        delete_paragraph_range(path, 2, 4)
        doc = Document(path)
        texts = [p.text for p in doc.paragraphs]
        assert texts == ["A", "B"]

    def test_delete_all_paragraphs(self, make_docx):
        """Delete all paragraphs from doc."""
        path = make_docx(paragraphs=["A", "B", "C"])
        delete_paragraph_range(path, 0, 2)
        doc = Document(path)
        assert len(doc.paragraphs) == 0 or all(p.text == "" for p in doc.paragraphs)


class TestDeleteRangeValidation:
    """Input validation."""

    def test_start_greater_than_end(self, make_docx):
        """start_index > end_index returns error."""
        path = make_docx(paragraphs=["A", "B", "C"])
        result = delete_paragraph_range(path, 2, 1)
        assert "error" in result.lower()

    def test_end_out_of_bounds(self, make_docx):
        """end_index beyond doc length returns error."""
        path = make_docx(paragraphs=["A", "B", "C"])
        result = delete_paragraph_range(path, 0, 10)
        assert "error" in result.lower()

    def test_negative_start(self, make_docx):
        """Negative start_index returns error."""
        path = make_docx(paragraphs=["A", "B", "C"])
        result = delete_paragraph_range(path, -1, 2)
        assert "error" in result.lower()

    def test_file_not_found(self, tmp_path):
        """Non-existent file returns error."""
        result = delete_paragraph_range(str(tmp_path / "missing.docx"), 0, 1)
        assert "error" in result.lower() or "not exist" in result.lower()


class TestDeleteRangePreservation:
    """Verify surrounding content is untouched."""

    def test_surrounding_paragraphs_unchanged(self, make_docx):
        """Paragraphs before start and after end retain their text and style."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Title"}]},
            "Delete me",
            "Delete me too",
            {"style": "Heading 2", "runs": [{"text": "Subtitle"}]},
        ])
        delete_paragraph_range(path, 1, 2)
        doc = Document(path)
        assert doc.paragraphs[0].text == "Title"
        assert doc.paragraphs[0].style.name == "Heading 1"
        assert doc.paragraphs[1].text == "Subtitle"
        assert doc.paragraphs[1].style.name == "Heading 2"

    def test_paragraph_count_reduced(self, make_docx):
        """Total paragraph count decreases by (end - start + 1)."""
        path = make_docx(paragraphs=["A", "B", "C", "D", "E"])
        delete_paragraph_range(path, 1, 3)
        doc = Document(path)
        assert len(doc.paragraphs) == 2
```

### Step 3: Run tests — verify they FAIL

```bash
uv run pytest tests/test_delete_paragraph_range.py -v
```

All tests should fail with `ImportError` (function doesn't exist yet).

### Step 4: Implement

#### 4a. Add utility function to `document_utils.py`

Add after the existing `replace_paragraph_range` function:

```python
def delete_paragraph_range(doc_path: str, start_index: int, end_index: int) -> str:
    """Delete paragraphs from start_index to end_index inclusive.

    Removes XML elements in reverse order to preserve internal indices.
    When making multiple range deletions, process higher indices first.

    Args:
        doc_path: Path to the Word document
        start_index: First paragraph index to delete (inclusive, 0-based)
        end_index: Last paragraph index to delete (inclusive, 0-based)

    Returns:
        Success or error message string
    """
    import os
    if not os.path.exists(doc_path):
        return f"Error: Document {doc_path} does not exist"

    try:
        doc = Document(doc_path)
        total = len(doc.paragraphs)

        if start_index < 0:
            return f"Error: start_index ({start_index}) must be >= 0"
        if end_index >= total:
            return f"Error: end_index ({end_index}) exceeds paragraph count ({total})"
        if start_index > end_index:
            return f"Error: start_index ({start_index}) > end_index ({end_index})"

        # Delete in reverse order to preserve indices
        for i in range(end_index, start_index - 1, -1):
            p = doc.paragraphs[i]._p
            p.getparent().remove(p)

        doc.save(doc_path)
        count = end_index - start_index + 1
        return f"Successfully deleted {count} paragraph(s) (indices {start_index}-{end_index})."
    except Exception as e:
        return f"Error: Failed to delete paragraph range: {str(e)}"
```

#### 4b. Add async wrapper to `content_tools.py`

```python
async def delete_paragraph_range_tool(filename: str, start_index: int, end_index: int) -> str:
    """Delete a range of paragraphs from a document."""
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}"

    return delete_paragraph_range(filename, start_index, end_index)
```

Add import at top of content_tools.py:
```python
from word_document_server.utils.document_utils import delete_paragraph_range
```

#### 4c. Register tool in `main.py`

Add inside `register_tools()`, near the existing `delete_paragraph` registration:

```python
@mcp.tool(
    annotations=ToolAnnotations(
        title="Delete Paragraph Range",
        destructiveHint=True,
    ),
)
def delete_paragraph_range(filename: str, start_index: int, end_index: int):
    """Delete a range of paragraphs (start to end index inclusive) from a document.

    Tip: When performing multiple range deletions in the same document, process
    higher indices first to prevent index shifting. This tool replaces the pattern
    of calling delete_paragraph in a bottom-up loop.
    """
    return content_tools.delete_paragraph_range_tool(filename, start_index, end_index)
```

### Step 5: Run tests — verify they PASS

```bash
uv run pytest tests/test_delete_paragraph_range.py -v
```

### Step 6: Run full test suite

```bash
uv run pytest tests/ -v
```

### Step 7: Update `changes.md`

Append to `changes.md`:

```markdown
### 7. `delete_paragraph_range` tool — batch paragraph deletion
**Branch:** `feat/delete-paragraph-range`
**Files:** `document_utils.py`, `content_tools.py`, `main.py`
**Issues resolved:** 1, 10, 11

New MCP tool that deletes a contiguous range of paragraphs (by start/end index, inclusive) in a single operation. Removes XML elements in reverse order to preserve indices internally. Tool docstring includes guidance on working backward when making multiple range operations, eliminating the workarounds documented in Issues 10 and 11.
```

### Step 8: Commit and push

```bash
git add tests/test_delete_paragraph_range.py word_document_server/utils/document_utils.py word_document_server/tools/content_tools.py word_document_server/main.py changes.md
git commit -m "feat: add delete_paragraph_range tool (Issues 1, 10, 11)"
git push -u origin feat/delete-paragraph-range
```
