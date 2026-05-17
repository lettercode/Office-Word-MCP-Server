# Plan 05: `get_document_info` — include outline (Issue 5)

**Status: COMPLETED** — Merged to main. See changes.md entry #11.

**Branch:** `feat/get-document-info-include-outline`
**Issue:** 5

---

## Context

`get_document_info` returns word count, paragraph count, and table count, but not the document's heading structure. When starting to edit a document, you need the outline first. Currently this requires a separate `get_document_outline` call. This plan adds an `include_outline` parameter to combine both in one call.

---

## Repository Info

- **Repository:** `<REPO_ROOT>`
- **Run commands with:** `uv run pytest` (or `python -m pytest` if uv unavailable)

## Architecture

- **Utility function:** `word_document_server/utils/document_utils.py` → `get_document_properties()`
- **Tool wrapper:** `word_document_server/tools/document_tools.py` → `get_document_info()`
- **MCP registration:** `word_document_server/main.py` inside `register_tools()`

### Current `get_document_properties()` return
```python
{
    "title": "", "author": "", "subject": "", "keywords": "",
    "created": "", "modified": "", "last_modified_by": "",
    "revision": 0, "page_count": N, "word_count": N,
    "paragraph_count": N, "table_count": N
}
```

---

## Files to Modify

1. `word_document_server/utils/document_utils.py` — add `include_outline` parameter to `get_document_properties()`
2. `word_document_server/tools/document_tools.py` — pass parameter through `get_document_info()`
3. `word_document_server/main.py` — update tool signature and docstring
4. `changes.md` — append entry after all tests pass

## Files to Create

1. `tests/test_get_document_info_outline.py`

---

## TDD Workflow

### Step 1: Create branch

```bash
git checkout -b feat/get-document-info-include-outline
```

### Step 2: Write failing tests

Create `tests/test_get_document_info_outline.py`:

```python
"""Tests for get_document_info with include_outline enhancement."""
import pytest
from docx import Document
from word_document_server.utils.document_utils import get_document_properties


class TestGetDocumentInfoDefault:
    """Backward-compatible default behavior."""

    def test_default_no_headings_key(self, make_docx):
        """Default response does not include headings key."""
        path = make_docx(paragraphs=["Content"])
        result = get_document_properties(path)
        assert "headings" not in result

    def test_default_still_has_word_count(self, make_docx):
        """Default response still includes standard fields."""
        path = make_docx(paragraphs=["Hello world"])
        result = get_document_properties(path)
        assert "word_count" in result
        assert "paragraph_count" in result


class TestGetDocumentInfoWithOutline:
    """include_outline=True behavior."""

    def test_include_outline_adds_headings(self, make_docx):
        """include_outline=True adds headings array to response."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Chapter 1"}]},
            "Content",
            {"style": "Heading 2", "runs": [{"text": "Section 1.1"}]},
        ])
        result = get_document_properties(path, include_outline=True)
        assert "headings" in result
        assert len(result["headings"]) == 2

    def test_headings_have_correct_fields(self, make_docx):
        """Each heading has index, text, style, level."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Title"}]},
        ])
        result = get_document_properties(path, include_outline=True)
        h = result["headings"][0]
        assert "index" in h
        assert "text" in h
        assert "style" in h
        assert "level" in h

    def test_heading_levels_correct(self, make_docx):
        """Heading 1 has level=1, Heading 2 has level=2."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "H1"}]},
            {"style": "Heading 2", "runs": [{"text": "H2"}]},
            {"style": "Heading 3", "runs": [{"text": "H3"}]},
        ])
        result = get_document_properties(path, include_outline=True)
        assert result["headings"][0]["level"] == 1
        assert result["headings"][1]["level"] == 2
        assert result["headings"][2]["level"] == 3

    def test_heading_index_matches_position(self, make_docx):
        """Heading index corresponds to paragraph position in document."""
        path = make_docx(paragraphs=[
            "Content before",
            {"style": "Heading 1", "runs": [{"text": "Title"}]},
            "Content after",
        ])
        result = get_document_properties(path, include_outline=True)
        assert result["headings"][0]["index"] == 1
        assert result["headings"][0]["text"] == "Title"

    def test_no_headings_returns_empty_array(self, make_docx):
        """Doc with no headings returns empty headings array."""
        path = make_docx(paragraphs=["Just normal text", "More text"])
        result = get_document_properties(path, include_outline=True)
        assert result["headings"] == []

    def test_standard_fields_still_present(self, make_docx):
        """include_outline=True still returns all standard fields."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Title"}]},
            "Content",
        ])
        result = get_document_properties(path, include_outline=True)
        assert "word_count" in result
        assert "paragraph_count" in result
        assert "table_count" in result

    def test_include_outline_false_same_as_default(self, make_docx):
        """include_outline=False explicitly does not include headings."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Title"}]},
        ])
        result = get_document_properties(path, include_outline=False)
        assert "headings" not in result
```

### Step 3: Run tests — verify they FAIL

```bash
uv run pytest tests/test_get_document_info_outline.py -v
```

Tests should fail because `get_document_properties()` doesn't accept `include_outline`.

### Step 4: Implement

#### 4a. Modify `get_document_properties()` in `document_utils.py`

Update signature:
```python
def get_document_properties(doc_path: str, include_outline: bool = False) -> Dict[str, Any]:
```

After building the existing result dict (before the return statement), add:

```python
if include_outline:
    headings = []
    for i, para in enumerate(doc.paragraphs):
        if para.style and para.style.name.startswith("Heading"):
            try:
                level = int(para.style.name.split(" ")[1])
            except (ValueError, IndexError):
                level = 0
            headings.append({
                "index": i,
                "text": para.text,
                "style": para.style.name,
                "level": level
            })
    result["headings"] = headings
```

Note: You'll need to assign the existing return value to a variable `result` first, then add headings conditionally, then return `result`.

#### 4b. Update async wrapper in `document_tools.py`

```python
async def get_document_info(filename: str, include_outline: bool = False) -> str:
    """Get information about a Word document."""
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    try:
        properties = get_document_properties(filename, include_outline=include_outline)
        return json.dumps(properties, indent=2)
    except Exception as e:
        return f"Failed to get document info: {str(e)}"
```

Add import if not already present:
```python
from word_document_server.utils.file_utils import ensure_docx_extension
```

#### 4c. Update tool registration in `main.py`

```python
@mcp.tool(
    annotations=ToolAnnotations(
        title="Get Document Info",
        readOnlyHint=True,
    ),
)
def get_document_info(filename: str, include_outline: bool = False):
    """Get information about a Word document.

    When include_outline is True, also returns a headings array with text,
    style, level, and paragraph index for each heading in the document.
    """
    return document_tools.get_document_info(filename, include_outline)
```

### Step 5: Run tests — verify they PASS

```bash
uv run pytest tests/test_get_document_info_outline.py -v
```

### Step 6: Run full test suite

```bash
uv run pytest tests/ -v
```

### Step 7: Update `changes.md`

Append to `changes.md`:

```markdown
### 11. `get_document_info` — include outline option
**Branch:** `feat/get-document-info-include-outline`
**Files:** `document_utils.py`, `document_tools.py`, `main.py`
**Issue resolved:** 5

Added `include_outline` boolean parameter (default False) to `get_document_info`.
When True, adds a `headings` array to the response containing each heading's text,
style, level, and paragraph index. Backward compatible — default behavior unchanged.
```

### Step 8: Commit and push

```bash
git add tests/test_get_document_info_outline.py word_document_server/utils/document_utils.py word_document_server/tools/document_tools.py word_document_server/main.py changes.md
git commit -m "feat: add include_outline option to get_document_info (Issue 5)"
git push -u origin feat/get-document-info-include-outline
```
