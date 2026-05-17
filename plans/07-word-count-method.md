# Plan 07: Word count method documentation (Issue 7)

**Branch:** `fix/word-count-method`
**Issue:** 7

---

## Context

`get_document_info` reports a word count that may not match Microsoft Word's built-in word count. Word's count has options to include or exclude footnotes, textboxes, and headers/footers. The MCP server uses `paragraph.text.split()` on body paragraphs only, which excludes table text, footnotes, headers/footers, and textboxes. This is not documented, and users relying on the count for journal submission limits may be misled.

---

## Repository Info

- **Repository:** `<REPO_ROOT>`
- **Run commands with:** `uv run pytest` (or `python -m pytest` if uv unavailable)

## Architecture

- **Utility function:** `word_document_server/utils/document_utils.py` → `get_document_properties()`
- **Tool wrapper:** `word_document_server/tools/document_tools.py` → `get_document_info()`
- **MCP registration:** `word_document_server/main.py` inside `register_tools()`

### Current word count implementation
```python
"word_count": sum(len(paragraph.text.split()) for paragraph in doc.paragraphs),
```

---

## Files to Modify

1. `word_document_server/utils/document_utils.py` — add metadata fields and table word count to `get_document_properties()`
2. `word_document_server/main.py` — update tool docstring
3. `changes.md` — append entry after all tests pass

## Files to Create

1. `tests/test_word_count.py`

---

## TDD Workflow

### Step 1: Create branch

```bash
git checkout -b fix/word-count-method
```

### Step 2: Write failing tests

Create `tests/test_word_count.py`:

```python
"""Tests for word count reporting in get_document_properties."""
import pytest
from docx import Document
from word_document_server.utils.document_utils import get_document_properties


class TestWordCountBody:
    """Body paragraph word counting."""

    def test_word_count_matches_split(self, make_docx):
        """Word count matches whitespace-split of body paragraphs."""
        path = make_docx(paragraphs=["Hello world", "One two three"])
        result = get_document_properties(path)
        assert result["word_count"] == 5  # 2 + 3

    def test_empty_paragraphs_zero_words(self, make_docx):
        """Empty paragraphs contribute 0 to word count."""
        path = make_docx(paragraphs=["Hello", "", "World", ""])
        result = get_document_properties(path)
        assert result["word_count"] == 2


class TestWordCountMetadata:
    """Word count metadata fields."""

    def test_word_count_method_present(self, make_docx):
        """Response includes word_count_method field."""
        path = make_docx(paragraphs=["Content"])
        result = get_document_properties(path)
        assert "word_count_method" in result
        assert isinstance(result["word_count_method"], str)

    def test_word_count_note_present(self, make_docx):
        """Response includes word_count_note field."""
        path = make_docx(paragraphs=["Content"])
        result = get_document_properties(path)
        assert "word_count_note" in result
        assert "table" in result["word_count_note"].lower() or "footnote" in result["word_count_note"].lower()


class TestTableWordCount:
    """Separate table word count."""

    def test_table_word_count_present(self, tmp_path):
        """Response includes table_word_count field."""
        path = tmp_path / "table_test.docx"
        doc = Document()
        doc.add_paragraph("Body text here")
        table = doc.add_table(rows=1, cols=1)
        table.cell(0, 0).text = "Table words here now"
        doc.save(str(path))
        result = get_document_properties(str(path))
        assert "table_word_count" in result
        assert result["table_word_count"] == 4  # "Table words here now"

    def test_table_word_count_excludes_body(self, tmp_path):
        """table_word_count only counts table text, not body."""
        path = tmp_path / "test.docx"
        doc = Document()
        doc.add_paragraph("Five words in the body")
        table = doc.add_table(rows=1, cols=1)
        table.cell(0, 0).text = "Two words"
        doc.save(str(path))
        result = get_document_properties(str(path))
        assert result["word_count"] == 5
        assert result["table_word_count"] == 2

    def test_no_tables_zero_count(self, make_docx):
        """Document with no tables has table_word_count=0."""
        path = make_docx(paragraphs=["Just body text"])
        result = get_document_properties(path)
        assert result["table_word_count"] == 0
```

### Step 3: Run tests — verify they FAIL

```bash
uv run pytest tests/test_word_count.py -v
```

Tests should fail because `word_count_method`, `word_count_note`, and `table_word_count` fields don't exist yet.

### Step 4: Implement

#### 4a. Modify `get_document_properties()` in `document_utils.py`

In the return dict, add three new fields after the existing `word_count`:

```python
return {
    "title": core_props.title or "",
    "author": core_props.author or "",
    "subject": core_props.subject or "",
    "keywords": core_props.keywords or "",
    "created": str(core_props.created) if core_props.created else "",
    "modified": str(core_props.modified) if core_props.modified else "",
    "last_modified_by": core_props.last_modified_by or "",
    "revision": core_props.revision or 0,
    "page_count": len(doc.sections),
    "word_count": sum(len(paragraph.text.split()) for paragraph in doc.paragraphs),
    "word_count_method": "body_paragraphs_whitespace_split",
    "word_count_note": (
        "Counts words in body paragraphs only using whitespace splitting. "
        "Does not include text in tables, headers, footers, footnotes, "
        "or textboxes. May differ from Microsoft Word's built-in word count."
    ),
    "table_word_count": sum(
        len(para.text.split())
        for table in doc.tables
        for row in table.rows
        for cell in row.cells
        for para in cell.paragraphs
    ),
    "paragraph_count": len(doc.paragraphs),
    "table_count": len(doc.tables)
}
```

#### 4b. Update tool docstring in `main.py`

```python
@mcp.tool(
    annotations=ToolAnnotations(
        title="Get Document Info",
        readOnlyHint=True,
    ),
)
def get_document_info(filename: str, include_outline: bool = False):
    """Get information about a Word document.

    Word count uses whitespace splitting on body paragraphs only.
    Table text is counted separately in table_word_count.
    This count may differ from Microsoft Word's built-in word count,
    which can optionally include footnotes, textboxes, and headers/footers.
    """
    return document_tools.get_document_info(filename, include_outline)
```

Note: If Plan 05 has not been merged yet, the `include_outline` parameter won't exist. In that case, keep the original signature without `include_outline` and add it to the docstring only.

### Step 5: Run tests — verify they PASS

```bash
uv run pytest tests/test_word_count.py -v
```

### Step 6: Run full test suite

```bash
uv run pytest tests/ -v
```

### Step 7: Update `changes.md`

Append to `changes.md`:

```markdown
### 13. Word count method documentation and table word count
**Branch:** `fix/word-count-method`
**Files:** `document_utils.py`, `main.py`
**Issue resolved:** 7

Added `word_count_method`, `word_count_note`, and `table_word_count` fields to the
`get_document_info` response. The note explains that word_count uses whitespace splitting
on body paragraphs only and may differ from Microsoft Word's built-in count. Table text
is now counted separately in `table_word_count`.
```

### Step 8: Commit and push

```bash
git add tests/test_word_count.py word_document_server/utils/document_utils.py word_document_server/main.py changes.md
git commit -m "fix: document word count method and add table_word_count (Issue 7)"
git push -u origin fix/word-count-method
```
