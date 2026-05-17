# Plan 04: `find_text_in_document` — include full paragraph text (Issue 4)

**Status: COMPLETED** — Merged to main. See changes.md entry #10.

**Branch:** `feat/find-text-include-full-paragraph`
**Issue:** 4

---

## Context

`find_text_in_document` returns a truncated `context` field (approximately 100 characters). After finding text, you almost always need a follow-up `get_paragraph_text_from_document` call to read the full paragraph. This doubles the number of calls for every search. This plan adds an `include_paragraph_text` parameter that returns full text instead of truncated context.

---

## Repository Info

- **Repository:** `<REPO_ROOT>`
- **Run commands with:** `uv run pytest` (or `python -m pytest` if uv unavailable)

## Architecture

- **Utility function:** `word_document_server/utils/extended_document_utils.py` → `find_text()`
- **Tool wrapper:** `word_document_server/tools/extended_document_tools.py` → `find_text_in_document()`
- **MCP registration:** `word_document_server/main.py` inside `register_tools()`

### Current `find_text()` behavior
Each occurrence returns:
```python
{"paragraph_index": i, "position": pos, "context": para.text[:100] + "..."}
```

---

## Files to Modify

1. `word_document_server/utils/extended_document_utils.py` — add `include_paragraph_text` parameter to `find_text()`
2. `word_document_server/tools/extended_document_tools.py` — pass parameter through `find_text_in_document()`
3. `word_document_server/main.py` — update tool signature and docstring
4. `changes.md` — append entry after all tests pass

## Files to Create

1. `tests/test_find_text_enhanced.py`

---

## TDD Workflow

### Step 1: Create branch

```bash
git checkout -b feat/find-text-include-full-paragraph
```

### Step 2: Write failing tests

Create `tests/test_find_text_enhanced.py`:

```python
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
```

### Step 3: Run tests — verify they FAIL

```bash
uv run pytest tests/test_find_text_enhanced.py -v
```

Tests should fail because `find_text()` doesn't accept `include_paragraph_text` parameter yet.

### Step 4: Implement

#### 4a. Modify `find_text()` in `extended_document_utils.py`

Update function signature:
```python
def find_text(doc_path: str, text_to_find: str, match_case: bool = True,
              whole_word: bool = False, include_paragraph_text: bool = False) -> Dict[str, Any]:
```

In the paragraph search loop, change the occurrence building logic:

```python
# Replace the current occurrence dict construction:
# OLD:
#   results["occurrences"].append({
#       "paragraph_index": i,
#       "position": pos,
#       "context": para.text[:100] + ("..." if len(para.text) > 100 else "")
#   })

# NEW:
occurrence = {
    "paragraph_index": i,
    "position": pos,
}
if include_paragraph_text:
    occurrence["text"] = para.text
    occurrence["style"] = para.style.name if para.style else "Normal"
else:
    occurrence["context"] = para.text[:100] + ("..." if len(para.text) > 100 else "")
results["occurrences"].append(occurrence)
```

Apply the same change to the table search loop (for occurrences found in tables).

#### 4b. Update async wrapper in `extended_document_tools.py`

```python
async def find_text_in_document(filename: str, text_to_find: str, match_case: bool = True,
                                whole_word: bool = False, include_paragraph_text: bool = False) -> str:
    """Find occurrences of specific text in a Word document."""
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    if not text_to_find:
        return "Search text cannot be empty"

    try:
        result = find_text(filename, text_to_find, match_case, whole_word, include_paragraph_text)
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Failed to search for text: {str(e)}"
```

#### 4c. Update tool registration in `main.py`

```python
@mcp.tool(
    annotations=ToolAnnotations(
        title="Find Text",
        readOnlyHint=True,
    ),
)
def find_text_in_document(filename: str, text_to_find: str, match_case: bool = True,
                          whole_word: bool = False, include_paragraph_text: bool = False):
    """Find occurrences of specific text in a Word document.

    When include_paragraph_text is True, each occurrence returns the full paragraph text
    and its style instead of a truncated 100-character context preview.
    """
    return extended_document_tools.find_text_in_document(
        filename, text_to_find, match_case, whole_word, include_paragraph_text
    )
```

### Step 5: Run tests — verify they PASS

```bash
uv run pytest tests/test_find_text_enhanced.py -v
```

### Step 6: Run full test suite

```bash
uv run pytest tests/ -v
```

### Step 7: Update `changes.md`

Append to `changes.md`:

```markdown
### 10. `find_text_in_document` — include full paragraph text option
**Branch:** `feat/find-text-include-full-paragraph`
**Files:** `extended_document_utils.py`, `extended_document_tools.py`, `main.py`
**Issue resolved:** 4

Added `include_paragraph_text` boolean parameter (default False) to `find_text_in_document`.
When True, each occurrence returns the full paragraph text and style name instead of a
truncated 100-character context. Backward compatible — default behavior unchanged.
```

### Step 8: Commit and push

```bash
git add tests/test_find_text_enhanced.py word_document_server/utils/extended_document_utils.py word_document_server/tools/extended_document_tools.py word_document_server/main.py changes.md
git commit -m "feat: add include_paragraph_text option to find_text (Issue 4)"
git push -u origin feat/find-text-include-full-paragraph
```
