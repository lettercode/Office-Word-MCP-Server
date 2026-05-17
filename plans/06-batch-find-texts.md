# Plan 06: `find_texts_in_document` — batch text search (Issue 6)

**Status: COMPLETED** — Merged to main. See changes.md entry #12.

**Branch:** `feat/batch-find-texts`
**Issue:** 6

---

## Context

When preparing for multiple edits, you need to locate many different text strings. Each requires a separate `find_text_in_document` call. A batch function loads the document once and searches for all strings in a single pass, reducing 12+ calls to 1.

---

## Repository Info

- **Repository:** `<REPO_ROOT>`
- **Run commands with:** `uv run pytest` (or `python -m pytest` if uv unavailable)

## Architecture

- **Utility functions (sync):** `word_document_server/utils/extended_document_utils.py`
- **Tool wrappers (async):** `word_document_server/tools/extended_document_tools.py`
- **MCP registration:** `word_document_server/main.py` inside `register_tools()`

### Existing `find_text()` return shape (per search string)
```python
{
    "query": "...",
    "match_case": True,
    "whole_word": False,
    "occurrences": [{"paragraph_index": N, "position": N, "context": "..."}],
    "total_count": N
}
```

---

## Files to Modify

1. `word_document_server/utils/extended_document_utils.py` — add `find_texts()` utility
2. `word_document_server/tools/extended_document_tools.py` — add async wrapper
3. `word_document_server/main.py` — register new tool
4. `changes.md` — append entry after all tests pass

## Files to Create

1. `tests/test_batch_find_texts.py`

---

## TDD Workflow

### Step 1: Create branch

```bash
git checkout -b feat/batch-find-texts
```

### Step 2: Write failing tests

Create `tests/test_batch_find_texts.py`:

```python
"""Tests for batch find_texts utility."""
import pytest
from docx import Document
from word_document_server.utils.extended_document_utils import find_texts


class TestBatchFindTexts:
    """Core batch search behavior."""

    def test_find_multiple_strings(self, make_docx):
        """Find 3 different strings, each returns correct paragraph_index."""
        path = make_docx(paragraphs=[
            "Alpha content here",
            "Beta content here",
            "Gamma content here",
        ])
        result = find_texts(path, ["Alpha", "Beta", "Gamma"])
        assert "Alpha" in result
        assert "Beta" in result
        assert "Gamma" in result
        assert result["Alpha"]["occurrences"][0]["paragraph_index"] == 0
        assert result["Beta"]["occurrences"][0]["paragraph_index"] == 1
        assert result["Gamma"]["occurrences"][0]["paragraph_index"] == 2

    def test_missing_string_returns_empty(self, make_docx):
        """String not in doc returns total_count=0, empty occurrences."""
        path = make_docx(paragraphs=["Hello world"])
        result = find_texts(path, ["Nonexistent"])
        assert result["Nonexistent"]["total_count"] == 0
        assert result["Nonexistent"]["occurrences"] == []

    def test_single_string_works(self, make_docx):
        """Single-item list works correctly."""
        path = make_docx(paragraphs=["Hello world"])
        result = find_texts(path, ["Hello"])
        assert result["Hello"]["total_count"] == 1

    def test_case_insensitive(self, make_docx):
        """match_case=False finds case-variant matches."""
        path = make_docx(paragraphs=["Hello World"])
        result = find_texts(path, ["hello", "WORLD"], match_case=False)
        assert result["hello"]["total_count"] == 1
        assert result["WORLD"]["total_count"] == 1

    def test_include_paragraph_text_propagates(self, make_docx):
        """include_paragraph_text=True returns full text for all results."""
        path = make_docx(paragraphs=["Alpha content", "Beta content"])
        result = find_texts(path, ["Alpha", "Beta"], include_paragraph_text=True)
        assert "text" in result["Alpha"]["occurrences"][0]
        assert result["Alpha"]["occurrences"][0]["text"] == "Alpha content"

    def test_empty_list_returns_empty_dict(self, make_docx):
        """Empty texts_to_find returns empty results dict."""
        path = make_docx(paragraphs=["Content"])
        result = find_texts(path, [])
        assert result == {} or result.get("results") == {}

    def test_duplicate_search_terms(self, make_docx):
        """Duplicate search terms don't cause errors."""
        path = make_docx(paragraphs=["Hello world"])
        result = find_texts(path, ["Hello", "Hello"])
        assert "Hello" in result
        assert result["Hello"]["total_count"] == 1

    def test_each_result_has_standard_fields(self, make_docx):
        """Each result has query, match_case, occurrences, total_count."""
        path = make_docx(paragraphs=["Target text"])
        result = find_texts(path, ["Target"])
        entry = result["Target"]
        assert "occurrences" in entry
        assert "total_count" in entry

    def test_file_not_found(self, tmp_path):
        """Non-existent file returns error."""
        result = find_texts(str(tmp_path / "missing.docx"), ["anything"])
        assert "error" in result

    def test_multiple_occurrences_same_string(self, make_docx):
        """String found in multiple paragraphs returns all occurrences."""
        path = make_docx(paragraphs=["Match here", "No match", "Match again"])
        result = find_texts(path, ["Match"])
        assert result["Match"]["total_count"] == 2
        indices = [o["paragraph_index"] for o in result["Match"]["occurrences"]]
        assert 0 in indices
        assert 2 in indices
```

### Step 3: Run tests — verify they FAIL

```bash
uv run pytest tests/test_batch_find_texts.py -v
```

### Step 4: Implement

#### 4a. Add utility function to `extended_document_utils.py`

```python
def find_texts(doc_path: str, texts_to_find: List[str], match_case: bool = True,
               include_paragraph_text: bool = False) -> Dict[str, Any]:
    """Find multiple text strings in a document in a single pass.

    Loads the document once and checks each paragraph against all search strings.

    Args:
        doc_path: Path to the Word document
        texts_to_find: List of strings to search for
        match_case: Case-sensitive matching (default True)
        include_paragraph_text: Include full paragraph text in results (default False)

    Returns:
        Dict keyed by search string. Each value has "occurrences" list and "total_count",
        matching the shape of find_text() results.
        Or dict with "error" key on failure.
    """
    import os
    if not os.path.exists(doc_path):
        return {"error": f"Document {doc_path} does not exist"}

    if not texts_to_find:
        return {}

    try:
        doc = Document(doc_path)

        # Initialize results for each unique search string
        unique_texts = list(dict.fromkeys(texts_to_find))  # preserve order, dedupe
        results = {}
        for text in unique_texts:
            results[text] = {
                "occurrences": [],
                "total_count": 0
            }

        # Single pass through all paragraphs
        for i, para in enumerate(doc.paragraphs):
            para_text = para.text

            for search_text in unique_texts:
                compare_para = para_text if match_case else para_text.lower()
                compare_search = search_text if match_case else search_text.lower()

                start_pos = 0
                while True:
                    pos = compare_para.find(compare_search, start_pos)
                    if pos == -1:
                        break

                    occurrence = {
                        "paragraph_index": i,
                        "position": pos,
                    }
                    if include_paragraph_text:
                        occurrence["text"] = para.text
                        occurrence["style"] = para.style.name if para.style else "Normal"
                    else:
                        occurrence["context"] = para.text[:100] + ("..." if len(para.text) > 100 else "")

                    results[search_text]["occurrences"].append(occurrence)
                    results[search_text]["total_count"] += 1
                    start_pos = pos + len(compare_search)

        return results
    except Exception as e:
        return {"error": f"Failed to search for texts: {str(e)}"}
```

#### 4b. Add async wrapper to `extended_document_tools.py`

Add import:
```python
from word_document_server.utils.extended_document_utils import get_paragraph_text, find_text, find_texts
```

Add function:
```python
async def find_texts_in_document_tool(filename: str, texts_to_find: list,
                                       match_case: bool = True,
                                       include_paragraph_text: bool = False) -> str:
    """Find occurrences of multiple text strings in a document."""
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return json.dumps({"error": f"Document {filename} does not exist"})

    try:
        result = find_texts(filename, texts_to_find, match_case, include_paragraph_text)
        return json.dumps(result, indent=2)
    except Exception as e:
        return json.dumps({"error": f"Failed to search for texts: {str(e)}"})
```

#### 4c. Register tool in `main.py`

```python
@mcp.tool(
    annotations=ToolAnnotations(
        title="Find Multiple Texts",
        readOnlyHint=True,
    ),
)
def find_texts_in_document(filename: str, texts_to_find: list[str], match_case: bool = True,
                           include_paragraph_text: bool = False):
    """Find occurrences of multiple text strings in a document in one call.

    More efficient than multiple find_text_in_document calls — loads the document once
    and searches for all strings in a single pass.

    Returns a dict keyed by search string, each containing occurrences and total_count.
    """
    return extended_document_tools.find_texts_in_document_tool(
        filename, texts_to_find, match_case, include_paragraph_text
    )
```

### Step 5: Run tests — verify they PASS

```bash
uv run pytest tests/test_batch_find_texts.py -v
```

### Step 6: Run full test suite

```bash
uv run pytest tests/ -v
```

### Step 7: Update `changes.md`

Append to `changes.md`:

```markdown
### 12. `find_texts_in_document` tool — batch text search
**Branch:** `feat/batch-find-texts`
**Files:** `extended_document_utils.py`, `extended_document_tools.py`, `main.py`
**Issue resolved:** 6

New MCP tool that searches for multiple text strings in a single document load. Returns a dict
keyed by search string, each with occurrences and total_count. Supports match_case and
include_paragraph_text options. Replaces the pattern of making many individual find_text calls.
```

### Step 8: Commit and push

```bash
git add tests/test_batch_find_texts.py word_document_server/utils/extended_document_utils.py word_document_server/tools/extended_document_tools.py word_document_server/main.py changes.md
git commit -m "feat: add batch find_texts_in_document tool (Issue 6)"
git push -u origin feat/batch-find-texts
```
