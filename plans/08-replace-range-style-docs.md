# Plan 08: `replace_paragraph_range` style behavior + docs (Issues 8 + 13)

**Branch:** `fix/replace-paragraph-range-style-docs`
**Issues:** 8 (bug — empty spacers undocumented), 13 (workaround — style behavior unclear)

---

## Context

**Issue 8:** When using `replace_paragraph_range`, you must explicitly include empty string entries in `new_paragraphs` to preserve spacing between paragraphs. This is not documented.

**Issue 13:** `replace_paragraph_range` has a `style` parameter but no `preserve_style` option. When `style` is not provided, new paragraphs appear to get "Normal" style by default, but this is undocumented. It's unclear what style new paragraphs receive.

This plan adds a `preserve_style` parameter and documents both the spacer and style behaviors.

---

## Repository Info

- **Repository:** `<REPO_ROOT>`
- **Run commands with:** `uv run pytest` (or `python -m pytest` if uv unavailable)

## Architecture

- **Utility function:** `word_document_server/utils/document_utils.py` → `replace_paragraph_range()`
- **Tool wrapper:** `word_document_server/tools/content_tools.py` → `replace_paragraph_range_tool()`
- **MCP registration:** `word_document_server/main.py` inside `register_tools()`

### Current `replace_paragraph_range()` signature
```python
def replace_paragraph_range(doc_path, start_index, end_index, new_paragraphs, style=None):
```

When `style` is None, the code uses a default style (likely "Normal" from `doc.add_paragraph(text)`).

---

## Files to Modify

1. `word_document_server/utils/document_utils.py` — add `preserve_style` parameter to `replace_paragraph_range()`
2. `word_document_server/tools/content_tools.py` — pass `preserve_style` through `replace_paragraph_range_tool()`
3. `word_document_server/main.py` — update tool signature and docstring
4. `changes.md` — append entry after all tests pass

## Files to Create

1. `tests/test_replace_range_style.py`

---

## TDD Workflow

### Step 1: Create branch

```bash
git checkout -b fix/replace-paragraph-range-style-docs
```

### Step 2: Write failing tests

Create `tests/test_replace_range_style.py`:

```python
"""Tests for replace_paragraph_range style behavior."""
import pytest
from docx import Document
from word_document_server.utils.document_utils import replace_paragraph_range


class TestReplaceRangeDefaultStyle:
    """Default style behavior when no style/preserve_style specified."""

    def test_default_style_is_normal(self, make_docx):
        """No style or preserve_style: new paragraphs get Normal style."""
        path = make_docx(paragraphs=[
            {"style": "Heading 2", "runs": [{"text": "Old heading"}]},
            "Old content",
        ])
        replace_paragraph_range(path, 0, 1, ["New text"])
        doc = Document(path)
        assert doc.paragraphs[0].style.name == "Normal"


class TestReplaceRangePreserveStyle:
    """preserve_style=True behavior."""

    def test_preserve_heading_style(self, make_docx):
        """Replace Heading 2 paragraphs with preserve_style=True, new paras get Heading 2."""
        path = make_docx(paragraphs=[
            {"style": "Heading 2", "runs": [{"text": "Old heading"}]},
            "Old content",
            "Another paragraph",
        ])
        replace_paragraph_range(path, 0, 0, ["New heading text"], preserve_style=True)
        doc = Document(path)
        assert doc.paragraphs[0].style.name == "Heading 2"
        assert doc.paragraphs[0].text == "New heading text"

    def test_style_param_overrides_preserve(self, make_docx):
        """Explicit style param takes precedence over preserve_style."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Title"}]},
            "Content",
        ])
        replace_paragraph_range(path, 0, 0, ["New text"],
                                style="Heading 3", preserve_style=True)
        doc = Document(path)
        assert doc.paragraphs[0].style.name == "Heading 3"

    def test_preserve_style_multiple_paragraphs(self, make_docx):
        """All new paragraphs get the preserved style."""
        path = make_docx(paragraphs=[
            {"style": "Heading 2", "runs": [{"text": "Old"}]},
            "Content after",
        ])
        replace_paragraph_range(path, 0, 0, ["New A", "New B"], preserve_style=True)
        doc = Document(path)
        assert doc.paragraphs[0].style.name == "Heading 2"
        assert doc.paragraphs[1].style.name == "Heading 2"

    def test_preserve_false_same_as_default(self, make_docx):
        """preserve_style=False behaves same as default (Normal)."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Title"}]},
        ])
        replace_paragraph_range(path, 0, 0, ["New text"], preserve_style=False)
        doc = Document(path)
        assert doc.paragraphs[0].style.name == "Normal"


class TestReplaceRangeEmptySpacers:
    """Empty string handling for spacer paragraphs (Issue 8)."""

    def test_empty_string_creates_spacer(self, make_docx):
        """Empty string in new_paragraphs creates empty paragraph."""
        path = make_docx(paragraphs=["Old content"])
        replace_paragraph_range(path, 0, 0, ["Content", "", "More content"])
        doc = Document(path)
        texts = [p.text for p in doc.paragraphs]
        assert texts == ["Content", "", "More content"]

    def test_spacers_get_style(self, make_docx):
        """Empty spacer paragraphs also get the specified/preserved style."""
        path = make_docx(paragraphs=[
            {"style": "Heading 2", "runs": [{"text": "Title"}]},
        ])
        replace_paragraph_range(path, 0, 0, ["Text", ""], preserve_style=True)
        doc = Document(path)
        # Even the empty spacer should have the preserved style
        assert doc.paragraphs[1].style.name == "Heading 2"


class TestReplaceRangeEmptyList:
    """Empty new_paragraphs list behavior."""

    def test_empty_list_deletes_range(self, make_docx):
        """new_paragraphs=[] effectively deletes the range."""
        path = make_docx(paragraphs=["A", "B", "C", "D"])
        replace_paragraph_range(path, 1, 2, [])
        doc = Document(path)
        texts = [p.text for p in doc.paragraphs]
        assert texts == ["A", "D"]
```

### Step 3: Run tests — verify they FAIL

```bash
uv run pytest tests/test_replace_range_style.py -v
```

Tests for `preserve_style` should fail (parameter doesn't exist). Some default tests may pass.

### Step 4: Implement

#### 4a. Modify `replace_paragraph_range()` in `document_utils.py`

Update the function signature:
```python
def replace_paragraph_range(doc_path, start_index, end_index, new_paragraphs, style=None, preserve_style=False):
```

Update the style resolution logic near the top of the function (before deletion/insertion):

```python
# Determine style for new paragraphs
if style:
    style_to_use = style
elif preserve_style:
    style_to_use = doc.paragraphs[start_index].style.name if doc.paragraphs[start_index].style else "Normal"
else:
    style_to_use = None  # will use doc.add_paragraph default (Normal)
```

Ensure the insertion loop uses `style_to_use`:
```python
for text in new_paragraphs:
    new_para = doc.add_paragraph(text, style=style_to_use)
    ...
```

#### 4b. Update async wrapper in `content_tools.py`

```python
async def replace_paragraph_range_tool(filename, start_index, end_index, new_paragraphs, style=None, preserve_style=False):
    """Replace a range of paragraphs with new paragraphs."""
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}"

    return replace_paragraph_range(filename, start_index, end_index, new_paragraphs, style, preserve_style)
```

#### 4c. Update tool registration in `main.py`

```python
@mcp.tool(
    annotations=ToolAnnotations(
        title="Replace Paragraph Range",
        destructiveHint=True,
    ),
)
def replace_paragraph_range(filename: str, start_index: int, end_index: int,
                            new_paragraphs: list[str], style: str = None,
                            preserve_style: bool = False):
    """Replace a range of paragraphs (start to end index inclusive) with new paragraphs.

    Style behavior:
    - Default: new paragraphs receive 'Normal' style
    - style parameter: applies specified style to all new paragraphs
    - preserve_style=True: copies style from the paragraph at start_index
    - style parameter takes precedence over preserve_style

    Note: Empty paragraphs used as spacing in the original document must be
    explicitly included as "" entries in new_paragraphs to preserve spacing.
    Passing new_paragraphs=[] effectively deletes the range.
    """
    return content_tools.replace_paragraph_range_tool(
        filename, start_index, end_index, new_paragraphs, style, preserve_style
    )
```

### Step 5: Run tests — verify they PASS

```bash
uv run pytest tests/test_replace_range_style.py -v
```

### Step 6: Run full test suite

```bash
uv run pytest tests/ -v
```

### Step 7: Update `changes.md`

Append to `changes.md`:

```markdown
### 14. `replace_paragraph_range` — preserve_style and documentation
**Branch:** `fix/replace-paragraph-range-style-docs`
**Files:** `document_utils.py`, `content_tools.py`, `main.py`
**Issues resolved:** 8, 13

Added `preserve_style` boolean parameter (default False) to `replace_paragraph_range`.
When True, copies the paragraph style from the first replaced paragraph to all new paragraphs.
Explicit `style` parameter overrides `preserve_style`. Updated tool docstring to document:
default style behavior (Normal), empty spacer paragraph requirement ("" entries), and that
passing an empty `new_paragraphs` list effectively deletes the range.
```

### Step 8: Commit and push

```bash
git add tests/test_replace_range_style.py word_document_server/utils/document_utils.py word_document_server/tools/content_tools.py word_document_server/main.py changes.md
git commit -m "fix: add preserve_style and document style behavior in replace_paragraph_range (Issues 8, 13)"
git push -u origin fix/replace-paragraph-range-style-docs
```
