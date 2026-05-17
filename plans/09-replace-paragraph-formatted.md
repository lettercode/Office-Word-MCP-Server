# Plan 09: `replace_paragraph_text` with markdown formatting (Issue 9)

**Branch:** `feat/replace-paragraph-formatted-text`
**Issue:** 9

---

## Context

When using `replace_paragraph_text` to replace a paragraph's content, all inline formatting (bold, italic) is lost. The replacement text becomes a single unformatted run. `preserve_style` only preserves paragraph-level style (e.g., "Normal", "Heading 2"), not character-level formatting. This is a problem for academic papers where statistical variables like *F* and *p* must be italic.

This plan adds a `parse_markdown` parameter that interprets `*italic*`, `**bold**`, and `***bold italic***` in the replacement text and creates properly formatted runs.

---

## Repository Info

- **Repository:** `<REPO_ROOT>`
- **Run commands with:** `uv run pytest` (or `python -m pytest` if uv unavailable)

## Architecture

- **Utility function:** `word_document_server/utils/document_utils.py` → `replace_paragraph_text()`
- **Tool wrapper:** `word_document_server/tools/content_tools.py` → `replace_paragraph_text_tool()`
- **MCP registration:** `word_document_server/main.py` inside `register_tools()`

### Current `replace_paragraph_text()` behavior
1. Clears all existing runs (sets text to "")
2. Sets the first run's text to `new_text` (or adds a new run)
3. Result: single unformatted run

### Note on duplicate function
`replace_paragraph_text()` is duplicated in `document_utils.py` (appears at ~line 496 and ~line 841). Only modify the first (primary) copy. If the second copy is called anywhere, it should be updated too or removed.

---

## Files to Modify

1. `word_document_server/utils/document_utils.py` — add `_parse_markdown_runs()` helper and `parse_markdown` parameter to `replace_paragraph_text()`
2. `word_document_server/tools/content_tools.py` — pass `parse_markdown` through
3. `word_document_server/main.py` — update tool signature and docstring
4. `changes.md` — append entry after all tests pass

## Files to Create

1. `tests/test_replace_paragraph_formatted.py`

---

## TDD Workflow

### Step 1: Create branch

```bash
git checkout -b feat/replace-paragraph-formatted-text
```

### Step 2: Write failing tests

Create `tests/test_replace_paragraph_formatted.py`:

```python
"""Tests for replace_paragraph_text with markdown formatting support."""
import pytest
from docx import Document
from word_document_server.utils.document_utils import replace_paragraph_text, _parse_markdown_runs


class TestParseMarkdownRuns:
    """Unit tests for the markdown parser."""

    def test_plain_text(self):
        """No markdown returns single plain run."""
        runs = _parse_markdown_runs("Hello world")
        assert len(runs) == 1
        assert runs[0]["text"] == "Hello world"
        assert runs[0]["bold"] is False
        assert runs[0]["italic"] is False

    def test_single_italic(self):
        """'Hello *world*' returns 2 runs, second is italic."""
        runs = _parse_markdown_runs("Hello *world*")
        assert len(runs) == 2
        assert runs[0]["text"] == "Hello "
        assert runs[0]["italic"] is False
        assert runs[1]["text"] == "world"
        assert runs[1]["italic"] is True
        assert runs[1]["bold"] is False

    def test_single_bold(self):
        """'Hello **world**' returns 2 runs, second is bold."""
        runs = _parse_markdown_runs("Hello **world**")
        assert len(runs) == 2
        assert runs[1]["text"] == "world"
        assert runs[1]["bold"] is True
        assert runs[1]["italic"] is False

    def test_bold_italic(self):
        """'Hello ***world***' returns 2 runs, second is bold+italic."""
        runs = _parse_markdown_runs("Hello ***world***")
        assert len(runs) == 2
        assert runs[1]["text"] == "world"
        assert runs[1]["bold"] is True
        assert runs[1]["italic"] is True

    def test_mixed_formatting(self):
        """'The *F* statistic (*p* < .001)' produces correct runs."""
        runs = _parse_markdown_runs("The *F* statistic (*p* < .001)")
        # Expected: "The " + italic "F" + " statistic (" + italic "p" + " < .001)"
        texts = [r["text"] for r in runs]
        assert "".join(texts) == "The F statistic (p < .001)"
        # Find italic runs
        italic_runs = [r for r in runs if r["italic"]]
        italic_texts = [r["text"] for r in italic_runs]
        assert "F" in italic_texts
        assert "p" in italic_texts

    def test_adjacent_bold_and_italic(self):
        """'**bold** and *italic*' produces 3 runs."""
        runs = _parse_markdown_runs("**bold** and *italic*")
        texts = [r["text"] for r in runs]
        assert "".join(texts) == "bold and italic"
        bold_runs = [r for r in runs if r["bold"]]
        assert any(r["text"] == "bold" for r in bold_runs)
        italic_runs = [r for r in runs if r["italic"]]
        assert any(r["text"] == "italic" for r in italic_runs)

    def test_empty_string(self):
        """Empty string returns single empty run."""
        runs = _parse_markdown_runs("")
        assert len(runs) == 1
        assert runs[0]["text"] == ""

    def test_no_closing_marker(self):
        """Unclosed marker is treated as literal text."""
        runs = _parse_markdown_runs("Hello *world")
        full_text = "".join(r["text"] for r in runs)
        assert full_text == "Hello *world"

    def test_asterisks_in_math(self):
        """Single asterisks not used as markers when not paired."""
        runs = _parse_markdown_runs("2 * 3 = 6")
        full_text = "".join(r["text"] for r in runs)
        assert full_text == "2 * 3 = 6"


class TestReplaceWithMarkdown:
    """Integration tests: markdown formatting applied to Word document."""

    def test_italic_run_in_document(self, make_docx):
        """parse_markdown=True creates italic run in saved document."""
        path = make_docx(paragraphs=["Original text"])
        replace_paragraph_text(path, 0, "Hello *world*", parse_markdown=True)
        doc = Document(path)
        para = doc.paragraphs[0]
        assert para.text == "Hello world"
        # Find the italic run
        italic_runs = [r for r in para.runs if r.italic]
        assert len(italic_runs) >= 1
        assert any(r.text == "world" for r in italic_runs)

    def test_bold_run_in_document(self, make_docx):
        """parse_markdown=True creates bold run in saved document."""
        path = make_docx(paragraphs=["Original text"])
        replace_paragraph_text(path, 0, "Hello **world**", parse_markdown=True)
        doc = Document(path)
        para = doc.paragraphs[0]
        bold_runs = [r for r in para.runs if r.bold]
        assert any(r.text == "world" for r in bold_runs)

    def test_bold_italic_in_document(self, make_docx):
        """parse_markdown=True creates bold+italic run."""
        path = make_docx(paragraphs=["Original"])
        replace_paragraph_text(path, 0, "Hello ***world***", parse_markdown=True)
        doc = Document(path)
        para = doc.paragraphs[0]
        bi_runs = [r for r in para.runs if r.bold and r.italic]
        assert any(r.text == "world" for r in bi_runs)

    def test_preserves_paragraph_style(self, make_docx):
        """Heading 2 style preserved with parse_markdown=True."""
        path = make_docx(paragraphs=[
            {"style": "Heading 2", "runs": [{"text": "Old heading"}]},
        ])
        replace_paragraph_text(path, 0, "New *heading*", preserve_style=True, parse_markdown=True)
        doc = Document(path)
        assert doc.paragraphs[0].style.name == "Heading 2"

    def test_default_no_parsing(self, make_docx):
        """parse_markdown=False (default) treats asterisks as literal text."""
        path = make_docx(paragraphs=["Original"])
        replace_paragraph_text(path, 0, "Hello *world*")
        doc = Document(path)
        assert doc.paragraphs[0].text == "Hello *world*"

    def test_statistical_text(self, make_docx):
        """Realistic: 'ANOVA (*F*(2, 38) = 108.37, *p* < .001)' formats correctly."""
        path = make_docx(paragraphs=["Old text"])
        new_text = "ANOVA (*F*(2, 38) = 108.37, *p* < .001)"
        replace_paragraph_text(path, 0, new_text, parse_markdown=True)
        doc = Document(path)
        para = doc.paragraphs[0]
        # Full text should have no asterisks
        assert para.text == "ANOVA (F(2, 38) = 108.37, p < .001)"
        # F and p should be italic
        italic_texts = [r.text for r in para.runs if r.italic]
        assert "F" in italic_texts
        assert "p" in italic_texts

    def test_multiple_paragraphs_unchanged(self, make_docx):
        """Other paragraphs in the document are not affected."""
        path = make_docx(paragraphs=["Keep me", "Replace me", "Keep me too"])
        replace_paragraph_text(path, 1, "New *italic* text", parse_markdown=True)
        doc = Document(path)
        assert doc.paragraphs[0].text == "Keep me"
        assert doc.paragraphs[2].text == "Keep me too"
```

### Step 3: Run tests — verify they FAIL

```bash
uv run pytest tests/test_replace_paragraph_formatted.py -v
```

### Step 4: Implement

#### 4a. Add `_parse_markdown_runs()` helper to `document_utils.py`

Add near the top of the file (after imports and before the main functions):

```python
import re

def _parse_markdown_runs(text: str) -> list:
    """Parse markdown-style inline formatting into a list of run specs.

    Supports:
    - ***bold italic*** (must be checked first — greedy)
    - **bold**
    - *italic*

    Returns:
        List of dicts: [{"text": "...", "bold": bool, "italic": bool}, ...]
    """
    if not text:
        return [{"text": "", "bold": False, "italic": False}]

    # Pattern matches ***bold italic***, **bold**, or *italic* (greedy, non-nested)
    # Order matters: check *** before ** before *
    pattern = r'(\*{3}(.+?)\*{3}|\*{2}(.+?)\*{2}|\*(.+?)\*)'

    runs = []
    last_end = 0

    for match in re.finditer(pattern, text):
        start = match.start()

        # Add plain text before this match
        if start > last_end:
            runs.append({
                "text": text[last_end:start],
                "bold": False,
                "italic": False
            })

        # Determine which group matched
        if match.group(2) is not None:
            # *** bold italic ***
            runs.append({
                "text": match.group(2),
                "bold": True,
                "italic": True
            })
        elif match.group(3) is not None:
            # ** bold **
            runs.append({
                "text": match.group(3),
                "bold": True,
                "italic": False
            })
        elif match.group(4) is not None:
            # * italic *
            runs.append({
                "text": match.group(4),
                "bold": False,
                "italic": True
            })

        last_end = match.end()

    # Add remaining text after last match
    if last_end < len(text):
        runs.append({
            "text": text[last_end:],
            "bold": False,
            "italic": False
        })

    # If no matches found, return the whole text as plain
    if not runs:
        runs.append({"text": text, "bold": False, "italic": False})

    return runs
```

#### 4b. Modify `replace_paragraph_text()` in `document_utils.py`

Update the function signature (update BOTH copies if they exist, or remove the duplicate):

```python
def replace_paragraph_text(doc_path, paragraph_index, new_text, preserve_style=True, parse_markdown=False):
```

Update the implementation:

```python
def replace_paragraph_text(doc_path, paragraph_index, new_text, preserve_style=True, parse_markdown=False):
    """Replace the text of a paragraph, optionally parsing markdown-style formatting.

    Args:
        doc_path: Path to the Word document
        paragraph_index: Index of the paragraph to replace
        new_text: New text content
        preserve_style: Preserve paragraph-level style (default True)
        parse_markdown: Parse *italic*, **bold**, ***bold italic*** (default False)
    """
    import os
    if not os.path.exists(doc_path):
        return f"Error: Document {doc_path} does not exist"

    try:
        doc = Document(doc_path)

        if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
            return f"Error: Invalid paragraph index {paragraph_index}. Document has {len(doc.paragraphs)} paragraphs."

        para = doc.paragraphs[paragraph_index]
        old_style = para.style

        # Clear all existing runs
        for run in para.runs:
            run.text = ""
        # Remove empty run elements
        for run in para.runs:
            if hasattr(run, '_r') and run._r.getparent() is not None:
                run._r.getparent().remove(run._r)

        if parse_markdown:
            # Parse markdown and create formatted runs
            run_specs = _parse_markdown_runs(new_text)
            for spec in run_specs:
                run = para.add_run(spec["text"])
                if spec["bold"]:
                    run.bold = True
                if spec["italic"]:
                    run.italic = True
        else:
            # Original behavior: single unformatted run
            para.add_run(new_text)

        if preserve_style and old_style:
            para.style = old_style

        doc.save(doc_path)
        return f"Successfully replaced paragraph {paragraph_index}."
    except Exception as e:
        return f"Error: Failed to replace paragraph text: {str(e)}"
```

**Important:** Check if there is a duplicate of `replace_paragraph_text` in `document_utils.py`. If so, update both or remove the duplicate to prevent confusion.

#### 4c. Update async wrapper in `content_tools.py`

```python
async def replace_paragraph_text_tool(filename, paragraph_index, new_text, preserve_style=True, parse_markdown=False):
    """Replace the text of a specific paragraph."""
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}"

    return replace_paragraph_text(filename, paragraph_index, new_text, preserve_style, parse_markdown)
```

#### 4d. Update tool registration in `main.py`

```python
@mcp.tool(
    annotations=ToolAnnotations(
        title="Replace Paragraph Text",
        destructiveHint=True,
    ),
)
def replace_paragraph_text(filename: str, paragraph_index: int, new_text: str,
                          preserve_style: bool = True, parse_markdown: bool = False):
    """Replace the text of a specific paragraph by index, optionally preserving its style.

    When parse_markdown is True, supports inline formatting:
    - *italic* for italic text
    - **bold** for bold text
    - ***bold italic*** for bold italic text

    Note: preserve_style preserves paragraph-level style (e.g., Heading 2).
    Character-level formatting from the original paragraph is not preserved.
    Use parse_markdown=True to apply formatting to the replacement text.
    """
    return content_tools.replace_paragraph_text_tool(
        filename, paragraph_index, new_text, preserve_style, parse_markdown
    )
```

### Step 5: Run tests — verify they PASS

```bash
uv run pytest tests/test_replace_paragraph_formatted.py -v
```

### Step 6: Run full test suite

```bash
uv run pytest tests/ -v
```

### Step 7: Update `changes.md`

Append to `changes.md`:

```markdown
### 15. `replace_paragraph_text` — markdown formatting support
**Branch:** `feat/replace-paragraph-formatted-text`
**Files:** `document_utils.py`, `content_tools.py`, `main.py`
**Issue resolved:** 9

Added `parse_markdown` boolean parameter (default False) to `replace_paragraph_text`.
When True, interprets `*italic*`, `**bold**`, and `***bold italic***` in the replacement
text and creates properly formatted Word runs. Added `_parse_markdown_runs()` helper
that splits text into run specs using regex. Backward compatible — default behavior unchanged.
```

### Step 8: Commit and push

```bash
git add tests/test_replace_paragraph_formatted.py word_document_server/utils/document_utils.py word_document_server/tools/content_tools.py word_document_server/main.py changes.md
git commit -m "feat: add parse_markdown option to replace_paragraph_text (Issue 9)"
git push -u origin feat/replace-paragraph-formatted-text
```
