"""Comment management tools for Word Document Server.

Provides tools for creating and managing comments in Word documents
using the docx-editor library. Complements the existing read-only
comment extraction tools in comment_tools.py.
"""
import os
import json
import getpass
from typing import Optional

from docx_editor import Document as DocxEditorDocument
from docx_editor.exceptions import TextNotFoundError, CommentError

from word_document_server.utils.file_utils import ensure_docx_extension, check_file_writeable
from word_document_server.utils.docx_zip_utils import (
    strip_meta_json,
    strip_orphan_comments_extensible,
)
from word_document_server.utils.anchor_utils import normalize_paragraph_runs_for_anchor


def _get_author(author: Optional[str]) -> str:
    """Get the author name, defaulting to the system username."""
    if author and author.strip():
        return author.strip()
    return getpass.getuser()


def _open_tracked_document(filename: str, author: str) -> DocxEditorDocument:
    """Open a document for editing."""
    return DocxEditorDocument.open(filename, author=author, force_recreate=True)


def _save_and_sanitize(doc: DocxEditorDocument, filename: str) -> None:
    """Save via docx-editor, then strip its stray meta.json from the zip.

    docx-editor writes its workspace meta.json inside the unpacked OPC tree
    and zips it into the .docx, which Word flags as "unreadable content".
    """
    doc.save()
    strip_meta_json(filename)


async def add_comment(
    filename: str,
    anchor_text: str,
    comment_text: str,
    author: str = None,
) -> str:
    """Add a comment anchored to specific text in a document.

    Args:
        filename: Path to the Word document
        anchor_text: Text to attach the comment to
        comment_text: The comment content
        author: Author name for the comment (defaults to system username)

    Returns:
        JSON string with result including comment_id
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})

    writeable, error = check_file_writeable(filename)
    if not writeable:
        return json.dumps({"success": False, "error": error})

    if not comment_text or not comment_text.strip():
        return json.dumps({"success": False, "error": "comment_text cannot be empty"})

    author = _get_author(author)
    # Normalize runs so a cross-run anchor resolves inside docx-editor's
    # single-run matcher. Returns False if the anchor is genuinely absent.
    if not normalize_paragraph_runs_for_anchor(filename, anchor_text):
        return json.dumps({"success": False, "error": f"Anchor text '{anchor_text}' not found in document"})
    doc = None
    try:
        doc = _open_tracked_document(filename, author)
        comment_id = doc.add_comment(anchor_text, comment_text)
        _save_and_sanitize(doc, filename)
        return json.dumps({"success": True, "comment_id": comment_id, "anchor_text": anchor_text, "author": author})
    except TextNotFoundError:
        return json.dumps({"success": False, "error": f"Anchor text '{anchor_text}' not found in document"})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to add comment: {str(e)}"})
    finally:
        if doc:
            try:
                doc.close(cleanup=True)
            except Exception:
                pass


async def reply_to_comment(
    filename: str,
    comment_id: int,
    reply_text: str,
    author: str = None,
) -> str:
    """Add a reply to an existing comment.

    Args:
        filename: Path to the Word document
        comment_id: ID of the comment to reply to
        reply_text: The reply content
        author: Author name for the reply (defaults to system username)

    Returns:
        JSON string with result including new reply comment_id
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})

    writeable, error = check_file_writeable(filename)
    if not writeable:
        return json.dumps({"success": False, "error": error})

    author = _get_author(author)
    doc = None
    try:
        doc = _open_tracked_document(filename, author)
        reply_id = doc.reply_to_comment(comment_id, reply_text)
        _save_and_sanitize(doc, filename)
        return json.dumps({"success": True, "reply_id": reply_id, "parent_comment_id": comment_id, "author": author})
    except CommentError as e:
        return json.dumps({"success": False, "error": str(e)})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to reply to comment: {str(e)}"})
    finally:
        if doc:
            try:
                doc.close(cleanup=True)
            except Exception:
                pass


async def resolve_comment(
    filename: str,
    comment_id: int,
) -> str:
    """Mark a comment as resolved.

    Args:
        filename: Path to the Word document
        comment_id: ID of the comment to resolve

    Returns:
        JSON string with result
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})

    writeable, error = check_file_writeable(filename)
    if not writeable:
        return json.dumps({"success": False, "error": error})

    doc = None
    try:
        doc = _open_tracked_document(filename, _get_author(None))
        result = doc.resolve_comment(comment_id)
        if result:
            _save_and_sanitize(doc, filename)
            return json.dumps({"success": True, "resolved": comment_id})
        else:
            return json.dumps({"success": False, "error": f"Comment {comment_id} not found"})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to resolve comment: {str(e)}"})
    finally:
        if doc:
            try:
                doc.close(cleanup=True)
            except Exception:
                pass


async def delete_comment(
    filename: str,
    comment_id: int,
) -> str:
    """Delete a comment from a document.

    Args:
        filename: Path to the Word document
        comment_id: ID of the comment to delete

    Returns:
        JSON string with result
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})

    writeable, error = check_file_writeable(filename)
    if not writeable:
        return json.dumps({"success": False, "error": error})

    doc = None
    try:
        doc = _open_tracked_document(filename, _get_author(None))
        result = doc.delete_comment(comment_id)
        if result:
            _save_and_sanitize(doc, filename)
            # docx-editor's delete_comment skips commentsExtensible.xml; scrub
            # any leftover <w16cex:commentExtensible> with no live durableId.
            strip_orphan_comments_extensible(filename)
            return json.dumps({"success": True, "deleted": comment_id})
        else:
            return json.dumps({"success": False, "error": f"Comment {comment_id} not found"})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to delete comment: {str(e)}"})
    finally:
        if doc:
            try:
                doc.close(cleanup=True)
            except Exception:
                pass
