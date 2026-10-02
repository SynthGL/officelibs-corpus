"""officecorpus Python-UNO script: open one document, apply one edit, save it.

LibreOffice runs this file inside its own embedded Python as a user-profile script
(``vnd.sun.star.script:officecorpus_uno.py$run?language=Python&location=user``);
``officecorpus_lo.py`` installs it into a fresh profile for every invocation.
Arguments arrive through environment variables because script URLs carry none:

- ``OFFICECORPUS_UNO_INPUT`` / ``OFFICECORPUS_UNO_OUTPUT``: absolute document paths
- ``OFFICECORPUS_UNO_TASK``: JSON file ``{"edit": {...}, "locator": {...}}`` holding the
  frozen edit plus the launcher-computed UNO locator
- ``OFFICECORPUS_UNO_STATUS``: JSON status file this script rewrites after every step:
  ``{"steps": [...], "locator": {...} | null, "error": {...} | null}``

UNO has no notion of OOXML runs (LibreOffice merges equal-attribute runs into one
portion), so text edits address the run as a character range inside its paragraph:
the launcher's offset is tried first and verified against the run's original text;
otherwise the unique occurrence of that text in the paragraph is used.
"""

from __future__ import annotations

import json
import os
import traceback

import uno
from com.sun.star.beans import PropertyValue

FILTERS = {
    "xlsx": "Calc MS Excel 2007 XML",
    "pptx": "Impress MS PowerPoint 2007 XML",
    "docx": "MS Word 2007 XML",
}


class LocateError(LookupError):
    """The edit target could not be identified unambiguously in the UNO model."""


class VerifyError(RuntimeError):
    """The edited range does not read back as the marker."""


def _property(name, value):
    item = PropertyValue()
    item.Name = name
    item.Value = value
    return item


def _write_status(path, status):
    scratch = path + ".partial"
    with open(scratch, "w", encoding="utf-8") as handle:
        json.dump(status, handle)
    os.replace(scratch, path)


def _paragraphs(text):
    enumeration = text.createEnumeration()
    while enumeration.hasMoreElements():
        element = enumeration.nextElement()
        if element.supportsService("com.sun.star.text.Paragraph"):
            yield element


def _nth_paragraph(text, index):
    for position, paragraph in enumerate(_paragraphs(text)):
        if position == index:
            return paragraph
    raise LocateError(f"paragraph {index} is missing from the UNO text")


class _Paragraph:
    """Character ranges of one paragraph, addressed in text-cursor units."""

    def __init__(self, text, index):
        self.text = text
        self.index = index
        self.refresh()

    def refresh(self):
        # Impress paragraph objects freeze their extent when enumerated, so a stale
        # one reads truncated text after an insertion; enumerate again after edits.
        self.paragraph = _nth_paragraph(self.text, self.index)

    def _cursor(self, count, expand):
        """Cursor from the paragraph start moved ``count`` units (goRight takes a short)."""
        cursor = self.text.createTextCursorByRange(self.paragraph.getStart())
        while count > 0:
            step = min(count, 32767)
            if not cursor.goRight(step, expand):
                return None
            count -= step
        return cursor

    def span(self, start, length):
        """Cursor over [start, start + length), or None if it leaves the paragraph."""
        # Text cursors walk across paragraph ends, and the Impress text API has no
        # range comparison: a range stays inside iff its text is a paragraph prefix.
        prefix = self._cursor(start + length, True)
        if prefix is None or not self.paragraph.getString().startswith(
            prefix.getString()
        ):
            return None
        cursor = self._cursor(start, False)
        if length:
            cursor.goRight(length, True)
        return cursor

    def occurrences(self, needle):
        """Cursor offsets whose range reads exactly ``needle`` (stops after two)."""
        found = []
        if not needle:
            return found
        start = 0
        while len(found) < 2:
            cursor = self.span(start, len(needle))
            if cursor is None:
                # A range that leaves the paragraph only moves further out from here.
                break
            if cursor.getString() == needle:
                found.append(start)
            start += 1
        return found


def _replace_run(para, locator, marker, found):
    """Replace the run's characters with ``marker`` keeping the run's attributes."""
    offset, length, run_text = locator["offset"], locator["length"], locator["run_text"]
    cursor = para.span(offset, length)
    if cursor is not None and cursor.getString() == run_text:
        method, start = "offset", offset
    else:
        hits = para.occurrences(run_text)
        if len(hits) != 1:
            raise LocateError(
                f"run text {run_text!r} not at offset {offset} and occurs "
                f"{'more than once' if hits else 'nowhere'} in the paragraph"
            )
        method, start = "unique-occurrence", hits[0]
    found.update(method=method, position=start)
    # Insert inside the run's own attribute span (after its first character) so the
    # marker inherits the run's formatting, then delete the original characters.
    insert_at = start + (1 if length else 0)
    para.text.insertString(para.span(insert_at, 0), marker, False)
    para.refresh()
    if length > 1:
        para.span(insert_at + len(marker), length - 1).setString("")
        para.refresh()
    if length:
        para.span(start, 1).setString("")
        para.refresh()
    written = para.span(start, len(marker))
    if written is None or written.getString() != marker:
        raise VerifyError("edited range does not read back as the marker")


def _sheet(sheets, name, locator):
    """The target sheet: by exact name, else the one case-insensitive name match, else
    by its workbook-order position when UNO holds as many sheets as the workbook.

    Excel sheet names are unique case-insensitively. LibreOffice's importer can also
    rename a sheet (seen: 'sheet1' read as 'sheet1_2'); the position then identifies
    it, and the saved name is the output's own, scored like any other difference.
    """
    if sheets.hasByName(name):
        return sheets.getByName(name), "name"
    names = list(sheets.getElementNames())
    folded = name.casefold()
    matches = [n for n in names if n.casefold() == folded]
    if len(matches) == 1:
        return sheets.getByName(matches[0]), "name-case-insensitive"
    if matches:
        raise LocateError(
            f"sheet {name!r} matches {len(matches)} sheets case-insensitively"
        )
    if len(names) == locator["sheet_count"]:
        return sheets.getByIndex(locator["sheet_index"]), "position"
    raise LocateError(
        f"sheet {name!r} is missing and UNO has {len(names)} sheets, the workbook "
        f"{locator['sheet_count']} (UNO sheets: {names!r})"
    )


def _edit_xlsx(document, edit, locator, found):
    sheet, found["sheet_method"] = _sheet(document.Sheets, edit["sheet"], locator)
    sheet.getCellRangeByName(edit["cell"]).setString(edit["value"])
    found["method"] = "cell"


def _shape(page, locator):
    index, name = locator["shape_index"], locator["shape_name"]
    if index < page.getCount() and page.getByIndex(index).Name == name:
        return page.getByIndex(index), "index"
    named = [
        page.getByIndex(position)
        for position in range(page.getCount())
        if page.getByIndex(position).Name == name
    ]
    if len(named) == 1:
        return named[0], "name"
    raise LocateError(
        f"shape {index} is not named {name!r} and {len(named)} shapes carry that name"
    )


def _edit_pptx(document, edit, locator, found):
    pages = document.DrawPages
    if not 1 <= edit["slide"] <= pages.getCount():
        raise LocateError(f"slide {edit['slide']} is missing")
    shape, found["shape_method"] = _shape(pages.getByIndex(edit["slide"] - 1), locator)
    para = _Paragraph(shape.getText(), edit["paragraph"])
    # The runner checks this against the input model's target paragraph.
    found["located_text"] = para.paragraph.getString()
    found["paragraph_text_match"] = (
        para.paragraph.getString() == locator["paragraph_text"]
    )
    _replace_run(para, locator, edit["text"], found)


def _writer_paragraph(text, index, expected):
    """Body paragraph index for the launcher's estimate, verified by its text.

    The launcher predicts how Writer splits and drops paragraphs on import; when the
    predicted paragraph reads differently, the nearest paragraph with exactly the
    expected text wins, else the estimate stands and the run check decides.
    """
    strings = [paragraph.getString() for paragraph in _paragraphs(text)]
    if index < len(strings) and strings[index] == expected:
        return index, "ordinal"
    distances = sorted(
        (abs(position - index), position)
        for position, value in enumerate(strings)
        if value == expected
    )
    if distances and (len(distances) == 1 or distances[0][0] != distances[1][0]):
        return distances[0][1], "nearest-text"
    if index < len(strings):
        return index, "ordinal-unverified"
    raise LocateError(f"paragraph {index} is missing and its text is not unique")


def _edit_docx(document, edit, locator, found):
    index, method = _writer_paragraph(
        document.Text, locator["lo_paragraph"], locator["paragraph_text"]
    )
    found.update(paragraph_method=method, lo_paragraph_used=index)
    para = _Paragraph(document.Text, index)
    # The runner checks this against the input model's target paragraph.
    found["located_text"] = para.paragraph.getString()
    # Edits must land as plain text, not as a tracked change.
    recording = document.RecordChanges
    document.RecordChanges = False
    try:
        _replace_run(para, locator, edit["text"], found)
    finally:
        document.RecordChanges = recording


EDITORS = {"xlsx": _edit_xlsx, "pptx": _edit_pptx, "docx": _edit_docx}


def _error(phase, exc):
    message = getattr(exc, "Message", None) or str(exc)
    kind = type(exc)
    own = kind in (LocateError, VerifyError)
    return {
        "phase": phase,
        "type": kind.__qualname__ if own else f"{kind.__module__}.{kind.__qualname__}",
        "message": message[:2000],
        "traceback": traceback.format_exc()[-4000:],
    }


def run(*_args):
    status_path = os.environ["OFFICECORPUS_UNO_STATUS"]
    status = {"steps": [], "locator": None, "error": None}
    context = uno.getComponentContext()
    desktop = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", context
    )
    phase = "open"
    try:
        with open(os.environ["OFFICECORPUS_UNO_TASK"], encoding="utf-8") as handle:
            task = json.load(handle)
        edit = task["edit"]
        document = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(os.environ["OFFICECORPUS_UNO_INPUT"]),
            "_blank",
            0,
            (_property("Hidden", True), _property("ReadOnly", False)),
        )
        if document is None:
            raise RuntimeError("LibreOffice could not load the input document")
        try:
            status["steps"].append("opened")
            _write_status(status_path, status)
            phase = "edit"
            status["locator"] = {}
            EDITORS[edit["format"]](document, edit, task["locator"], status["locator"])
            status["steps"].append("edited")
            _write_status(status_path, status)
            phase = "save"
            document.storeToURL(
                uno.systemPathToFileUrl(os.environ["OFFICECORPUS_UNO_OUTPUT"]),
                (
                    _property("FilterName", FILTERS[edit["format"]]),
                    _property("Overwrite", True),
                ),
            )
            status["steps"].append("saved")
        finally:
            document.close(True)
    except Exception as exc:  # noqa: BLE001 - every failure is evidence for the launcher
        status["error"] = _error(phase, exc)
    finally:
        _write_status(status_path, status)
        desktop.terminate()


g_exportedScripts = (run,)
