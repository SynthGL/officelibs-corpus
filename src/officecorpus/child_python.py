"""Child process for Python adapters: open one file, apply the edit, save it.

Runs inside the adapter's own virtual environment, so it imports nothing from
officecorpus. Usage: `child_python.py <adapter> <source> <target> --edit <edit.json>`.
Protocol: JSON lines on stdout ({"event": "opened"}, {"event": "edited"},
{"event": "saved"}, or {"event": "error", "phase": "open" | "edit" | "save", ...});
`--version <adapter>` prints the installed distribution version.

Every edit goes through the library's public API for that operation (see EDIT_API).
"""

import json
import os
import posixpath
import subprocess
import sys
import tempfile
import traceback
import zipfile
from importlib import metadata
from pathlib import Path
from xml.etree import ElementTree as ET

DISTRIBUTIONS = {
    "openpyxl": "openpyxl",
    "wolfxl": "wolfxl",
    "python-pptx": "python-pptx",
    "wolfppt-wheel": "wolfppt",
    "python-docx": "python-docx",
    "docxtpl": "docxtpl",
    "aspose-cells-foss": "aspose-cells-foss",
    "wolfdocx-wheel": "wolfdocx",
}
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def emit(**payload):
    sys.stdout.write(json.dumps(payload) + "\n")
    sys.stdout.flush()


def open_document(adapter, source):
    if adapter == "openpyxl":
        import openpyxl

        return openpyxl.load_workbook(source)
    if adapter == "wolfxl":
        import wolfxl

        # modify=True is WolfXL's documented mode for editing an existing workbook.
        return wolfxl.load_workbook(source, modify=True)
    if adapter == "python-pptx":
        from pptx import Presentation

        return Presentation(source)
    if adapter == "wolfppt-wheel":
        from wolfppt import Presentation

        return Presentation(source)
    if adapter == "python-docx":
        import docx

        return docx.Document(source)
    if adapter == "docxtpl":
        from docxtpl import DocxTemplate

        return DocxTemplate(source)
    if adapter == "aspose-cells-foss":
        from aspose.cells_foss import Workbook

        return Workbook(source)
    raise SystemExit(f"unknown adapter {adapter}")


def _slide_run(presentation, edit):
    """python-pptx-compatible API: slide by order, top-level p:sp by shape_id, a:r by index."""
    slide = presentation.slides[edit["slide"] - 1]
    shape = next(
        s
        for s in slide.shapes
        if s.shape_id == edit["shape_id"] and s.element.tag.endswith("}sp")
    )
    return shape.text_frame.paragraphs[edit["paragraph"]].runs[edit["run"]]


def apply_edit(adapter, document, edit):
    if adapter in ("openpyxl", "wolfxl"):
        document[edit["sheet"]][edit["cell"]] = edit["value"]
    elif adapter == "aspose-cells-foss":
        sheet = document.get_worksheet_by_name(edit["sheet"])
        sheet.cells[edit["cell"]].put_value(edit["value"])
    elif adapter in ("python-pptx", "wolfppt-wheel"):
        _slide_run(document, edit).text = edit["text"]
    elif adapter == "python-docx":
        # Document.paragraphs and Paragraph.runs are the direct w:p / w:r children.
        document.paragraphs[edit["paragraph"]].runs[edit["run"]].text = edit["text"]
    elif adapter == "docxtpl":
        # get_docx() is docxtpl's documented handle for editing the base document;
        # it is a python-docx Document, so the edit is python-docx's Run.text.
        # save() without render() reloads the template from disk and drops edits made
        # through get_docx(), so the documented flow renders (empty context) first.
        paragraph = document.get_docx().paragraphs[edit["paragraph"]]
        paragraph.runs[edit["run"]].text = edit["text"]
        document.render({})
    else:
        raise SystemExit(f"unknown adapter {adapter}")


def error(phase, exc):
    emit(
        event="error",
        phase=phase,
        type=type(exc).__module__ + "." + type(exc).__qualname__,
        message=str(exc)[:2000],
        traceback=traceback.format_exc()[-4000:],
    )


class WolfdocxRefused(Exception):
    pass


def _wolfdocx_cli(*args):
    cli = Path(sys.executable).parent / "wolfdocx"
    done = subprocess.run(
        [str(cli), *args], capture_output=True, text=True, check=False
    )
    try:
        return json.loads(done.stdout)
    except json.JSONDecodeError:
        raise WolfdocxRefused(
            f"wolfdocx {args[0]} {args[1]} exited {done.returncode}: "
            f"{(done.stderr or done.stdout)[-1500:]}"
        ) from None


def _docx_locator(source, edit):
    """Main part name, 1-based index among all w:p of that part, and the run's text."""
    with zipfile.ZipFile(source) as archive:
        names = {n.lower(): n for n in archive.namelist()}
        rels = ET.fromstring(archive.read(names["_rels/.rels"]))
        main = next(
            r.get("Target").lstrip("/")
            for r in rels
            if (r.get("Type") or "").endswith("/officeDocument")
        )
        part = names[posixpath.normpath(main).lower()]
        root = ET.fromstring(archive.read(part))
    body = next(c for c in root if c.tag.endswith("}body"))
    paragraph = [c for c in body if c.tag.endswith("}p")][edit["paragraph"]]
    run = [c for c in paragraph if c.tag.endswith("}r")][edit["run"]]
    ns = paragraph.tag[: -len("p")]
    index = list(root.iter(ns + "p")).index(paragraph) + 1
    text = "".join(t.text or "" for t in run if t.tag == ns + "t")
    return part, index, text


def run_wolfdocx(source, target, edit):
    """WolfDocx `author report-targets` then `author report` (documented CLI)."""
    try:
        part, index, old_text = _docx_locator(source, edit)
        found = _wolfdocx_cli("author", "report-targets", str(source), "--part", part)
        if not found.get("passed"):
            raise WolfdocxRefused(f"report-targets refused: {found.get('code')}")
    except BaseException as exc:  # noqa: BLE001 - every failure is evidence
        error("open", exc)
        return 3
    emit(event="opened")
    try:
        candidates = [
            t
            for t in found.get("targets", [])
            if t.get("kind") == "text" and t.get("paragraph") == index
        ]
        if not candidates:
            raise WolfdocxRefused(
                f"paragraph {index} of {part} is not a report-authoring text target"
            )
        with tempfile.TemporaryDirectory(prefix="officecorpus-wolfdocx-") as tmp:
            request = Path(tmp) / "request.json"
            request.write_text(
                json.dumps(
                    {
                        "source": str(source),
                        "destination": str(target),
                        "source_sha256": found["source_sha256"],
                        "edits": [
                            {
                                "kind": "text",
                                "part": part,
                                "paragraph": index,
                                "expected_old_text": old_text,
                                "structure_fingerprint_sha256": candidates[0][
                                    "structure_fingerprint_sha256"
                                ],
                                "text": edit["text"],
                            }
                        ],
                    }
                ),
                encoding="utf-8",
            )
            report = _wolfdocx_cli("author", "report", str(request))
        if not report.get("passed") or not os.path.exists(target):
            raise WolfdocxRefused(f"author report refused: {report.get('code')}")
    except BaseException as exc:  # noqa: BLE001
        error("edit", exc)
        return 4
    emit(event="edited")
    emit(event="saved")
    return 0


def main():
    if sys.argv[1] == "--version":
        emit(event="version", version=metadata.version(DISTRIBUTIONS[sys.argv[2]]))
        return 0
    adapter, source, target, flag, edit_path = sys.argv[1:6]
    if flag != "--edit":
        raise SystemExit(
            "usage: child_python.py ADAPTER SOURCE TARGET --edit EDIT_JSON"
        )
    with open(edit_path, encoding="utf-8") as handle:
        edit = json.load(handle)
    if adapter == "wolfdocx-wheel":
        return run_wolfdocx(source, target, edit)
    try:
        document = open_document(adapter, source)
    except BaseException as exc:  # noqa: BLE001 - every failure is evidence
        error("open", exc)
        return 3
    emit(event="opened")
    try:
        apply_edit(adapter, document, edit)
    except BaseException as exc:  # noqa: BLE001
        error("edit", exc)
        return 4
    emit(event="edited")
    try:
        document.save(target)
    except BaseException as exc:  # noqa: BLE001
        error("save", exc)
        return 5
    emit(event="saved")
    return 0


if __name__ == "__main__":
    sys.exit(main())
