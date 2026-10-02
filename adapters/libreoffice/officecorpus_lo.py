"""LibreOffice adapter launcher: one-edit round trip through headless soffice.

Usage: python adapters/libreoffice/officecorpus_lo.py FMT SOURCE TARGET --edit EDIT.json

Reads the input package to translate the frozen OOXML edit into a UNO locator
(sheet position, shape position, LibreOffice paragraph ordinal, run character offset), installs
``officecorpus_uno.py`` into a fresh LibreOffice profile, runs it as a user macro,
and replays its status as JSON lines on stdout:

    {"event": "opened"}
    {"event": "locator", ...}     # how the target was found in the UNO model
    {"event": "edited"}
    {"event": "saved"}
    {"event": "error", "phase": "open"|"edit"|"save", "type", "message", "traceback"}

Exit status: 0 saved, 2 usage, 3 open failure, 5 edit failure, 4 save failure.
"""

from __future__ import annotations

import argparse
import json
import os
import posixpath
import shutil
import signal
import subprocess
import sys
import tempfile
import traceback
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Any

SCRIPT = Path(__file__).with_name("officecorpus_uno.py")
MACRO_URL = "vnd.sun.star.script:officecorpus_uno.py$run?language=Python&location=user"
MAC_SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
EXIT_CODES = {"open": 3, "edit": 5, "save": 4}

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

# Top-level p:spTree children that LibreOffice imports as one draw-page shape each.
DRAWABLE = {
    P + "sp",
    P + "grpSp",
    P + "graphicFrame",
    P + "cxnSp",
    P + "pic",
    P + "contentPart",
    MC + "AlternateContent",
}
# Body-level wrappers whose w:p LibreOffice imports as ordinary body paragraphs.
DOCX_BLOCK_WRAPPERS = {
    W + "sdt",
    W + "sdtContent",
    W + "customXml",
    W + "ins",
    W + "moveTo",
}
# Run content that occupies one character, as LibreOffice reads it back.
DOCX_CHARS = {
    W + "tab": "\t",
    W + "br": "\n",
    W + "cr": "\n",
    W + "noBreakHyphen": "\u2011",
    W + "softHyphen": "\u00ad",
}
# Breaks LibreOffice turns into a paragraph property, splitting the paragraph there.
DOCX_SPLIT_BREAKS = {"page", "column"}
# Subtrees that carry no paragraph text (properties, field codes, drawings). Deleted
# text stays: LibreOffice keeps tracked deletions in the paragraph text.
DOCX_SKIP = {
    W + "pPr",
    W + "rPr",
    W + "instrText",
    W + "drawing",
    W + "pict",
    W + "object",
    MC + "AlternateContent",
}


def emit(**payload: Any) -> None:
    sys.stdout.write(json.dumps(payload) + "\n")
    sys.stdout.flush()


class LocatorError(LookupError):
    """The frozen edit does not resolve to a target inside the input package."""


def _xml(package: zipfile.ZipFile, name: str) -> ET.Element:
    return ET.fromstring(package.read(name))


def _splits(node: ET.Element) -> bool:
    return node.tag == W + "br" and node.get(W + "type") in DOCX_SPLIT_BREAKS


def _has_content(run: ET.Element) -> bool:
    ignored = {W + "rPr", W + "lastRenderedPageBreak"}
    return any(child.tag not in ignored and not _splits(child) for child in run)


def _docx_segments(
    element: ET.Element, mark: ET.Element | None = None
) -> tuple[list[str], tuple[int, int] | None]:
    """Texts of the LibreOffice paragraphs one w:p imports as, plus where ``mark`` starts.

    A page or column break with run content on both sides splits the paragraph; the
    mark position is (segment index, character offset inside that segment).
    """
    texts: list[list[str]] = [[]]
    content = [False]
    where = None

    def walk(node: ET.Element) -> None:
        nonlocal where
        if node is mark:
            where = (len(texts) - 1, len("".join(texts[-1])))
        if node.tag in DOCX_SKIP:
            return
        if node.tag == W + "r" and _has_content(node):
            content[-1] = True
        if node.tag in {W + "t", W + "delText"}:
            texts[-1].append(node.text or "")
        elif _splits(node):
            texts.append([])
            content.append(False)
        elif node.tag in DOCX_CHARS:
            texts[-1].append(DOCX_CHARS[node.tag])
        elif node.tag == W + "sym":
            texts[-1].append(chr(int(node.get(W + "char", "0"), 16)))
        for child in node:
            walk(child)

    walk(element)
    first, last = 0, len(texts) - 1
    while first < last and not content[first]:
        first += 1
    while last > first and not content[last]:
        last -= 1
    if where is not None:
        where = (min(max(where[0], first), last) - first, where[1])
    return ["".join(parts) for parts in texts[first : last + 1]], where


def _lo_paragraph_count(paragraph: ET.Element) -> int:
    """How many Writer body paragraphs one block-level w:p becomes."""
    segments, _ = _docx_segments(paragraph)
    section_break = paragraph.find(f"{W}pPr/{W}sectPr") is not None
    if section_break and segments == [""] and paragraph.find(f".//{W}r") is None:
        return 0  # Writer drops an empty paragraph that only carries a section break.
    return len(segments)


def _docx_locator(package: zipfile.ZipFile, edit: dict[str, Any]) -> dict[str, Any]:
    body = _xml(package, "word/document.xml").find(W + "body")
    if body is None:
        raise LocatorError("word/document.xml has no w:body")
    paragraphs = body.findall(W + "p")
    if not 0 <= edit["paragraph"] < len(paragraphs):
        raise LocatorError(f"body paragraph {edit['paragraph']} is missing")
    target = paragraphs[edit["paragraph"]]
    runs = target.findall(W + "r")
    if not 0 <= edit["run"] < len(runs):
        raise LocatorError(f"run {edit['run']} is missing from body paragraph")
    run = runs[edit["run"]]

    ordinal = 0

    def count(node: ET.Element) -> bool:
        nonlocal ordinal
        for child in node:
            if child is target:
                return True
            if child.tag == W + "p":
                ordinal += _lo_paragraph_count(child)
            elif child.tag in DOCX_BLOCK_WRAPPERS and count(child):
                return True
        return False

    count(body)
    segments, (segment, offset) = _docx_segments(target, mark=run)
    run_text = "".join(_docx_segments(run)[0])
    return {
        "lo_paragraph": ordinal + segment,
        "offset": offset,
        "length": len(run_text),
        "run_text": run_text,
        "paragraph_text": segments[segment],
    }


def _pptx_slide_part(package: zipfile.ZipFile, slide: int) -> str:
    presentation = _xml(package, "ppt/presentation.xml")
    ids = presentation.findall(f"{P}sldIdLst/{P}sldId")
    if not 1 <= slide <= len(ids):
        raise LocatorError(f"slide {slide} is missing from p:sldIdLst")
    rel_id = ids[slide - 1].get(R + "id")
    for rel in _xml(package, "ppt/_rels/presentation.xml.rels"):
        if rel.get("Id") == rel_id:
            target = rel.get("Target", "")
            if target.startswith("/"):
                return target.lstrip("/")
            return posixpath.normpath(posixpath.join("ppt", target))
    raise LocatorError(f"slide relationship {rel_id} is missing")


def _pptx_paragraph_text(paragraph: ET.Element, stop: ET.Element | None = None) -> str:
    parts: list[str] = []
    for child in paragraph:
        if child is stop:
            break
        if child.tag in {A + "r", A + "fld"}:
            parts.append(child.findtext(A + "t") or "")
        elif child.tag == A + "br":
            parts.append("\n")
    return "".join(parts)


def _pptx_locator(package: zipfile.ZipFile, edit: dict[str, Any]) -> dict[str, Any]:
    tree = _xml(package, _pptx_slide_part(package, edit["slide"])).find(
        f"{P}cSld/{P}spTree"
    )
    if tree is None:
        raise LocatorError("slide has no p:spTree")
    index = 0
    for child in tree:
        if child.tag not in DRAWABLE:
            continue
        properties = (
            child.find(f"{P}nvSpPr/{P}cNvPr") if child.tag == P + "sp" else None
        )
        if properties is not None and properties.get("id") == str(edit["shape_id"]):
            shape, name = child, properties.get("name", "")
            break
        index += 1
    else:
        raise LocatorError(f"top-level p:sp with id {edit['shape_id']} is missing")
    paragraphs = shape.findall(f"{P}txBody/{A}p")
    if not 0 <= edit["paragraph"] < len(paragraphs):
        raise LocatorError(f"paragraph {edit['paragraph']} is missing from the shape")
    paragraph = paragraphs[edit["paragraph"]]
    runs = paragraph.findall(A + "r")
    if not 0 <= edit["run"] < len(runs):
        raise LocatorError(f"run {edit['run']} is missing from the paragraph")
    run = runs[edit["run"]]
    run_text = run.findtext(A + "t") or ""
    return {
        "shape_index": index,
        "shape_name": name,
        "offset": len(_pptx_paragraph_text(paragraph, stop=run)),
        "length": len(run_text),
        "run_text": run_text,
        "paragraph_text": _pptx_paragraph_text(paragraph),
    }


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _xlsx_locator(package: zipfile.ZipFile, edit: dict[str, Any]) -> dict[str, Any]:
    """The target sheet's position among the workbook's sheets, and their count.

    The UNO script uses the position only when no UNO sheet carries the target name
    (exactly or case-insensitively) and UNO has as many sheets as the workbook.
    """
    workbook = next(
        (
            rel.get("Target", "")
            for rel in _xml(package, "_rels/.rels")
            if rel.get("Type", "").endswith("/officeDocument")
        ),
        None,
    )
    if workbook is None:
        raise LocatorError("the package has no officeDocument relationship")
    root = _xml(package, posixpath.normpath(workbook.lstrip("/")))
    container = next((c for c in root if _local(c.tag) == "sheets"), None)
    names = [
        s.get("name", "")
        for s in (container if container is not None else [])
        if _local(s.tag) == "sheet"
    ]
    if edit["sheet"] not in names:
        raise LocatorError(f"sheet {edit['sheet']!r} is missing from the workbook")
    return {"sheet_index": names.index(edit["sheet"]), "sheet_count": len(names)}


def locate(fmt: str, source: Path, edit: dict[str, Any]) -> dict[str, Any]:
    if edit.get("format") != fmt:
        raise LocatorError(f"edit format {edit.get('format')!r} does not match {fmt!r}")
    with zipfile.ZipFile(source) as package:
        if fmt == "xlsx":
            return _xlsx_locator(package, edit)
        if fmt == "docx":
            return _docx_locator(package, edit)
        return _pptx_locator(package, edit)


def soffice() -> str:
    return shutil.which("soffice") or MAC_SOFFICE


def run_soffice(
    source: Path, target: Path, task: dict[str, Any]
) -> tuple[dict | None, int, str]:
    """Run the UNO script in a fresh profile; return (status, exit code, output tail)."""
    with tempfile.TemporaryDirectory(prefix="officecorpus-lo-") as scratch:
        root = Path(scratch)
        profile = root / "profile"
        scripts = profile / "user" / "Scripts" / "python"
        scripts.mkdir(parents=True)
        shutil.copyfile(SCRIPT, scripts / SCRIPT.name)
        task_path = root / "task.json"
        task_path.write_text(json.dumps(task), encoding="utf-8")
        status_path = root / "status.json"
        environment = {
            **os.environ,
            "OFFICECORPUS_UNO_INPUT": str(source.resolve()),
            "OFFICECORPUS_UNO_OUTPUT": str(target.resolve()),
            "OFFICECORPUS_UNO_TASK": str(task_path),
            "OFFICECORPUS_UNO_STATUS": str(status_path),
        }
        completed = subprocess.run(
            [
                soffice(),
                f"-env:UserInstallation={profile.as_uri()}",
                "--headless",
                "--invisible",
                "--norestore",
                "--nologo",
                "--nolockcheck",
                MACRO_URL,
            ],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            env=environment,
            check=False,
        )
        tail = completed.stdout.decode("utf-8", "replace")[-2000:]
        status = (
            json.loads(status_path.read_text("utf-8"))
            if status_path.is_file()
            else None
        )
    return status, completed.returncode, tail


def _error_payload(phase: str, exc: BaseException) -> dict[str, str]:
    kind = type(exc)
    return {
        "phase": phase,
        "type": kind.__qualname__
        if kind is LocatorError
        else f"{kind.__module__}.{kind.__qualname__}",
        "message": str(exc)[:2000],
        "traceback": traceback.format_exc()[-4000:],
    }


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("fmt", choices=["xlsx", "pptx", "docx"])
    parser.add_argument("source", type=Path)
    parser.add_argument("target", type=Path)
    parser.add_argument("--edit", type=Path, required=True)
    args = parser.parse_args(argv)
    # The runner stops the process group with SIGTERM; unwind so the profile is removed.
    signal.signal(signal.SIGTERM, lambda *_: sys.exit(143))

    try:
        edit = json.loads(args.edit.read_text("utf-8"))
        locator = locate(args.fmt, args.source, edit)
    except (
        LookupError,
        KeyError,
        TypeError,
        ValueError,
        OSError,
        ET.ParseError,
        zipfile.BadZipFile,
    ) as exc:
        emit(event="error", **_error_payload("edit", exc))
        return EXIT_CODES["edit"]

    status, returncode, tail = run_soffice(
        args.source, args.target, {"edit": edit, "locator": locator}
    )
    if status is None:
        emit(
            event="error",
            phase="open",
            type="NoStatus",
            message=f"soffice exited {returncode} without a script status: {tail[-1000:]}",
            traceback="",
        )
        return EXIT_CODES["open"]

    steps = status.get("steps", [])
    found = status.get("locator")
    for step in steps:
        emit(event=step)
        if step == "opened" and (found is not None or locator):
            details = {k: v for k, v in locator.items() if k != "paragraph_text"}
            if "run_text" in details:
                details["run_text"] = details["run_text"][:200]
            emit(event="locator", **details, **(found or {}))
    error = status.get("error")
    if error is None and "saved" not in steps:
        phase = (
            "open"
            if "opened" not in steps
            else "edit"
            if "edited" not in steps
            else "save"
        )
        error = {
            "phase": phase,
            "type": "Incomplete",
            "message": f"soffice exited {returncode} after steps {steps}: {tail[-1000:]}",
            "traceback": "",
        }
    if error is not None:
        emit(event="error", **error)
        return EXIT_CODES.get(error.get("phase"), 1)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
