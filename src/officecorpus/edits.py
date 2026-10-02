"""The one edit every adapter applies to a corpus file, and the content model it implies.

`target(fmt, path)` reads the INPUT package (standard library only, relationships
followed by type, never by part name) and returns the edit spec every adapter receives,
or None when the file has no valid target (excluded for every adapter):

  xlsx  {"format": "xlsx", "kind": "set-cell", "sheet", "cell": "A<n>", "value"}
        sheet: the first worksheet in workbook order whose part exists (chartsheets,
        dialogsheets and macrosheets are skipped); n = the largest row number in its
        sheetData (0 when empty) + 2, so the cell is always new.
  pptx  {"format": "pptx", "kind": "set-run-text", "slide", "shape_id", "paragraph",
        "run", "text"}
        slides in presentation order (`slide` is 1-based); the first top-level p:sp of
        p:spTree (groups and mc:AlternateContent are not entered) holding a paragraph
        with a direct a:r whose a:t is non-empty; the first such paragraph and run.
  docx  {"format": "docx", "kind": "set-run-text", "paragraph", "run", "text"}
        the first direct w:body/w:p with a direct w:r holding a non-empty w:t; the
        first such run. Indexes count direct children only.

`expected(fmt, input_path, input_model, edit)` returns the content model a correct
edit produces. xlsx: the input model with the new string cell added (`cells`), so
the new cell is expected to look like an absent cell styled by its row or column, as
Excel styles it; `new_cell_default_style` detects (unscored) an output whose new cell
carries the workbook default style instead.
pptx/docx: the run-text change is applied to a copy of the input package (the run
keeps a:rPr/w:rPr; every other child of the run is replaced by one a:t/w:t holding
the marker) and that package's model is extracted, because the models merge
equal-format runs and derive other features (hyperlink text, comment anchors)
from run text, so only the model's own extraction knows every observable effect.

`applied(fmt, expected_model, output_model, edit)` decides the edit-applied check.
"""

from __future__ import annotations

import copy
import json
import posixpath
import re
import tempfile
import zipfile
from pathlib import Path
from typing import Any
from urllib.parse import unquote
from xml.etree import ElementTree as ET

MARKER = "officelibs edit 7f3a"
# Where the docx model puts an anchored drawing or object in run text.
OBJECT_PLACEHOLDER = "\ufffc"
MAX_ROW = 1048576
_ROOT_RELS = "_rels/.rels"
_REL_TYPE = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"


def _local(tag: Any) -> str:
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def _rid(element: Any) -> str | None:
    """The r:id attribute (transitional or strict relationships namespace)."""
    for name, value in element.attrib.items():
        if name.startswith("{") and name.endswith("}id") and "relationships" in name:
            return value
    return None


class _Package:
    def __init__(self, path: Path) -> None:
        with zipfile.ZipFile(path) as archive:
            self.data = {
                info.filename: archive.read(info)
                for info in archive.infolist()
                if not info.is_dir()
            }
        self._lower = {unquote(name).lower(): name for name in self.data}

    def name(self, part: str) -> str | None:
        return self._lower.get(unquote(part.lstrip("/")).lower())

    def xml(self, part: str) -> Any:
        return ET.fromstring(self.data[part])

    def rels(self, source: str) -> dict[str, tuple[str, str]]:
        """Relationship id -> (type suffix, resolved existing part name) for `source`."""
        directory, base = posixpath.split(source)
        rels_name = self.name(
            _ROOT_RELS
            if source == ""
            else posixpath.join(directory, "_rels", base + ".rels")
        )
        if rels_name is None:
            return {}
        out: dict[str, tuple[str, str]] = {}
        for rel in self.xml(rels_name):
            if _local(rel.tag) != "Relationship" or rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target") or ""
            if target.startswith("/"):
                resolved = target.lstrip("/")
            else:
                resolved = posixpath.normpath(posixpath.join(directory, target))
            found = self.name(resolved)
            if found is not None:
                out[rel.get("Id") or ""] = (
                    (rel.get("Type") or "").rsplit("/", 1)[-1],
                    found,
                )
        return out

    def main_part(self) -> str | None:
        for kind, part in self.rels("").values():
            if kind == "officeDocument":
                return part
        return None


def _children(element: Any, name: str) -> list[Any]:
    return [child for child in element if _local(child.tag) == name]


def _child(element: Any, name: str) -> Any:
    found = _children(element, name)
    return found[0] if found else None


def _cell_row(ref: str) -> int | None:
    match = re.fullmatch(r"\$?[A-Za-z]{1,3}\$?(\d+)", ref.strip())
    return int(match.group(1)) if match else None


def _xlsx_target(package: _Package) -> dict[str, Any] | None:
    workbook = package.main_part()
    if workbook is None:
        return None
    rels = package.rels(workbook)
    sheets = _child(package.xml(workbook), "sheets")
    for sheet in [] if sheets is None else _children(sheets, "sheet"):
        rel = rels.get(_rid(sheet) or "")
        if rel is None or rel[0] != "worksheet":
            continue
        data = _child(package.xml(rel[1]), "sheetData")
        largest, previous = 0, 0
        for row in [] if data is None else _children(data, "row"):
            number = int(row.get("r")) if row.get("r") else previous + 1
            previous = number
            largest = max(largest, number)
            for cell in _children(row, "c"):
                largest = max(largest, _cell_row(cell.get("r") or "") or 0)
        if largest + 2 > MAX_ROW:
            return None
        return {
            "format": "xlsx",
            "kind": "set-cell",
            "sheet": sheet.get("name"),
            "cell": f"A{largest + 2}",
            "value": MARKER,
        }
    return None


def _run_index(paragraph: Any, run_tag: str) -> int | None:
    """Index among direct runs of the first run with a non-empty direct text child."""
    for index, run in enumerate(_children(paragraph, run_tag)):
        if any(t.text for t in _children(run, "t")):
            return index
    return None


def _slides(package: _Package) -> list[str]:
    presentation = package.main_part()
    if presentation is None:
        return []
    rels = package.rels(presentation)
    ids = _child(package.xml(presentation), "sldIdLst")
    slides = []
    for sld in [] if ids is None else _children(ids, "sldId"):
        rel = rels.get(_rid(sld) or "")
        if rel is not None and rel[0] == "slide":
            slides.append(rel[1])
    return slides


def _pptx_target(package: _Package) -> dict[str, Any] | None:
    for number, slide in enumerate(_slides(package), start=1):
        csld = _child(package.xml(slide), "cSld")
        tree = None if csld is None else _child(csld, "spTree")
        for shape in [] if tree is None else _children(tree, "sp"):
            nv = _child(shape, "nvSpPr")
            cnvpr = None if nv is None else _child(nv, "cNvPr")
            body = _child(shape, "txBody")
            if cnvpr is None or body is None or not (cnvpr.get("id") or "").isdigit():
                continue
            for p_index, paragraph in enumerate(_children(body, "p")):
                run = _run_index(paragraph, "r")
                if run is not None:
                    return {
                        "format": "pptx",
                        "kind": "set-run-text",
                        "slide": number,
                        "shape_id": int(cnvpr.get("id")),
                        "paragraph": p_index,
                        "run": run,
                        "text": MARKER,
                    }
    return None


def _docx_target(package: _Package) -> dict[str, Any] | None:
    document = package.main_part()
    if document is None:
        return None
    body = _child(package.xml(document), "body")
    for p_index, paragraph in enumerate([] if body is None else _children(body, "p")):
        run = _run_index(paragraph, "r")
        if run is not None:
            return {
                "format": "docx",
                "kind": "set-run-text",
                "paragraph": p_index,
                "run": run,
                "text": MARKER,
            }
    return None


def target(fmt: str, input_path: Path) -> dict[str, Any] | None:
    """The edit spec for one input file, or None when it has no valid target."""
    package = _Package(Path(input_path))
    return {"xlsx": _xlsx_target, "pptx": _pptx_target, "docx": _docx_target}[fmt](
        package
    )


# --- expected model ---------------------------------------------------------------


def _edit_run(run: Any, text: str) -> None:
    """Set a run's text: keep its properties, replace every other child with one text."""
    ns = run.tag[: -len(_local(run.tag))]
    for child in list(run):
        if _local(child.tag) != "rPr":
            run.remove(child)
    t = run.makeelement(ns + "t", {})
    t.text = text
    run.append(t)


def _edited_package(
    fmt: str, input_path: Path, edit: dict[str, Any], out: Path
) -> None:
    from lxml import (
        etree,
    )  # lxml keeps prefixes and declarations; stdlib ET renames them

    package = _Package(input_path)
    if fmt == "docx":
        part = package.main_part()
        assert part is not None
        root = etree.fromstring(package.data[part])
        body = _child(root, "body")
        run = _children(_children(body, "p")[edit["paragraph"]], "r")[edit["run"]]
    else:
        part = _slides(package)[edit["slide"] - 1]
        root = etree.fromstring(package.data[part])
        tree = _child(_child(root, "cSld"), "spTree")
        shape = next(
            sp
            for sp in _children(tree, "sp")
            if _child(_child(sp, "nvSpPr"), "cNvPr").get("id") == str(edit["shape_id"])
        )
        paragraph = _children(_child(shape, "txBody"), "p")[edit["paragraph"]]
        run = _children(paragraph, "r")[edit["run"]]
    _edit_run(run, edit["text"])
    edited = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    with zipfile.ZipFile(input_path) as src, zipfile.ZipFile(out, "w") as dst:
        for info in src.infolist():
            dst.writestr(
                info,
                edited if info.filename == part else src.read(info),
                compress_type=zipfile.ZIP_DEFLATED,
            )


def expected(
    fmt: str, input_path: Path, input_model: dict[str, Any], edit: dict[str, Any]
) -> dict[str, Any]:
    """The content model a correct application of `edit` to the input produces."""
    if fmt == "xlsx":
        model = copy.deepcopy(input_model)
        model["cells"].setdefault(edit["sheet"], {})[edit["cell"]] = {
            "type": "s",
            "value": edit["value"],
        }
        return model
    from .content import model_module

    with tempfile.TemporaryDirectory(prefix="officecorpus-expected-") as tmp:
        edited = Path(tmp) / f"expected.{fmt}"
        _edited_package(fmt, Path(input_path), edit, edited)
        model = model_module(fmt).extract(edited)
    return json.loads(json.dumps(model, ensure_ascii=False))


def new_cell_default_style(
    fmt: str,
    expected_model: dict[str, Any],
    output_model: dict[str, Any],
    edit: dict[str, Any],
) -> bool:
    """Whether the new xlsx cell carries the workbook default style, not the inherited one.

    The edit sets a value and says nothing about style. Excel gives a new cell the row
    or column style an absent cell shows there, which is what `expected` assumes.
    Writing the cell without a style index gives it the workbook default (cellXfs[0])
    instead. Where the two differ (a styled column or row), that output still differs
    from the expected model under `cell_styles`; this only names the cause, as an
    unscored diagnostic.
    """
    if fmt != "xlsx":
        return False
    out_styles = output_model.get("cell_styles") or {}
    default = out_styles.get("default")
    sheet_styles = (out_styles.get("sheets") or {}).get(edit["sheet"]) or {}
    exp_styles = expected_model.get("cell_styles") or {}
    return (
        default is not None
        and sheet_styles.get(edit["cell"]) == default
        and exp_styles.get("default") == default
    )


# --- edit-applied ------------------------------------------------------------------


def _escape(key: str) -> str:
    return key.replace("~", "~0").replace("/", "~1")


def _paragraphs(value: Any, path: str) -> list[tuple[str, str]]:
    """(pointer, text) of every model paragraph (a dict with a `runs` list) under value."""
    found: list[tuple[str, str]] = []
    if isinstance(value, dict):
        runs = value.get("runs")
        if isinstance(runs, list):
            found.append(
                (path, "".join(r.get("text", "") for r in runs if isinstance(r, dict)))
            )
        for key, item in value.items():
            found += _paragraphs(item, f"{path}/{_escape(str(key))}")
    elif isinstance(value, list):
        for index, item in enumerate(value):
            found += _paragraphs(item, f"{path}/{index}")
    return found


def _scope(fmt: str, model: dict[str, Any], edit: dict[str, Any]) -> tuple[Any, str]:
    if fmt == "pptx":
        slides = model.get("slides") or []
        index = edit["slide"] - 1
        return (slides[index] if index < len(slides) else None), f"/slides/{index}"
    return model.get("body"), "/body"


def _without_objects(text: str) -> str:
    return text.replace(OBJECT_PLACEHOLDER, "")


# Line and page breaks as the models spell them (\v, \f) versus a UNO paragraph
# string (\n, and no character for a page break, where Writer splits the paragraph).
_LOCATOR_TEXT = str.maketrans(
    {"\v": "\n", "\r": "\n", "\u2028": "\n", "\f": None, OBJECT_PLACEHOLDER: None}
)


def comparable_text(text: str) -> str:
    """Paragraph text for checking an adapter-side locator against the input model."""
    return text.translate(_LOCATOR_TEXT)


def target_paragraph(
    fmt: str,
    input_model: dict[str, Any],
    expected_model: dict[str, Any],
    edit: dict[str, Any],
) -> dict[str, Any] | None:
    """Model path and INPUT text of the paragraph the edit targets (pptx/docx).

    The path is where the expected model carries the marker; the text is the input
    model's paragraph at that path. An adapter that has to locate the target by its
    own means (LibreOffice's UNO model has no OOXML ids) is checked against it.
    """
    if fmt == "xlsx":
        return None
    scope, prefix = _scope(fmt, expected_model, edit)
    path = next((p for p, text in _paragraphs(scope, prefix) if MARKER in text), None)
    if path is None:
        raise ValueError("the expected model does not carry the marker")
    in_scope, _ = _scope(fmt, input_model, edit)
    return {"path": path, "text": dict(_paragraphs(in_scope, prefix)).get(path)}


def applied(
    fmt: str,
    expected_model: dict[str, Any],
    output_model: dict[str, Any],
    edit: dict[str, Any],
) -> dict[str, Any]:
    """edit-applied: the output shows the edit at its target.

    xlsx: the target cell holds the marker as a string. pptx/docx: the paragraph that
    carries the marker in the expected model has the same text in the output, at the
    same model path or, if earlier content moved, anywhere on the same slide / in the
    body (`found_at` records where). Object placeholders (U+FFFC, where the models put
    an anchored drawing or object in run text) are ignored here: whether an anchor kept
    its position is content preservation, scored by content-preserved, not the edit.
    """
    if fmt == "xlsx":
        cell = output_model.get("cells", {}).get(edit["sheet"], {}).get(edit["cell"])
        actual = (
            None
            if cell is None
            else {"type": cell.get("type"), "value": cell.get("value")}
        )
        want = {"type": "s", "value": edit["value"]}
        return {
            "applied": actual == want,
            "path": f"/cells/{_escape(edit['sheet'])}/{edit['cell']}",
            "expected": want,
            "actual": actual,
        }
    scope, prefix = _scope(fmt, expected_model, edit)
    want = next(
        ((path, text) for path, text in _paragraphs(scope, prefix) if MARKER in text),
        None,
    )
    if want is None:
        raise ValueError("the expected model does not carry the marker")
    out_scope, _ = _scope(fmt, output_model, edit)
    out_paragraphs = dict(_paragraphs(out_scope, prefix))
    actual = out_paragraphs.get(want[0])
    target = _without_objects(want[1])
    found_at = (
        want[0]
        if actual is not None and _without_objects(actual) == target
        else next(
            (p for p, t in out_paragraphs.items() if _without_objects(t) == target),
            None,
        )
    )
    return {
        "applied": found_at is not None,
        "path": want[0],
        "expected": want[1],
        "actual": actual,
        "found_at": found_at,
    }
