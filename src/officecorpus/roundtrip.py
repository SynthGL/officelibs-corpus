"""Real-world one-edit round trip: every adapter opens each corpus file, makes one small
precomputed edit through the library's public API, saves, and the output is scored.

The edit (officecorpus.edits) is computed once per file from the input alone and is
identical for every adapter: xlsx sets a new string cell below the first worksheet's
data; pptx and docx set the text of one existing run. A byte copy of the input cannot
pass, so every library must parse and serialize the document.

Scored checks per (adapter, file), in order:
  opened              the adapter loaded the file without an exception
  saved               the adapter applied the edit and wrote an output file without an
                      exception
  package-valid       the output is a readable OPC package (content types, root and
                      internal relationships resolve) with no problem the input lacked
  edit-applied        the output's content model shows the marker at the edit target
                      (officecorpus.edits.applied)
  content-preserved   the output's content model equals the expected model, the input's
                      model with the edit applied:
                      model.<format>.compare(expected, extract(output)) == []

Unscored diagnostics (recorded with full detail, never ranked on):
  render-similarity   LibreOffice PDF renders of input and output match (render.py
                      thresholds). Unscored because the oracle, LibreOffice, is itself a
                      library under test and renders its own output. The edit changes a
                      few words, so a correct output differs slightly from the input.
  parts-preserved     every input part has an output counterpart (same name, same
                      relationship path, or identical bytes under a new name)
  semantic-preserved  every mapped part is equal after XML canonicalization
                      (ooxml.canonical_part); the edited part always differs

A file without a valid edit target is unscored for every adapter (reason
`no-edit-target`) and no adapter runs on it. The input and expected models are
extracted once per file; an input whose model cannot be extracted is excluded from the
content checks (edit-applied and content-preserved) with the reason recorded. A file is
fully preserved when every scored check succeeds.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import os
import platform
import queue
import shutil
import subprocess
import sys
import threading
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any

from . import __version__, content, edits, ooxml, render
from .adapters import Prepared, adapters_for, get_adapter, libreoffice_version, prepare
from .corpus import files_digest, load_manifest, sha256_file, verify_holdout

SCHEMA_VERSION = 4
LANE = "real-world-roundtrip"
NO_EDIT_TARGET = "no-edit-target"
LOCATOR_MISMATCH = "locator mismatch"
# Error types the LibreOffice launcher/UNO script raise when they cannot find the target.
_LOCATOR_ERRORS = ("LocatorError", "LocateError")
CHECKS = (
    "opened",
    "saved",
    "package-valid",
    "edit-applied",
    "content-preserved",
    "render-similarity",
    "parts-preserved",
    "semantic-preserved",
)
SCORED_CHECKS = (
    "opened",
    "saved",
    "package-valid",
    "edit-applied",
    "content-preserved",
)
BYTE_ONLY_CHECKS = ("parts-preserved", "semantic-preserved")
NORMALIZATION = [
    (
        "One edit per file, identical for every adapter and computed from the input alone "
        "(src/officecorpus/edits.py): xlsx sets a new string cell A<n> on the first worksheet, n two "
        "rows below its last sheetData row; pptx and docx replace the text of the first run with "
        "visible text (pptx: first slide with a top-level p:sp holding one, docx: first top-level body "
        "paragraph holding one) with the marker. A file without such a target is unscored for every "
        "adapter (no-edit-target)."
    ),
    (
        "content-preserved compares content models: officecorpus.model.<format>.extract() reads the "
        "output into a model of what a user of the document can observe, and compare() lists every "
        "difference from the expected model; the check passes when the list is empty. The expected "
        "model is the input model with the edit applied: xlsx adds the new string cell to the input "
        "model; pptx and docx extract the model of a copy of the input whose target run holds only "
        "its properties and the marker text, so run merging and derived text follow the model's own "
        "rules."
    ),
    (
        "edit-applied passes when the output model holds the marker at the target: xlsx, the target "
        "cell is the marker string; pptx and docx, the paragraph that carries the marker in the "
        "expected model has the same text in the output, at the same model path or elsewhere on the "
        "same slide or in the body."
    ),
    (
        "xlsx new cell style: the expected new cell carries the style Excel gives it, the row or "
        "column style an absent cell shows there; any other style on it is a cell_styles difference. "
        "Unscored diagnostic: every xlsx result records in content-preserved whether the new cell "
        "carries the workbook default style (cellXfs[0], what a write without a style index gives) "
        "where that differs from the inherited style (new_cell_default_style), and each summary "
        "counts those files."
    ),
    (
        "LibreOffice pptx/docx: UNO exposes no OOXML ids, so the adapter locates the target by "
        "position. The result is unscored for LibreOffice (reason 'locator mismatch') when that lookup "
        "fails, when the pptx shape at the target position does not carry the target's name, or when "
        "the located paragraph's text differs from the input model's target paragraph text "
        "(object placeholders U+FFFC dropped, \\v \\r U+2028 read as \\n, \\f dropped). A wrong "
        "guess is never scored as a preservation failure."
    ),
    (
        "LibreOffice xlsx: the adapter looks the target sheet up by its exact name, else by the one "
        "case-insensitive match (Excel sheet names are unique case-insensitively), else by its "
        "position in workbook order when UNO holds as many sheets as the workbook (LibreOffice's "
        "importer can rename a sheet, e.g. 'sheet1' to 'sheet1_2'). Each result records which "
        "(details.locator.sheet_method); a renamed sheet in the output stays a scored difference. "
        "Several case-insensitive matches or a different sheet count is an adapter error."
    ),
    (
        "xlsx: a difference in a cell whose input cell holds a formula is attributed to the feature "
        "formula_cached_values instead of cells. Excel recomputes formulas on open, so a missing or "
        "stale cached value is invisible in Excel, but readers that do not recalculate (pandas, file "
        "previewers, search indexers) see the cached value or nothing."
    ),
    (
        "The model never records how a package is serialized: part names, relationship ids, namespace "
        "prefixes, attribute order, XML layout, and whether a default value is written or omitted do "
        "not enter it, so an equivalent rewrite scores the same as a byte copy."
    ),
    (
        "The exact equivalence rules (inheritance, defaults, colors, numbers, formulas, ranges, binary "
        "parts) are the docstring of each model module: src/officecorpus/model/xlsx.py, "
        "src/officecorpus/model/pptx.py, src/officecorpus/model/docx.py. Their FEATURES tuples name the "
        "features every difference is attributed to."
    ),
    (
        "A feature is present in an input when compare() attributes at least one difference to it after "
        "every scalar of the input model that is not null, false, or the empty string is replaced by a "
        "probe value and every non-empty list gains a probe item."
    ),
    (
        "The input and expected models are extracted once per file. An input whose model cannot be "
        "extracted is excluded from edit-applied and content-preserved (reason recorded); an output "
        "whose model cannot be extracted fails both."
    ),
    (
        "render-similarity, parts-preserved, and semantic-preserved are unscored diagnostics: the render "
        "oracle is LibreOffice, itself a library under test, and the two byte-only checks compare part "
        "names and canonical XML, which only byte copying reliably passes."
    ),
]


def _check(
    name: str, outcome: str, detail: Any = None, scored: bool | None = None
) -> dict[str, Any]:
    """One check record; only SCORED_CHECKS are scored unless `scored` says otherwise."""
    category = {
        "opened": "adapter",
        "saved": "adapter",
        "package-valid": "package",
        "edit-applied": "content",
        "content-preserved": "content",
        "render-similarity": "render",
        "parts-preserved": "byte-only",
        "semantic-preserved": "byte-only",
    }[name]
    return {
        "name": name,
        "category": category,
        "outcome": outcome,
        "scored": name in SCORED_CHECKS if scored is None else scored,
        "detail": detail,
    }


class _Redactor:
    def __init__(self, root: Path, work: Path) -> None:
        self.pairs = [
            (str(work), "<work>"),
            (str(root), "<repo>"),
            (str(Path.home()), "~"),
        ]

    def __call__(self, text: str) -> str:
        for old, new in self.pairs:
            text = text.replace(old, new)
        return text


def _events(stdout: str) -> list[dict[str, Any]]:
    events = []
    for line in stdout.splitlines():
        line = line.strip()
        if line.startswith("{"):
            try:
                events.append(json.loads(line))
            except json.JSONDecodeError:
                continue
    return events


def _require_new_or_empty(path: Path) -> None:
    if path.exists() and (
        not path.is_dir() or path.is_symlink() or any(path.iterdir())
    ):
        raise ValueError("output must be a new or empty non-symlink directory")


def _runner_identity(root: Path, fmt: str) -> str:
    digest = hashlib.sha256()
    files = sorted((root / "src/officecorpus").glob("*.py")) + [
        root / "src/officecorpus/model/__init__.py",
        root / f"src/officecorpus/model/{fmt}.py",
        root / "adapters/node/roundtrip.cjs",
        root / "adapters/node/package-lock.json",
        *sorted((root / "adapters/libreoffice").glob("*.py")),
    ]
    for path in files:
        digest.update(
            path.relative_to(root).as_posix().encode() + b"\0" + path.read_bytes()
        )
    return digest.hexdigest()


def _content_child(args: list[str], timeout: float) -> tuple[dict[str, Any], float]:
    """Run one content-model child; return its event (or a synthesized error) and wall time."""
    child = render.run_child(
        [sys.executable, "-m", "officecorpus.content", *args], timeout
    )
    events = _events(child["stdout"])
    if child["timed_out"]:
        event = {
            "event": "error",
            "type": "Timeout",
            "message": f"model extraction exceeded {timeout:g}s",
        }
    elif events:
        event = events[-1]
    else:
        event = {
            "event": "error",
            "type": "ChildFailed",
            "message": f"exit {child['returncode']}: {child['stderr'][-500:]}",
        }
    return event, child["elapsed_ms"]


def _environment() -> dict[str, Any]:
    cpu = None
    if platform.system() == "Darwin":
        try:
            cpu = subprocess.run(
                ["sysctl", "-n", "machdep.cpu.brand_string"],
                capture_output=True,
                text=True,
                check=True,
            ).stdout.strip()
        except (OSError, subprocess.CalledProcessError):
            cpu = None
    return {
        "os": f"{platform.system()} {platform.release()}",
        "platform": platform.platform(),
        "arch": platform.machine(),
        "cpu": cpu,
        "logical_cpus": os.cpu_count(),
        "python": platform.python_version(),
        "implementation": platform.python_implementation(),
    }


def run_roundtrip(
    root: Path,
    *,
    fmt: str,
    adapter_ids: list[str],
    split: str,
    output: Path,
    timeout: float,
    render_timeout: float,
    content_timeout: float,
    jobs: int,
    limit: int | None,
) -> Path:
    _require_new_or_empty(output)
    output.mkdir(parents=True, exist_ok=True)
    output = output.resolve()
    manifest = load_manifest(root)
    verify_holdout(root, manifest)
    entries = [
        f
        for f in manifest["files"]
        if f["format"] == fmt and (split == "all" or f["split"] == split)
    ]
    if limit is not None:
        entries = entries[:limit]
    if "all" in adapter_ids:
        specs = adapters_for(fmt)
    else:
        specs = [get_adapter(fmt, a) for a in dict.fromkeys(adapter_ids)]
    created = dt.datetime.now(dt.UTC)
    prepared = [prepare(root, spec) for spec in specs]

    work = output / "work"
    work.mkdir()
    redact = _Redactor(root, work)
    profiles: queue.Queue[Path] = queue.Queue()
    for index in range(max(1, jobs)):
        profile = work / "lo-profiles" / f"p{index}"
        profile.mkdir(parents=True)
        profiles.put(profile)

    def with_profile(fn: Any) -> Any:
        profile = profiles.get()
        try:
            return fn(profile)
        finally:
            profiles.put(profile)

    # Phase 1: every input once: edit spec, input and expected models (scored lane) and
    # oracle render (diagnostic). Files without an edit target are not run at all.
    input_renders: dict[str, dict[str, Any]] = {}
    input_models: dict[str, dict[str, Any]] = {}

    def prepare_input(index: int, entry: dict[str, Any]) -> None:
        case = work / "input" / f"{index:03d}"
        prepared_dir = case / "prepared"
        prepared_dir.mkdir(parents=True)
        source = case / Path(entry["path"]).name
        shutil.copyfile(root / entry["path"], source)
        event, elapsed = _content_child(
            ["prepare", fmt, str(source), str(prepared_dir)], content_timeout
        )
        model: dict[str, Any] = {
            "fixture": entry["path"],
            "elapsed_ms": elapsed,
            "edit": event.get("edit"),
        }
        if event.get("event") != "prepared":
            model.update(
                outcome=NO_EDIT_TARGET,
                reason="edit target could not be computed: "
                + redact(f"{event.get('type')}: {event.get('message', '')}")[:500],
            )
        elif event["edit"] is None:
            model.update(
                outcome=NO_EDIT_TARGET, reason="the input has no valid edit target"
            )
        elif "input_features" in event:
            model.update(
                outcome="success",
                input_features=event["input_features"],
                target_paragraph=event.get("target_paragraph"),
                prepared_dir=prepared_dir,
                edit_json=prepared_dir / "edit.json",
            )
        else:
            error = event.get("model_error") or {}
            model.update(
                outcome="excluded",
                reason="input content model could not be extracted: "
                + redact(f"{error.get('type')}: {error.get('message', '')}")[:500],
                edit_json=prepared_dir / "edit.json",
            )
        input_models[entry["path"]] = model
        if model["outcome"] == NO_EDIT_TARGET:
            return

        pdf, result = with_profile(
            lambda p: render.to_pdf(source, case / "pdf", p, render_timeout)
        )
        record: dict[str, Any] = {
            "fixture": entry["path"],
            "source_sha256": entry["sha256"],
            "elapsed_ms": result["elapsed_ms"],
            "timed_out": result["timed_out"],
        }
        if pdf is None:
            record.update(
                outcome="failure", reason="LibreOffice produced no PDF for the input"
            )
        else:
            import pymupdf

            with pymupdf.open(pdf) as doc:
                record["pages"] = doc.page_count
            record.update(outcome="success", output_sha256=sha256_file(pdf), pdf=pdf)
        input_renders[entry["path"]] = record

    with ThreadPoolExecutor(max_workers=max(1, jobs)) as pool:
        list(pool.map(lambda pair: prepare_input(*pair), enumerate(entries)))

    # Phase 2: every available adapter on every file with an edit target.
    tasks = [
        (prep, index, entry)
        for prep in prepared
        if prep.available
        for index, entry in enumerate(entries)
        if input_models[entry["path"]]["outcome"] != NO_EDIT_TARGET
    ]
    results: list[dict[str, Any] | None] = [None] * len(tasks)
    lock = threading.Lock()

    def run_task(task_index: int) -> None:
        prep, index, entry = tasks[task_index]
        result = _roundtrip_one(
            root,
            prep,
            index,
            entry,
            work,
            timeout,
            render_timeout,
            content_timeout,
            with_profile,
            input_renders,
            input_models,
            redact,
        )
        with lock:
            results[task_index] = result

    try:
        with ThreadPoolExecutor(max_workers=max(1, jobs)) as pool:
            list(pool.map(run_task, range(len(tasks))))
    finally:
        for prep in prepared:
            if prep.pool is not None:
                prep.pool.close()

    final_results = [r for r in results if r is not None]
    final_results += [
        _no_target_row(prep.spec.id, entry, input_models[entry["path"]])
        for prep in prepared
        if prep.available
        for entry in entries
        if input_models[entry["path"]]["outcome"] == NO_EDIT_TARGET
    ]
    render_evidence = [
        {k: v for k, v in input_renders[e["path"]].items() if k != "pdf"}
        for e in entries
        if e["path"] in input_renders
    ]
    model_module = content.model_module(fmt)
    features = content.features(fmt)
    private = ("prepared_dir", "edit_json")
    content_inputs = [
        {k: v for k, v in input_models[e["path"]].items() if k not in private}
        for e in entries
    ]

    report = {
        "schema_version": SCHEMA_VERSION,
        "benchmark": "OfficeCorpus",
        "benchmark_version": __version__,
        "lane": LANE,
        "format": fmt,
        "status": _status(final_results),
        "created_at_utc": created.isoformat(),
        "finished_at_utc": dt.datetime.now(dt.UTC).isoformat(),
        "environment": _environment(),
        "configuration": {
            "format": fmt,
            "split": split,
            "limit": limit,
            "adapters": [p.spec.id for p in prepared],
            "timeout_seconds": timeout,
            "render_timeout_seconds": render_timeout,
            "content_timeout_seconds": content_timeout,
            "jobs": jobs,
            "cold_subprocess": True,
            "iterations": 1,
            "isolation": (
                "each (adapter, file) runs in a fresh child process group with its own copy of the input; "
                "the group is killed on timeout. Docker adapters instead serve files from a persistent "
                "helper process per container (see each adapter's receipt.execution and receipt.timing)"
            ),
            "checks": list(CHECKS),
            "scored_checks": list(SCORED_CHECKS),
            "normalization": NORMALIZATION,
            "edit": {
                "marker": edits.MARKER,
                "source": "src/officecorpus/edits.py",
                "source_sha256": sha256_file(root / "src/officecorpus/edits.py"),
                "rules": "module docstring of the source above",
                "transport": (
                    "the spec is written once per file as edit.json; Python, Node and LibreOffice "
                    "adapters receive its path, Docker helpers receive it base64-encoded on the "
                    "request line"
                ),
            },
            "content_model": {
                "module": model_module.__name__,
                "source": f"src/officecorpus/model/{fmt}.py",
                "source_sha256": sha256_file(root / f"src/officecorpus/model/{fmt}.py"),
                "features": list(features),
                "rules": "module docstring of the source above",
            },
            "render": {
                "scored": False,
                "unscored_reason": (
                    "the oracle is LibreOffice, which is also a library under test and renders its own output"
                ),
                "oracle": "LibreOffice headless PDF export",
                "oracle_version": libreoffice_version(),
                "rasterizer": "pymupdf " + _pymupdf_version(),
                **render.thresholds(),
            },
            "runner_sha256": _runner_identity(root, fmt),
            "reproduce": _reproduce(
                fmt,
                prepared,
                split,
                timeout,
                render_timeout,
                content_timeout,
                jobs,
                limit,
            ),
        },
        "adapters": [_adapter_row(p) for p in prepared],
        "manifest": {
            "path": "manifest.json",
            "manifest_sha256": sha256_file(root / "manifest.json"),
            "files_digest": files_digest(manifest["files"]),
            "holdout_digest": manifest["holdout"]["digest"],
            "holdout_count": manifest["holdout"]["count"],
        },
        "fixtures": [
            {
                "id": e["path"],
                "path": e["path"],
                "sha256": e["sha256"],
                "size": e["size"],
                "source": e["source"],
                "license": e["license"],
                "split": e["split"],
                "features": e["features"],
            }
            for e in entries
        ],
        "content_inputs": {
            "files": len(entries),
            "no_edit_target": sum(
                1 for r in content_inputs if r["outcome"] == NO_EDIT_TARGET
            ),
            "extracted": sum(1 for r in content_inputs if r["outcome"] == "success"),
            "excluded": [
                {
                    "fixture": r["fixture"],
                    "outcome": r["outcome"],
                    "reason": r["reason"],
                }
                for r in content_inputs
                if r["outcome"] != "success"
            ],
            "inputs": content_inputs,
        },
        "results": final_results,
        "summary": [
            _summary(p, entries, final_results, input_models, features)
            for p in prepared
        ],
        "optional_evidence": {"render": render_evidence},
    }
    shutil.rmtree(work)
    run_json = output / "run.json"
    run_json.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    return run_json


def _pymupdf_version() -> str:
    import pymupdf

    return pymupdf.VersionBind


def _reproduce(
    fmt: str,
    prepared: list[Prepared],
    split: str,
    timeout: float,
    render_timeout: float,
    content_timeout: float,
    jobs: int,
    limit: int | None,
) -> str:
    parts = ["uv run officecorpus roundtrip", f"--format {fmt}"]
    parts += [f"--adapter {p.spec.id}" for p in prepared]
    parts += [
        f"--split {split}",
        f"--timeout {timeout:g}",
        f"--render-timeout {render_timeout:g}",
        f"--content-timeout {content_timeout:g}",
        f"--jobs {jobs}",
    ]
    if limit is not None:
        parts.append(f"--limit {limit}")
    parts.append("--output <new-dir>")
    return " ".join(parts)


def _adapter_row(prep: Prepared) -> dict[str, Any]:
    return {
        "name": prep.spec.id,
        "library": prep.spec.library,
        "language": prep.spec.language,
        "mode": prep.spec.mode,
        "available": prep.available,
        "supported": prep.spec.kind != "unsupported",
        "version": prep.version,
        "reason": prep.reason,
        "receipt": prep.receipt,
    }


def _status(results: list[dict[str, Any]]) -> str:
    run = [r for r in results if r["outcome"] != "unscored"]
    if not run:
        return "failure"
    return "success" if all(r["outcome"] == "success" for r in run) else "failure"


def _no_target_row(
    adapter_id: str,
    entry: dict[str, Any],
    input_model: dict[str, Any],
    reason: str = NO_EDIT_TARGET,
    detail: Any = None,
) -> dict[str, Any]:
    return {
        "adapter": adapter_id,
        "lane": LANE,
        "fixture": entry["path"],
        "input_sha256": entry["sha256"],
        "source": entry["source"],
        "outcome": "unscored",
        "reason": reason,
        "detail": input_model["reason"] if detail is None else detail,
        "edit": input_model["edit"] if reason != NO_EDIT_TARGET else None,
        "checks": [],
        "samples": [],
    }


def _locator_mismatch(
    spec: Any,
    events: list[dict[str, Any]],
    error_event: dict[str, Any] | None,
    input_model: dict[str, Any],
) -> dict[str, Any] | None:
    """Why LibreOffice's own target lookup cannot be trusted for this file, or None.

    LibreOffice's UNO model has no OOXML ids, so the adapter finds the target by
    position (adapters/libreoffice/). A wrong guess must not be scored as a LibreOffice
    preservation failure: the file is unscored for it when the lookup fails, when the
    pptx shape at the target position does not carry the target's name, or when the
    located paragraph's text differs from the input model's target paragraph
    (edits.comparable_text: object placeholders dropped, break characters unified).
    """
    if spec.kind != "soffice" or spec.format == "xlsx":
        return None
    if error_event is not None and error_event.get("type") in _LOCATOR_ERRORS:
        return {
            "error": f"{error_event.get('type')}: {error_event.get('message', '')}"[
                :500
            ]
        }
    locator = next((e for e in events if e.get("event") == "locator"), None)
    target = input_model.get("target_paragraph")
    if locator is None or target is None or "located_text" not in locator:
        return None
    info = {k: v for k, v in locator.items() if k not in ("event", "located_text")}
    if spec.format == "pptx" and locator.get("shape_method") != "index":
        return {
            "why": "the shape at the target position does not carry its name",
            "locator": info,
        }
    located = edits.comparable_text(locator["located_text"])
    expected = edits.comparable_text(target["text"] or "")
    if located == expected:
        return None
    return {
        "why": "the located paragraph's text differs from the target paragraph",
        "target_path": target["path"],
        "target_text": expected[:300],
        "located_text": located[:300],
        "locator": info,
    }


def _summary(
    prep: Prepared,
    entries: list[dict[str, Any]],
    results: list[dict[str, Any]],
    input_models: dict[str, dict[str, Any]],
    features: tuple[str, ...],
) -> dict[str, Any]:
    every = [r for r in results if r["adapter"] == prep.spec.id]
    rows = [r for r in every if r["outcome"] != "unscored"]
    mismatched = [r["fixture"] for r in every if r["reason"] == LOCATOR_MISMATCH]
    content_files = sum(
        1 for e in entries if input_models[e["path"]]["outcome"] == "success"
    ) - len(mismatched)
    base = {
        "adapter": prep.spec.id,
        "library": prep.spec.library,
        "files_total": len(entries),
        "no_edit_target": sum(
            1 for e in entries if input_models[e["path"]]["outcome"] == NO_EDIT_TARGET
        ),
        "locator_mismatch": len(mismatched),
        "locator_mismatch_files": mismatched,
        "content_files": content_files,
    }
    if not prep.available:
        status = "unsupported" if prep.spec.kind == "unsupported" else "unavailable"
        return {**base, "status": status, "reason": prep.reason}

    def check(row: dict[str, Any], name: str) -> dict[str, Any]:
        return next(c for c in row["checks"] if c["name"] == name)

    def passed(name: str) -> int:
        return sum(1 for r in rows if check(r, name)["outcome"] == "success")

    present = dict.fromkeys(features, 0)
    preserved = dict.fromkeys(features, 0)
    differing_files: dict[str, int] = {}
    fully = 0
    for row in rows:
        lane = check(row, "content-preserved")
        if not lane["scored"]:
            continue
        if row["outcome"] == "success":
            fully += 1
        differing = lane["detail"]["differing_features"]
        for feature in differing or {}:
            differing_files[feature] = differing_files.get(feature, 0) + 1
        for feature in lane["detail"]["input_features"]:
            present[feature] += 1
            # No output, or an output whose model could not be extracted, preserves nothing.
            if differing is not None and feature not in differing:
                preserved[feature] += 1
    content_preserved = passed("content-preserved")
    edit_applied = passed("edit-applied")
    opened = passed("opened")
    default_style_files = [
        r["fixture"]
        for r in rows
        if isinstance(check(r, "content-preserved")["detail"], dict)
        and check(r, "content-preserved")["detail"].get("new_cell_default_style")
    ]
    extra: dict[str, Any] = {}
    if prep.spec.format == "xlsx":
        extra["new_cell_default_style"] = len(default_style_files)
        extra["new_cell_default_style_files"] = default_style_files
    return {
        **base,
        "status": "measured",
        "files_attempted": len(rows),
        "opened": opened,
        "opened_label": f"{opened}/{len(rows)}",
        "saved": passed("saved"),
        "package_valid": passed("package-valid"),
        "edit_applied": edit_applied,
        "edit_applied_label": f"{edit_applied}/{content_files}",
        "content_preserved": content_preserved,
        "content_preserved_label": f"{content_preserved}/{content_files}",
        "fully_preserved": fully,
        "fully_preserved_label": f"{fully}/{content_files}",
        "feature_preservation": {
            feature: {
                "preserved": preserved[feature],
                "present": present[feature],
                "rate": round(preserved[feature] / present[feature], 4)
                if present[feature]
                else None,
            }
            for feature in features
        },
        "files_differing_by_feature": dict(
            sorted(differing_files.items(), key=lambda item: (-item[1], item[0]))
        ),
        **extra,
        "errors": sum(1 for r in rows if r["outcome"] == "error"),
        "timeouts": sum(1 for r in rows if r["outcome"] == "timeout"),
        "median_elapsed_ms": _median([r["samples"][0]["elapsed_ms"] for r in rows]),
        "diagnostics": {
            "scored": False,
            "render_similar": passed("render-similarity"),
            "render_unscored_input": sum(
                1
                for r in rows
                if check(r, "render-similarity")["outcome"] == "unscored"
            ),
            "parts_preserved": passed("parts-preserved"),
            "semantic_preserved": passed("semantic-preserved"),
            "byte_identical_outputs": sum(
                1 for r in rows if r["samples"][0]["output_identical_to_input"]
            ),
        },
    }


def _median(values: list[float]) -> float | None:
    if not values:
        return None
    ordered = sorted(values)
    mid = len(ordered) // 2
    return (
        ordered[mid]
        if len(ordered) % 2
        else round((ordered[mid - 1] + ordered[mid]) / 2, 3)
    )


def _roundtrip_one(
    root: Path,
    prep: Prepared,
    index: int,
    entry: dict[str, Any],
    work: Path,
    timeout: float,
    render_timeout: float,
    content_timeout: float,
    with_profile: Any,
    input_renders: dict[str, dict[str, Any]],
    input_models: dict[str, dict[str, Any]],
    redact: _Redactor,
) -> dict[str, Any]:
    spec = prep.spec
    case = work / spec.id / f"{index:03d}"
    in_dir, out_dir = case / "in", case / "out"
    in_dir.mkdir(parents=True)
    out_dir.mkdir()
    name = Path(entry["path"]).name
    source = in_dir / name
    shutil.copyfile(root / entry["path"], source)
    target = out_dir / f"{Path(name).stem}.{spec.format}"
    input_model = input_models[entry["path"]]
    edit_json: Path = input_model["edit_json"]

    if prep.pool is not None:
        child = prep.pool.run(source, target, timeout, edit=input_model["edit"])
    else:
        child = render.run_child(prep.command(root, source, target, edit_json), timeout)
    events = _events(child["stdout"])
    names = {e.get("event") for e in events}
    error_event = next((e for e in events if e.get("event") == "error"), None)
    if input_model["outcome"] == "success" and not child["timed_out"]:
        mismatch = _locator_mismatch(spec, events, error_event, input_model)
        if mismatch is not None:
            return _no_target_row(
                spec.id, entry, input_model, LOCATOR_MISMATCH, mismatch
            )
    output_ok = target.exists() and target.stat().st_size > 0 and not child["timed_out"]
    opened = "opened" in names
    saved = "saved" in names and output_ok

    checks = [
        _check(
            "opened",
            "success" if opened else "failure",
            _error_text(error_event, "open"),
        ),
        _check(
            "saved",
            "success" if saved else "failure",
            _error_text(error_event, "edit") or _error_text(error_event, "save"),
        ),
    ]
    sample: dict[str, Any] = {
        "iteration": 1,
        "phase": "measured",
        "elapsed_ms": child["elapsed_ms"],
        "output_sha256": sha256_file(target) if output_ok else None,
        "output_bytes": target.stat().st_size if output_ok else None,
        "output_identical_to_input": output_ok
        and sha256_file(target) == entry["sha256"],
        "details": {
            "command": [
                redact(arg)
                for arg in prep.command(
                    root, Path("<input>"), Path("<output>"), Path("<edit.json>")
                )
            ],
            "returncode": child["returncode"],
            "timed_out": child["timed_out"],
            "events": [e.get("event") for e in events],
            "error": _redacted_error(error_event, redact),
            "stderr_excerpt": redact(child["stderr"][-1500:]) or None,
        },
    }
    locator = next((e for e in events if e.get("event") == "locator"), None)
    if locator is not None:
        # How LibreOffice found the target (xlsx: sheet_method name, case-insensitive
        # name, or position).
        sample["details"]["locator"] = {
            k: v for k, v in locator.items() if k not in ("event", "located_text")
        }

    if saved:
        checks += _score_output(
            root / entry["path"],
            target,
            entry,
            case,
            render_timeout,
            with_profile,
            input_renders,
        )
        checks += _content_checks(
            spec.format, input_model, target, content_timeout, redact
        )
    else:
        why = "timed out" if child["timed_out"] else "no output to score"
        if error_event is not None and error_event.get("phase") == "edit":
            why = "the adapter failed to apply the edit: " + str(
                _error_text(error_event, "edit")
            )
        checks += [
            _check(n, "failure", why)
            for n in ("package-valid", "parts-preserved", "semantic-preserved")
        ]
        checks += _content_checks(spec.format, input_model, None, 0, redact, why)
        render_input = input_renders[entry["path"]]
        if render_input["outcome"] == "success":
            checks.append(_check("render-similarity", "failure", why))
        else:
            checks.append(
                _check(
                    "render-similarity", "unscored", "oracle could not render the input"
                )
            )
    checks.sort(key=lambda c: CHECKS.index(c["name"]))
    if child["timed_out"]:
        outcome, reason = "timeout", f"adapter exceeded {timeout:g}s"
    elif not (opened and saved):
        outcome = "error"
        if not opened:
            reason = "adapter failed to open"
        elif error_event is not None and error_event.get("phase") == "edit":
            reason = "adapter failed to apply the edit"
        else:
            reason = "adapter failed to save"
    elif all(c["outcome"] == "success" for c in checks if c["scored"]):
        outcome, reason = "success", None
    else:
        failed = [
            c["name"] for c in checks if c["scored"] and c["outcome"] != "success"
        ]
        outcome, reason = "failure", "failed: " + ", ".join(failed)
    sample.update(outcome=outcome, reason=reason)
    return {
        "adapter": spec.id,
        "lane": LANE,
        "fixture": entry["path"],
        "input_sha256": entry["sha256"],
        "source": entry["source"],
        "edit": input_model["edit"],
        "outcome": outcome,
        "reason": reason,
        "checks": checks,
        "samples": [sample],
    }


def _error_text(event: dict[str, Any] | None, phase: str) -> str | None:
    if event is None or event.get("phase") != phase:
        return None
    return f"{event.get('type')}: {event.get('message', '')[:300]}"


def _redacted_error(
    event: dict[str, Any] | None, redact: _Redactor
) -> dict[str, Any] | None:
    if event is None:
        return None
    return {
        "phase": event.get("phase"),
        "type": event.get("type"),
        "message": redact(str(event.get("message", "")))[:1000],
        "traceback_tail": redact(str(event.get("traceback", "")))[-1500:],
    }


def _content_checks(
    fmt: str,
    input_model: dict[str, Any],
    output_path: Path | None,
    timeout: float,
    redact: _Redactor,
    no_output: str | None = None,
) -> list[dict[str, Any]]:
    """edit-applied and content-preserved: the output's model against the expected model."""
    if input_model["outcome"] != "success":
        excluded = {"excluded": True, "reason": input_model["reason"]}
        return [
            _check(name, "unscored", excluded, scored=False)
            for name in ("edit-applied", "content-preserved")
        ]
    features = input_model["input_features"]

    def failed(reason: str | None, **extra: Any) -> list[dict[str, Any]]:
        return [
            _check("edit-applied", "failure", {"reason": reason, **extra}),
            _check(
                "content-preserved",
                "failure",
                {
                    "reason": reason,
                    "differing_features": None,
                    "input_features": features,
                    "sample": [],
                    **extra,
                },
            ),
        ]

    if output_path is None:
        return failed(no_output)
    event, elapsed = _content_child(
        ["compare", fmt, str(input_model["prepared_dir"]), str(output_path)], timeout
    )
    if event.get("event") != "compared":
        return failed(
            "output content model could not be extracted: "
            + redact(f"{event.get('type')}: {event.get('message', '')}")[:500],
            elapsed_ms=elapsed,
        )
    applied = event["edit_applied"]
    detail: dict[str, Any] = {
        "differing_features": event["differing_features"],
        "diff_count": event["diff_count"],
        "input_features": features,
        "sample": event["sample"],
        "elapsed_ms": elapsed,
    }
    if "new_cell_default_style" in event:
        detail["new_cell_default_style"] = event["new_cell_default_style"]
    return [
        _check("edit-applied", "success" if applied["applied"] else "failure", applied),
        _check(
            "content-preserved",
            "failure" if event["diff_count"] else "success",
            detail,
        ),
    ]


def _score_output(
    input_path: Path,
    output_path: Path,
    entry: dict[str, Any],
    case: Path,
    render_timeout: float,
    with_profile: Any,
    input_renders: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    checks: list[dict[str, Any]] = []
    inp = ooxml.read_package(input_path)
    try:
        out = ooxml.read_package(output_path)
    except ooxml.PackageError as exc:
        reason = f"output is not a readable package: {exc}"
        checks += [
            _check(n, "failure", reason)
            for n in ("package-valid", "parts-preserved", "semantic-preserved")
        ]
        out = None
    if out is not None:
        inherited = set(entry.get("input_package_problems", []))
        new_problems = [p for p in ooxml.package_problems(out) if p not in inherited]
        if ooxml.main_part(out) is None:
            new_problems.append("no main document part")
        checks.append(
            _check(
                "package-valid",
                "failure" if new_problems else "success",
                {
                    "new_problems": new_problems[:25],
                    "new_problem_count": len(new_problems),
                }
                if new_problems
                else None,
            )
        )
        mapping = ooxml.map_parts(inp, out)
        method_counts: dict[str, int] = {}
        for method in mapping.methods.values():
            method_counts[method] = method_counts.get(method, 0) + 1
        checks.append(
            _check(
                "parts-preserved",
                "failure" if mapping.missing else "success",
                {
                    "input_parts": len(inp.data),
                    "matched_by": method_counts,
                    "missing": mapping.missing[:25],
                    "missing_count": len(mapping.missing),
                    "added_count": len(mapping.added),
                },
            )
        )
        differing: list[str] = []
        compared = 0
        for key, out_key in sorted(mapping.mapping.items()):
            compared += 1
            if (
                not ooxml.is_rels_part(key)
                and key != ooxml.CT_PART
                and inp.content_type(key) != out.content_type(out_key)
            ):
                differing.append(f"{key} (content type)")
                continue
            if ooxml.canonical_part(inp, key, mapping.mapping) != ooxml.canonical_part(
                out, out_key, None
            ):
                differing.append(key)
        checks.append(
            _check(
                "semantic-preserved",
                "failure" if differing else "success",
                {
                    "compared_parts": compared,
                    "equal_parts": compared - len(differing),
                    "input_parts": len(inp.data),
                    "differing": differing[:25],
                    "differing_count": len(differing),
                },
            )
        )

    render_input = input_renders[entry["path"]]
    if render_input["outcome"] != "success":
        checks.append(
            _check(
                "render-similarity",
                "unscored",
                "oracle could not render the input",
                scored=False,
            )
        )
        return checks
    pdf, result = with_profile(
        lambda p: render.to_pdf(output_path, case / "pdf", p, render_timeout)
    )
    if pdf is None:
        reason = (
            "oracle render timed out"
            if result["timed_out"]
            else "oracle produced no PDF for the output"
        )
        checks.append(_check("render-similarity", "failure", reason))
        return checks
    comparison = render.compare_pdfs(render_input["pdf"], pdf)
    checks.append(
        _check(
            "render-similarity",
            "success" if comparison["pass"] else "failure",
            comparison,
        )
    )
    return checks
