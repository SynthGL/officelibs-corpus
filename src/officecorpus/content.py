"""Content lane: extract content models (`officecorpus.model.<format>`) and compare them.

The runner calls this module in a child process per input and per output, so a model
that is slow or crashes on one package cannot stall or take down the run. Protocol:
one JSON line on stdout.

  python -m officecorpus.content prepare <format> <package> <dir>
      Compute the input's edit spec (officecorpus.edits.target) and write it to
      <dir>/edit.json. When there is one, extract the input model (<dir>/model.json)
      and the model the edit implies (<dir>/expected.json). Prints
      {"event": "prepared", "edit": {...} | null, "input_features": [...]} or, when
      the input model cannot be extracted, "model_error" instead of "input_features".
  python -m officecorpus.content compare <format> <dir> <package>
      Extract the output model, compare it with the expected model, check the edit,
      and print {"event": "compared", "differing_features": {...}, "diff_count": N,
      "sample": [...], "edit_applied": {...}}; xlsx adds the unscored diagnostic
      "new_cell_default_style" (edits.new_cell_default_style).

Either command prints {"event": "error", "type": ..., "message": ...} when it raises
(`ValueError` is the models' documented signal for an unreadable package).

Every model passes through JSON before `compare`, so stored and freshly extracted
models have identical value types.

xlsx formula cached values: a difference under `/cells/<sheet>/<cell>` whose INPUT
cell holds a formula (the input model's `formulas` feature) is attributed to the
feature `formula_cached_values` instead of `cells`. The model is untouched; only the
attribution changes, so the cached result of a formula is scored separately from
typed data.
"""

from __future__ import annotations

import importlib
import json
import sys
import traceback
from pathlib import Path
from types import ModuleType
from typing import Any

FORMULA_CACHE_FEATURE = "formula_cached_values"
SAMPLE_DIFFS = 10
SAMPLE_VALUE_CHARS = 200
# Replaces every non-empty scalar when probing which features an input contains.
_PROBE = "\x00officecorpus-probe\x00"


def model_module(fmt: str) -> ModuleType:
    if fmt not in ("xlsx", "pptx", "docx"):
        raise ValueError(f"no content model for {fmt!r}")
    return importlib.import_module(f"officecorpus.model.{fmt}")


def features(fmt: str) -> tuple[str, ...]:
    """Scored features: the model's FEATURES plus the xlsx formula-cache relabel."""
    base = model_module(fmt).FEATURES
    if fmt != "xlsx":
        return base
    at = base.index("formulas") + 1
    return (*base[:at], FORMULA_CACHE_FEATURE, *base[at:])


def _unescape(segment: str) -> str:
    return segment.replace("~1", "/").replace("~0", "~")


def relabel(
    fmt: str, input_model: dict[str, Any], diffs: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    """Attribute xlsx cell differences at input formula cells to formula_cached_values."""
    if fmt != "xlsx":
        return diffs
    formulas = input_model.get("formulas") or {}
    for diff in diffs:
        parts = diff["path"].split("/")
        if (
            len(parts) >= 4
            and parts[1] == "cells"
            and _unescape(parts[3]) in formulas.get(_unescape(parts[2]), {})
        ):
            diff["feature"] = FORMULA_CACHE_FEATURE
    return diffs


def _json_roundtrip(value: Any) -> Any:
    return json.loads(json.dumps(value, ensure_ascii=False))


def _hollow(value: Any) -> Any:
    if isinstance(value, dict):
        return {key: _hollow(item) for key, item in value.items()}
    if isinstance(value, list):
        return [*(_hollow(item) for item in value), _PROBE] if value else value
    if value is None or value is False or value == "":
        return value
    return _PROBE


def present_features(fmt: str, model: dict[str, Any]) -> list[str]:
    """Features the model attributes at least one non-empty value to.

    Every scalar that is not null, false, or the empty string is replaced by a probe
    value and every non-empty list gains one probe item; the model's own `compare`
    then attributes each resulting difference to a feature, so presence uses exactly
    the feature attribution that scoring uses (a list contributes its own feature,
    e.g. a deck's slide list, as well as the features of its items).
    """
    diffs = model_module(fmt).compare(model, _hollow(model))
    found = {diff["feature"] for diff in relabel(fmt, model, diffs)}
    return [feature for feature in features(fmt) if feature in found]


def _clip(value: Any) -> Any:
    if isinstance(value, str):
        text = value
    elif isinstance(value, (dict, list)):
        text = json.dumps(value, ensure_ascii=False, sort_keys=True)
    else:
        return value
    if len(text) <= SAMPLE_VALUE_CHARS:
        return value
    return text[: SAMPLE_VALUE_CHARS - 1] + "\u2026"


def summarize(fmt: str, diffs: list[dict[str, Any]]) -> dict[str, Any]:
    counts: dict[str, int] = {}
    for diff in diffs:
        counts[diff["feature"]] = counts.get(diff["feature"], 0) + 1
    order = {feature: index for index, feature in enumerate(features(fmt))}
    return {
        "differing_features": dict(
            sorted(
                counts.items(),
                key=lambda item: (order.get(item[0], len(order)), item[0]),
            )
        ),
        "diff_count": len(diffs),
        "sample": [
            {
                "feature": diff["feature"],
                "path": diff["path"],
                "kind": diff["kind"],
                "before": _clip(diff["before"]),
                "after": _clip(diff["after"]),
            }
            for diff in diffs[:SAMPLE_DIFFS]
        ],
    }


def _emit(**payload: Any) -> None:
    sys.stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def _read(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _write(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")


def _prepare(fmt: str, package: Path, out: Path) -> None:
    from . import edits

    edit = edits.target(fmt, package)
    _write(out / "edit.json", edit)
    if edit is None:
        _emit(event="prepared", edit=None)
        return
    try:
        model = _json_roundtrip(model_module(fmt).extract(package))
        expected = edits.expected(fmt, package, model, edit)
        paragraph = edits.target_paragraph(fmt, model, expected, edit)
    except Exception as exc:  # noqa: BLE001 - recorded as the exclusion reason
        _emit(
            event="prepared",
            edit=edit,
            model_error={"type": type(exc).__name__, "message": str(exc)[:1000]},
        )
        return
    _write(out / "model.json", model)
    _write(out / "expected.json", expected)
    _emit(
        event="prepared",
        edit=edit,
        input_features=present_features(fmt, model),
        target_paragraph=paragraph,
    )


def _compare(fmt: str, prepared: Path, package: Path) -> None:
    from . import edits

    edit = _read(prepared / "edit.json")
    model = _read(prepared / "model.json")
    expected = _read(prepared / "expected.json")
    output = _json_roundtrip(model_module(fmt).extract(package))
    diffs = relabel(fmt, model, model_module(fmt).compare(expected, output))
    extra: dict[str, Any] = {}
    if fmt == "xlsx":
        extra["new_cell_default_style"] = edits.new_cell_default_style(
            fmt, expected, output, edit
        )
    _emit(
        event="compared",
        **summarize(fmt, diffs),
        edit_applied=edits.applied(fmt, expected, output, edit),
        **extra,
    )


def main(argv: list[str]) -> int:
    command, fmt, first, second = argv
    model_module(fmt)
    try:
        if command == "prepare":
            _prepare(fmt, Path(first), Path(second))
        elif command == "compare":
            _compare(fmt, Path(first), Path(second))
        else:
            raise SystemExit(f"unknown command {command!r}")
    except Exception as exc:  # noqa: BLE001 - any model failure is reported, not raised
        _emit(
            event="error",
            type=type(exc).__name__,
            message=str(exc)[:1000],
            traceback_tail=traceback.format_exc()[-1500:],
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
