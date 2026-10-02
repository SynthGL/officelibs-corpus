"""officecorpus command line."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="officecorpus")
    sub = parser.add_subparsers(dest="command", required=True)

    fetch = sub.add_parser("fetch-sources", help="materialize pinned upstream sources")
    fetch.add_argument("--upstream", type=Path, default=REPO_ROOT / ".upstream")

    build = sub.add_parser(
        "build-corpus",
        help="select files, write corpus/, manifest.json, NOTICE, LICENSES/",
    )
    build.add_argument("--upstream", type=Path, default=REPO_ROOT / ".upstream")
    build.add_argument(
        "--engines-root",
        type=Path,
        required=True,
        help="directory containing the wolfxl, wolfppt, and wolfdocx repositories (scanned read-only)",
    )

    roundtrip = sub.add_parser(
        "roundtrip",
        help="run the one-edit round-trip lane (one run.json for all given adapters)",
    )
    roundtrip.add_argument("--format", required=True, choices=("xlsx", "pptx", "docx"))
    roundtrip.add_argument(
        "--adapter",
        required=True,
        action="append",
        help="adapter id; repeat for several adapters, or 'all' for every adapter of the format",
    )
    roundtrip.add_argument(
        "--split", default="holdout", choices=("holdout", "dev", "all")
    )
    roundtrip.add_argument(
        "--output", type=Path, required=True, help="new or empty directory"
    )
    roundtrip.add_argument(
        "--timeout", type=float, default=120.0, help="per-file adapter timeout, seconds"
    )
    roundtrip.add_argument(
        "--render-timeout",
        type=float,
        default=180.0,
        help="per-file PDF conversion timeout",
    )
    roundtrip.add_argument(
        "--content-timeout",
        type=float,
        default=300.0,
        help="per-package content-model extraction timeout, seconds",
    )
    roundtrip.add_argument("--jobs", type=int, default=4, help="concurrent files")
    roundtrip.add_argument(
        "--limit", type=int, default=None, help="first N files only (smoke runs)"
    )

    args = parser.parse_args(argv)
    if args.command == "fetch-sources":
        from .corpus import fetch_sources

        fetch_sources(REPO_ROOT, args.upstream)
    elif args.command == "build-corpus":
        from .corpus import build_corpus

        manifest = build_corpus(REPO_ROOT, args.upstream, args.engines_root)
        json.dump(
            {"counts": manifest["counts"], "holdout": manifest["holdout"]},
            sys.stdout,
            indent=2,
        )
        print()
    elif args.command == "roundtrip":
        from .roundtrip import run_roundtrip

        path = run_roundtrip(
            REPO_ROOT,
            fmt=args.format,
            adapter_ids=args.adapter,
            split=args.split,
            output=args.output,
            timeout=args.timeout,
            render_timeout=args.render_timeout,
            content_timeout=args.content_timeout,
            jobs=args.jobs,
            limit=args.limit,
        )
        print(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
