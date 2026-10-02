"""Fetch pinned upstream sources and build the corpus, manifest, split, and notices."""

from __future__ import annotations

import hashlib
import json
import re
import shutil
import subprocess
import tarfile
import urllib.request
import zipfile
from pathlib import Path
from typing import Any

from . import ooxml
from .split import ENGINE_REPOS, scan_engine

FORMATS = ("xlsx", "pptx", "docx")
MACRO_EXTENSIONS = {
    "xlsx": ("xlsm", "xltm"),
    "pptx": ("pptm", "potm", "ppsm"),
    "docx": ("docm", "dotm"),
}
MANIFEST_SCHEMA = "officecorpus.manifest/v1"
# Upstream file names that announce a deliberately damaged or hostile input.
INTENTIONALLY_BROKEN = re.compile(
    r"corrupt|invalid|broken|truncat|fuzz|crash|ofz\d|malformed|damaged|bogus|garbage|poc[-_.]",
    re.IGNORECASE,
)
LICENSE_TEXT_NAMES = {
    ("poi", "legal/LICENSE"): "Apache-2.0.txt",
    ("poi", "legal/NOTICE"): "NOTICE-apache-poi.txt",
    ("libreoffice", "COPYING.MPL"): "MPL-2.0.txt",
    ("openpyxl", "LICENCE.rst"): "MIT-openpyxl.txt",
    ("python-pptx", "LICENSE"): "MIT-python-pptx.txt",
    ("python-docx", "LICENSE"): "MIT-python-docx.txt",
}
GOVDOCS_LICENSE_NAME = "LicenseRef-Govdocs1-Public.txt"
GOVDOCS_LICENSE_TEXT = """Govdocs1 redistribution statement (LicenseRef-Govdocs1-Public)

Source: https://digitalcorpora.org/corpora/file-corpora/files/

Digital Corpora describes Govdocs1 as "a corpus of 1 million documents that are
freely available for research and may be (to the best of our knowledge) freely
redistributed." The documents were collected from publicly accessible web
servers in the .gov domain. Works of the U.S. federal government are not subject
to copyright in the United States (17 U.S.C. 105); documents from state or local
government servers may carry other terms, and Digital Corpora's statement above
is the basis on which this corpus redistributes them.

Citation: Simson Garfinkel, Paul Farrell, Vassil Roussev, and George Dinolt,
"Bringing Science to Digital Forensics with Standardized Forensic Corpora",
DFRWS 2009, Montreal, Canada.
"""


def load_sources(root: Path) -> dict[str, Any]:
    return json.loads((root / "sources.json").read_text(encoding="utf-8"))


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


# ---------------------------------------------------------------------------
# Fetch
# ---------------------------------------------------------------------------


def fetch_sources(root: Path, upstream: Path) -> None:
    """Materialize every pinned source under `upstream/<source-id>`."""
    config = load_sources(root)
    upstream.mkdir(parents=True, exist_ok=True)
    for source in config["sources"]:
        target = upstream / source["id"]
        if source["kind"] == "git":
            _fetch_git(source, target)
        elif source["id"] == "govdocs1":
            for fmt, archive in source["archives"].items():
                zip_path = target / f"{fmt}.zip"
                _download(archive["url"], zip_path, archive.get("sha256"))
                with zipfile.ZipFile(zip_path) as handle:
                    handle.extractall(target / fmt)
        else:
            archive = source["archives"]["xlsx"]
            tar_path = upstream / f"{source['id']}.tar.gz"
            _download(archive["url"], tar_path, archive.get("sha256"))
            target.mkdir(parents=True, exist_ok=True)
            with tarfile.open(tar_path) as handle:
                for member in handle.getmembers():
                    parts = member.name.split("/")[archive.get("strip_components", 0) :]
                    if not parts or not parts[0]:
                        continue
                    member.name = "/".join(parts)
                    handle.extract(member, target, filter="data")


def _download(url: str, dest: Path, expected_sha256: str | None) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    if not dest.exists():
        with urllib.request.urlopen(url) as response, dest.open("wb") as handle:
            shutil.copyfileobj(response, handle)
    if expected_sha256 and sha256_file(dest) != expected_sha256:
        raise RuntimeError(f"{dest.name}: SHA-256 mismatch against sources.json")


def _fetch_git(source: dict[str, Any], target: Path) -> None:
    def git(*args: str) -> None:
        subprocess.run(["git", "-C", str(target), *args], check=True)

    if not (target / ".git").exists():
        target.mkdir(parents=True, exist_ok=True)
        git("init", "-q")
        git("remote", "add", "origin", source["repo"] + ".git")
        git("config", "remote.origin.promisor", "true")
        git("config", "remote.origin.partialclonefilter", "blob:none")
    if "sparse_patterns" in source:
        git("sparse-checkout", "set", "--no-cone", *source["sparse_patterns"])
    else:
        git("sparse-checkout", "set", *source["sparse"])
    git("fetch", "-q", "--depth", "1", "--filter=blob:none", "origin", source["commit"])
    git("checkout", "-q", "--detach", source["commit"])


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------


def _candidates(
    source: dict[str, Any], fmt: str, upstream: Path
) -> list[tuple[Path, str, str]]:
    """(file, upstream path, corpus-relative stem) for every file with a relevant extension."""
    extensions = (fmt, *MACRO_EXTENSIONS[fmt])
    out: list[tuple[Path, str, str]] = []
    base = upstream / source["id"]
    if source["id"] == "govdocs1":
        roots = [(base / fmt, "")]
    else:
        roots = [(base / rel, rel) for rel in source.get("paths", {}).get(fmt, [])]
    for root, rel in roots:
        if not root.is_dir():
            continue
        for path in sorted(root.rglob("*")):
            if not path.is_file() or path.suffix.lower().lstrip(".") not in extensions:
                continue
            inner = path.relative_to(root).as_posix()
            upstream_path = f"{rel}/{inner}" if rel else inner
            out.append((path, upstream_path, inner))
    return out


def _flatten(inner: str, taken: set[str]) -> str:
    """Corpus file name: basename, widened with parent directories only on collision."""
    parts = inner.split("/")
    for width in range(1, len(parts) + 1):
        name = "__".join(parts[-width:])
        if name.lower() not in taken:
            taken.add(name.lower())
            return name
    raise RuntimeError(f"cannot name {inner}")


def _greedy(
    pool: list[dict[str, Any]], quota: int, coverage: dict[str, int]
) -> list[dict[str, Any]]:
    """Pick files that add the most not-yet-covered features (rarity weighted)."""
    chosen: list[dict[str, Any]] = []
    remaining = list(pool)
    while remaining and len(chosen) < quota:

        def score(entry: dict[str, Any]) -> tuple[float, int, int, str]:
            gain = sum(1.0 / (1 + coverage.get(f, 0)) for f in entry["features"])
            return (
                -gain,
                -len(entry["features"]),
                entry["size"],
                entry["upstream_path"],
            )

        remaining.sort(key=score)
        pick = remaining.pop(0)
        chosen.append(pick)
        for feature in pick["features"]:
            coverage[feature] = coverage.get(feature, 0) + 1
    return chosen


def build_corpus(root: Path, upstream: Path, engines_root: Path) -> dict[str, Any]:
    config = load_sources(root)
    selection = config["selection"]
    sources = {s["id"]: s for s in config["sources"]}
    order = selection["source_order"]
    corpus_dir = root / "corpus"
    if corpus_dir.exists():
        shutil.rmtree(corpus_dir)
    excluded: list[dict[str, Any]] = []
    files: list[dict[str, Any]] = []
    candidate_counts: dict[str, dict[str, int]] = {}

    for fmt in FORMATS:
        seen_sha: dict[str, str] = {}
        eligible_by_source: dict[str, list[dict[str, Any]]] = {}
        quotas = selection["quotas"][fmt]
        candidate_counts[fmt] = {}
        for source_id in order:
            if source_id not in quotas:
                continue
            source = sources[source_id]
            candidates = _candidates(source, fmt, upstream)
            candidate_counts[fmt][source_id] = len(candidates)
            eligible: list[dict[str, Any]] = []
            for path, upstream_path, inner in candidates:
                record = {
                    "format": fmt,
                    "source": source_id,
                    "upstream_path": upstream_path,
                }
                size = path.stat().st_size
                sha = sha256_file(path)
                record["sha256"] = sha
                reason = None
                if path.suffix.lower().lstrip(".") != fmt:
                    reason = f"macro-enabled: {path.suffix.lower()} extension"
                elif INTENTIONALLY_BROKEN.search(inner):
                    reason = "intentionally-malformed: upstream file name marks it as a damaged or hostile input"
                elif size > selection["max_file_bytes"]:
                    reason = (
                        f"too-large: {size} bytes exceeds {selection['max_file_bytes']}"
                    )
                if reason is None:
                    reason, pkg = ooxml.screen(path)
                    if reason is None and pkg is not None:
                        bad = _malformed_xml_part(pkg)
                        if bad:
                            reason = f"corrupt: part {bad} is not well-formed XML"
                if reason is None and sha in seen_sha:
                    reason = f"duplicate: identical bytes to {seen_sha[sha]}"
                if reason is not None:
                    excluded.append({**record, "reason": reason})
                    continue
                assert pkg is not None
                seen_sha[sha] = f"{source_id}:{upstream_path}"
                eligible.append(
                    {
                        **record,
                        "file": path,
                        "inner": inner,
                        "size": size,
                        "features": ooxml.detect_features(pkg, fmt),
                        "inventory": ooxml.part_inventory(pkg),
                        "input_package_problems": ooxml.package_problems(pkg),
                    }
                )
            eligible_by_source[source_id] = eligible

        coverage: dict[str, int] = {}
        target = selection["per_format_target"]
        picked: list[dict[str, Any]] = []
        shortfall = 0
        for source_id in order:
            if source_id in quotas:
                chosen = _greedy(
                    eligible_by_source[source_id],
                    quotas[source_id] + shortfall,
                    coverage,
                )
                shortfall = quotas[source_id] + shortfall - len(chosen)
                picked.extend(chosen)
        if len(picked) < target:  # refill from any source in order
            chosen_ids = {id(e) for e in picked}
            rest = [
                e
                for s in order
                if s in quotas
                for e in eligible_by_source[s]
                if id(e) not in chosen_ids
            ]
            picked.extend(_greedy(rest, target - len(picked), coverage))

        taken: dict[str, set[str]] = {}
        for entry in picked:
            name = _flatten(entry["inner"], taken.setdefault(entry["source"], set()))
            rel = Path("corpus") / fmt / entry["source"] / name
            (root / rel).parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(entry["file"], root / rel)
            source = sources[entry["source"]]
            upstream_ref: dict[str, Any] = {"path": entry["upstream_path"]}
            if source["kind"] == "git" or source["id"] == "openpyxl":
                upstream_ref.update(repo=source["repo"], commit=source["commit"])
            if source["id"] == "govdocs1":
                upstream_ref.update(
                    archive=source["archives"][fmt]["url"],
                    archive_sha256=source["archives"][fmt]["sha256"],
                )
            files.append(
                {
                    "path": rel.as_posix(),
                    "format": fmt,
                    "sha256": entry["sha256"],
                    "size": entry["size"],
                    "source": entry["source"],
                    "upstream": upstream_ref,
                    "license": source["license"],
                    "features": entry["features"],
                    "parts": entry["inventory"]["part_count"],
                    "input_package_problems": entry["input_package_problems"],
                }
            )

    files.sort(key=lambda f: f["path"])
    split_info = _assign_split(files, engines_root)
    excluded.sort(key=lambda e: (e["format"], e["source"], e["upstream_path"]))
    manifest = _manifest(config, files, excluded, candidate_counts, split_info)
    _write_json(root / "manifest.json", manifest)
    holdout = [f for f in files if f["split"] == "holdout"]
    (root / "holdout.txt").write_text(
        "".join(f"{f['sha256']}  {f['path']}\n" for f in holdout), encoding="utf-8"
    )
    write_notices(root, upstream, config)
    return manifest


def _malformed_xml_part(pkg: ooxml.Package) -> str | None:
    from lxml import etree

    for key in sorted(pkg.data):
        if ooxml.is_xml_part(pkg, key):
            try:
                etree.fromstring(pkg.data[key], ooxml._XML_PARSER)
            except etree.XMLSyntaxError:
                return key
    return None


def _assign_split(files: list[dict[str, Any]], engines_root: Path) -> dict[str, Any]:
    targets: dict[int, set[str]] = {}
    for entry in files:
        targets.setdefault(entry["size"], set()).add(entry["sha256"])
    scans = []
    exposure: dict[str, list[str]] = {}
    for name in ENGINE_REPOS:
        repo = engines_root / name
        if not repo.is_dir():
            raise RuntimeError(
                f"engine repository {name} not found under the engines root"
            )
        scan = scan_engine(repo, targets)
        scans.append(
            {
                "repo": name,
                "commit": scan.commit,
                "working_tree_files_hashed": scan.working_tree_files_hashed,
                "history_blobs_hashed": scan.history_blobs_hashed,
                "matched_corpus_files": len(scan.matches),
            }
        )
        for sha, locations in scan.matches.items():
            exposure.setdefault(sha, []).extend(locations)
    for entry in files:
        locations = sorted(set(exposure.get(entry["sha256"], [])))
        entry["split"] = "dev" if locations else "holdout"
        if locations:
            entry["engine_exposure"] = locations
    return {"engines": scans}


def holdout_digest(files: list[dict[str, Any]]) -> str:
    lines = sorted(
        f"{f['sha256']}  {f['path']}" for f in files if f["split"] == "holdout"
    )
    return hashlib.sha256("\n".join(lines).encode()).hexdigest()


def files_digest(files: list[dict[str, Any]]) -> str:
    return hashlib.sha256(
        json.dumps(files, sort_keys=True, separators=(",", ":")).encode()
    ).hexdigest()


def _manifest(
    config: dict[str, Any],
    files: list[dict[str, Any]],
    excluded: list[dict[str, Any]],
    candidate_counts: dict[str, dict[str, int]],
    split_info: dict[str, Any],
) -> dict[str, Any]:
    counts: dict[str, Any] = {}
    for entry in files:
        fmt_counts = counts.setdefault(
            entry["format"], {"total": 0, "by_source": {}, "by_split": {}}
        )
        fmt_counts["total"] += 1
        by_source = fmt_counts["by_source"].setdefault(
            entry["source"], {"holdout": 0, "dev": 0}
        )
        by_source[entry["split"]] += 1
        fmt_counts["by_split"][entry["split"]] = (
            fmt_counts["by_split"].get(entry["split"], 0) + 1
        )
    return {
        "schema": MANIFEST_SCHEMA,
        "sources": [
            {
                k: v
                for k, v in s.items()
                if k
                in ("id", "title", "repo", "commit", "archives", "license", "homepage")
            }
            for s in config["sources"]
        ],
        "selection": {
            **config["selection"],
            "candidates_found": candidate_counts,
            "policy": (
                "Per format and source: screen out encrypted, macro-enabled, intentionally malformed, corrupt "
                "(unreadable zip, missing content types or main part, any XML part not well-formed), oversized, "
                "and duplicate files; then greedily pick the file whose detected features add the most "
                "rarity-weighted coverage (1/(1+times already covered) per feature), ties broken by feature count, "
                "smaller size, then upstream path. Unfilled quota rolls over to the next source."
            ),
        },
        "split": {
            "rule": (
                "dev if the file's SHA-256 matches any working-tree file or any blob reachable from any ref in "
                "the git history of the wolfxl, wolfppt, or wolfdocx engine repositories; holdout otherwise."
            ),
            **split_info,
        },
        "counts": counts,
        "files_digest": files_digest(files),
        "holdout": {
            "count": sum(1 for f in files if f["split"] == "holdout"),
            "digest": holdout_digest(files),
            "digest_rule": "sha256 of newline-joined sorted '<sha256>  <path>' lines of holdout files (holdout.txt)",
        },
        "files": files,
        "excluded": excluded,
    }


def write_notices(root: Path, upstream: Path, config: dict[str, Any]) -> None:
    licenses = root / "LICENSES"
    if licenses.exists():
        shutil.rmtree(licenses)
    licenses.mkdir()
    for (source_id, rel), name in LICENSE_TEXT_NAMES.items():
        shutil.copyfile(upstream / source_id / rel, licenses / name)
    (licenses / GOVDOCS_LICENSE_NAME).write_text(GOVDOCS_LICENSE_TEXT, encoding="utf-8")
    lines = [
        "officelibs-corpus",
        "",
        "This repository redistributes Office Open XML test and real-world documents",
        "collected from the upstream sources below. Each document keeps the license of",
        "its source; manifest.json records the source, pinned revision, upstream path,",
        "and SPDX license identifier of every file. License texts are in LICENSES/.",
        "",
    ]
    for source in config["sources"]:
        pin = source.get("commit") or ", ".join(
            f"{fmt}: sha256 {a['sha256']}"
            for fmt, a in source.get("archives", {}).items()
        )
        location = source.get("repo") or source.get("homepage")
        texts = [
            name for (sid, _), name in LICENSE_TEXT_NAMES.items() if sid == source["id"]
        ]
        if source["id"] == "govdocs1":
            texts = [GOVDOCS_LICENSE_NAME]
        lines += [
            f"{source['title']} ({source['id']})",
            f"  Location: {location}",
            f"  Pinned:   {pin}",
            f"  License:  {source['license']} (LICENSES/{', LICENSES/'.join(texts)})",
            f"  Corpus:   corpus/<format>/{source['id']}/",
            f"  {source['attribution']}",
            "",
        ]
    lines += [
        "Apache POI NOTICE (reproduced as required by Apache-2.0 section 4(d)) is in",
        "LICENSES/NOTICE-apache-poi.txt.",
        "",
        "LibreOffice test documents are covered by the MPL-2.0. Their Source Code Form",
        "is the file itself; the upstream location of each file is recorded in",
        "manifest.json.",
        "",
    ]
    (root / "NOTICE").write_text("\n".join(lines), encoding="utf-8")


def _write_json(path: Path, data: Any) -> None:
    path.write_text(
        json.dumps(data, indent=2, sort_keys=False) + "\n", encoding="utf-8"
    )


def load_manifest(root: Path) -> dict[str, Any]:
    return json.loads((root / "manifest.json").read_text(encoding="utf-8"))


def verify_holdout(root: Path, manifest: dict[str, Any]) -> None:
    """Refuse to run against a corpus whose holdout list or bytes drifted from the frozen digest."""
    if holdout_digest(manifest["files"]) != manifest["holdout"]["digest"]:
        raise RuntimeError(
            "manifest holdout digest mismatch: the holdout list changed after freezing"
        )
    for entry in manifest["files"]:
        if sha256_file(root / entry["path"]) != entry["sha256"]:
            raise RuntimeError(f"{entry['path']}: bytes do not match manifest SHA-256")
