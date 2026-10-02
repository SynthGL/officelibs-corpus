"""Holdout/dev split: find corpus SHA-256 digests already present in engine repositories.

A corpus file is `dev` when the same bytes exist anywhere in an engine repository's
working tree or anywhere in its git history (every reachable blob of every ref),
because the engine may have been developed or tuned against it. Everything else is
`holdout`. The scan is read-only: it walks files and runs `git rev-list` and
`git cat-file`, nothing else.
"""

from __future__ import annotations

import hashlib
import os
import subprocess
from dataclasses import dataclass, field
from pathlib import Path

ENGINE_REPOS = ("wolfxl", "wolfppt", "wolfdocx")


@dataclass
class EngineScan:
    name: str
    commit: str | None
    working_tree_files_hashed: int = 0
    history_blobs_hashed: int = 0
    matches: dict[str, list[str]] = field(default_factory=dict)  # sha256 -> locations


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def scan_engine(repo: Path, targets: dict[int, set[str]]) -> EngineScan:
    """Hash every working-tree file and history blob whose size equals a target size."""
    commit = None
    if (repo / ".git").exists():
        commit = subprocess.run(
            ["git", "-C", str(repo), "rev-parse", "HEAD"],
            capture_output=True,
            text=True,
            check=True,
        ).stdout.strip()
    scan = EngineScan(repo.name, commit)
    wanted = {sha for shas in targets.values() for sha in shas}

    for dirpath, dirnames, filenames in os.walk(repo):
        if ".git" in dirnames:
            dirnames.remove(".git")
        for filename in filenames:
            path = Path(dirpath) / filename
            try:
                if path.is_symlink() or not path.is_file():
                    continue
                size = path.stat().st_size
            except OSError:
                continue
            if size not in targets:
                continue
            sha = _hash_file(path)
            scan.working_tree_files_hashed += 1
            if sha in wanted:
                scan.matches.setdefault(sha, []).append(f"{repo.name}:working-tree")

    if commit is None:
        return scan
    objects = subprocess.run(
        ["git", "-C", str(repo), "rev-list", "--all", "--objects"],
        capture_output=True,
        check=True,
    ).stdout.decode("utf-8", "surrogateescape")
    paths: dict[str, str] = {}
    for line in objects.splitlines():
        oid, _, name = line.partition(" ")
        if name:
            paths.setdefault(oid, name)
    check = subprocess.run(
        [
            "git",
            "-C",
            str(repo),
            "cat-file",
            "--batch-check=%(objectname) %(objecttype) %(objectsize)",
        ],
        input="\n".join(paths).encode(),
        capture_output=True,
        check=True,
    ).stdout.decode()
    candidates = []
    for line in check.splitlines():
        parts = line.split()
        if len(parts) == 3 and parts[1] == "blob" and int(parts[2]) in targets:
            candidates.append(parts[0])
    if candidates:
        with subprocess.Popen(
            ["git", "-C", str(repo), "cat-file", "--batch"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
        ) as proc:
            assert proc.stdin is not None and proc.stdout is not None
            for oid in candidates:
                proc.stdin.write(oid.encode() + b"\n")
                proc.stdin.flush()
                header = proc.stdout.readline().split()
                size = int(header[2])
                body = proc.stdout.read(size)
                proc.stdout.read(1)
                sha = hashlib.sha256(body).hexdigest()
                scan.history_blobs_hashed += 1
                if sha in wanted:
                    scan.matches.setdefault(sha, []).append(f"{repo.name}:git-history")
            proc.stdin.close()
    return scan
