"""Containerized adapters (JVM, .NET, Go): image provisioning and persistent helper sessions.

Each helper under `adapters/<runtime>/<library>/` is a small program that serves many
round trips from one process: every stdin line is `<ext>\\t<base64 input>\\t<base64 edit JSON>`.
The helper opens the input, applies the one edit spec through the library's public API, saves,
and answers with JSON lines (`opened`, `saved` or `error`, then `done` carrying the in-process
open plus edit plus save time and the base64 output). One container therefore handles many
files, which keeps a remote Docker host usable; a file that exceeds the timeout kills its
container, and the next file starts a fresh one.

The Docker context comes from `OFFICECORPUS_DOCKER_CONTEXT` (default: the current context).
Images are tagged `officecorpus/<library>:<version>` and labelled with the sha256 of the helper
source; a missing image, or one built from different source, is rebuilt before the run.
"""

from __future__ import annotations

import base64
import hashlib
import json
import os
import queue
import subprocess
import threading
import time
import uuid
from pathlib import Path
from typing import Any

SOURCE_LABEL = "org.officecorpus.helper-sha256"
BUILD_TIMEOUT = 3600.0
CALL_TIMEOUT = 600.0
RUN_FLAGS = ("--rm", "-i", "--network", "none")


def context() -> str | None:
    return os.environ.get("OFFICECORPUS_DOCKER_CONTEXT") or None


def docker(*args: str) -> list[str]:
    ctx = context()
    return ["docker", *(["--context", ctx] if ctx else []), *args]


def _call(args: list[str], timeout: float = CALL_TIMEOUT) -> str:
    return subprocess.run(
        args, capture_output=True, text=True, check=True, timeout=timeout
    ).stdout


def helper_sha256(helper: Path) -> str:
    """Digest of every helper source file (path and content), in sorted order."""
    digest = hashlib.sha256()
    for path in sorted(p for p in helper.rglob("*") if p.is_file()):
        digest.update(path.relative_to(helper).as_posix().encode() + b"\0")
        digest.update(hashlib.sha256(path.read_bytes()).hexdigest().encode() + b"\n")
    return digest.hexdigest()


def _inspect(image: str) -> dict[str, Any] | None:
    try:
        out = _call(docker("image", "inspect", image))
    except subprocess.CalledProcessError:
        return None
    return json.loads(out)[0]


def provision(root: Path, helper_dir: str, image: str) -> tuple[str, dict[str, Any]]:
    """Ensure the image exists and matches the helper source; return (library version, receipt)."""
    helper = root / helper_dir
    source_sha = helper_sha256(helper)
    info = _inspect(image)
    rebuilt = False
    if (
        info is None
        or (info["Config"].get("Labels") or {}).get(SOURCE_LABEL) != source_sha
    ):
        _call(
            docker(
                "build", "-q", "-t", image, "--label", f"{SOURCE_LABEL}={source_sha}"
            )
            + [str(helper)],
            BUILD_TIMEOUT,
        )
        info = _inspect(image)
        rebuilt = True
    assert info is not None
    out = _call(docker("run", *RUN_FLAGS, image, "--version"))
    version = json.loads(out.strip().splitlines()[-1])["version"]
    server = json.loads(
        _call(
            docker(
                "info",
                "--format",
                '{"os":{{json .OSType}},"arch":{{json .Architecture}},'
                '"kernel":{{json .KernelVersion}},"server_version":{{json .ServerVersion}},'
                '"cpus":{{json .NCPU}},"memory_bytes":{{json .MemTotal}}}',
            )
        )
    )
    labels = info["Config"].get("Labels") or {}
    receipt = {
        "image": image,
        "image_id": info["Id"],
        "image_platform": f"{info.get('Os')}/{info.get('Architecture')}",
        "image_labels": {
            k: v for k, v in sorted(labels.items()) if "officecorpus" in k
        },
        "image_rebuilt_for_run": rebuilt,
        "helper_source": helper_dir,
        "helper_source_sha256": source_sha,
        "docker_context": context() or "current",
        "docker_server": server,
        "run_flags": list(RUN_FLAGS),
        "execution": (
            "one long-lived helper process per container serves many files over stdin/stdout; "
            "each file gets a private input copy; a timeout or crash kills the container and the "
            "next file starts a fresh one"
        ),
        "timing": (
            "elapsed_ms is the helper's in-process open, edit, and save time on the Docker host; it excludes "
            "runtime startup and transfer, so it is not comparable with cold-subprocess adapters"
        ),
    }
    return version, receipt


class _Session:
    def __init__(self, image: str) -> None:
        self.name = f"officecorpus-{uuid.uuid4().hex[:12]}"
        self.argv = docker("run", *RUN_FLAGS, "--name", self.name, image)
        self.proc = subprocess.Popen(
            self.argv,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            start_new_session=True,
        )
        self.lines: queue.Queue[bytes | None] = queue.Queue()
        self.stderr = bytearray()
        self.stderr_lock = threading.Lock()
        threading.Thread(target=self._pump_stdout, daemon=True).start()
        threading.Thread(target=self._pump_stderr, daemon=True).start()

    def _pump_stdout(self) -> None:
        assert self.proc.stdout is not None
        for line in self.proc.stdout:
            self.lines.put(line)
        self.lines.put(None)

    def _pump_stderr(self) -> None:
        assert self.proc.stderr is not None
        for line in self.proc.stderr:
            with self.stderr_lock:
                self.stderr += line
                del self.stderr[:-65536]

    def _stderr_tail(self) -> str:
        with self.stderr_lock:
            text = self.stderr.decode("utf-8", "replace")
            self.stderr.clear()
        return text

    def request(
        self, source: Path, target: Path, timeout: float, edit: dict[str, Any]
    ) -> tuple[dict[str, Any], bool]:
        """Run one round trip; return (child-style result, whether the session is reusable)."""
        self._stderr_tail()
        payload = (
            source.suffix.lstrip(".").encode()
            + b"\t"
            + base64.b64encode(source.read_bytes())
            + b"\t"
            + base64.b64encode(json.dumps(edit).encode())
            + b"\n"
        )
        start = time.perf_counter()
        events: list[dict[str, Any]] = []
        done: dict[str, Any] | None = None
        timed_out = False
        try:
            assert self.proc.stdin is not None
            self.proc.stdin.write(payload)
            self.proc.stdin.flush()
        except (BrokenPipeError, OSError):
            pass
        else:
            # The helper consumes the request as it arrives, so the clock starts once it is sent.
            deadline = time.perf_counter() + timeout
            while True:
                remaining = deadline - time.perf_counter()
                if remaining <= 0:
                    timed_out = True
                    break
                try:
                    line = self.lines.get(timeout=remaining)
                except queue.Empty:
                    timed_out = True
                    break
                if line is None:
                    break
                try:
                    event = json.loads(line)
                except json.JSONDecodeError:
                    continue
                if event.get("event") == "done":
                    done = event
                    break
                events.append(event)
        wall_ms = (time.perf_counter() - start) * 1000
        if done is not None and done.get("output"):
            target.write_bytes(base64.b64decode(done["output"]))
        if done is None:
            self.close(kill=True)
            returncode = self.proc.returncode
        else:
            returncode = 0
        return (
            {
                "returncode": returncode,
                "timed_out": timed_out,
                "elapsed_ms": round(done["elapsed_ms"], 3)
                if done
                else round(wall_ms, 3),
                "transfer_wall_ms": round(wall_ms, 3),
                "stdout": "".join(json.dumps(e) + "\n" for e in events),
                "stderr": self._stderr_tail(),
            },
            done is not None,
        )

    def close(self, kill: bool = False) -> None:
        if kill:
            subprocess.run(
                docker("kill", self.name),
                capture_output=True,
                check=False,
                timeout=CALL_TIMEOUT,
            )
            self.proc.kill()
        elif self.proc.stdin is not None:
            try:
                self.proc.stdin.close()
            except OSError:
                pass
        try:
            self.proc.wait(timeout=CALL_TIMEOUT)
        except subprocess.TimeoutExpired:
            subprocess.run(
                docker("kill", self.name),
                capture_output=True,
                check=False,
                timeout=CALL_TIMEOUT,
            )
            self.proc.kill()
            self.proc.wait()


class SessionPool:
    """Idle helper sessions for one image; each borrowing thread gets a session of its own."""

    def __init__(self, image: str) -> None:
        self.image = image
        self.idle: list[_Session] = []
        self.lock = threading.Lock()

    def run(
        self, source: Path, target: Path, timeout: float, *, edit: dict[str, Any]
    ) -> dict[str, Any]:
        with self.lock:
            session = self.idle.pop() if self.idle else None
        if session is None:
            session = _Session(self.image)
        result, reusable = session.request(source, target, timeout, edit)
        if reusable:
            with self.lock:
                self.idle.append(session)
        return result

    def close(self) -> None:
        with self.lock:
            sessions, self.idle = self.idle, []
        for session in sessions:
            session.close()
