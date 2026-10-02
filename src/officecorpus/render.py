"""Render oracle: headless LibreOffice to PDF, pymupdf rasterization, per-page pixel difference.

Thresholds (documented in README):
- pages are rasterized at RENDER_DPI in RGB;
- a pixel differs when any channel differs by more than PIXEL_DELTA (0-255);
- a page passes when at most PAGE_TOLERANCE of its pixels differ;
- a file passes when input and output PDFs have the same page count, every
  compared page has the same size, and every compared page passes. At most
  MAX_PAGES leading pages are compared; the page-count check covers all pages.
"""

from __future__ import annotations

import os
import shutil
import signal
import subprocess
import time
from pathlib import Path
from typing import Any

import numpy as np
import pymupdf

from .adapters import soffice_binary

RENDER_DPI = 50
PIXEL_DELTA = 32
PAGE_TOLERANCE = 0.01
MAX_PAGES = 20

# Annotation appearance-stream warnings (e.g. Screen annotations for media) do not affect rasterization.
pymupdf.TOOLS.mupdf_display_errors(False)


def run_child(
    cmd: list[str], timeout: float, cwd: Path | None = None
) -> dict[str, Any]:
    """Run one isolated child process group; collect exit status, output, wall time, peak RSS."""
    start = time.perf_counter()
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        stdin=subprocess.DEVNULL,
        cwd=cwd,
        start_new_session=True,
    )
    timed_out = False
    try:
        stdout, stderr = proc.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        timed_out = True
        try:
            os.killpg(proc.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
        stdout, stderr = proc.communicate()
    elapsed_ms = (time.perf_counter() - start) * 1000
    # Reap any surviving group members (e.g. soffice helpers) so nothing leaks into the next file.
    try:
        os.killpg(proc.pid, signal.SIGKILL)
    except (ProcessLookupError, PermissionError):
        pass
    return {
        "returncode": proc.returncode,
        "timed_out": timed_out,
        "elapsed_ms": round(elapsed_ms, 3),
        "stdout": stdout.decode("utf-8", "replace"),
        "stderr": stderr.decode("utf-8", "replace"),
    }


def to_pdf(
    source: Path, out_dir: Path, profile: Path, timeout: float
) -> tuple[Path | None, dict[str, Any]]:
    """Convert one document to PDF with a private LibreOffice profile."""
    out_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        soffice_binary(),
        f"-env:UserInstallation={profile.as_uri()}",
        "--headless",
        "--norestore",
        "--nolockcheck",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_dir),
        str(source),
    ]
    result = run_child(cmd, timeout)
    pdf = out_dir / (source.stem + ".pdf")
    if result["timed_out"] or not pdf.exists() or pdf.stat().st_size == 0:
        return None, result
    return pdf, result


def _page_pixels(page: Any) -> np.ndarray:
    pix = page.get_pixmap(dpi=RENDER_DPI, colorspace=pymupdf.csRGB, alpha=False)
    return np.frombuffer(pix.samples, dtype=np.uint8).reshape(
        pix.height, pix.width, pix.n
    )


def compare_pdfs(reference: Path, candidate: Path) -> dict[str, Any]:
    with pymupdf.open(reference) as ref, pymupdf.open(candidate) as cand:
        ref_pages, cand_pages = ref.page_count, cand.page_count
        pages: list[dict[str, Any]] = []
        for index in range(min(ref_pages, cand_pages, MAX_PAGES)):
            a = _page_pixels(ref[index])
            b = _page_pixels(cand[index])
            if a.shape != b.shape:
                pages.append(
                    {
                        "page": index + 1,
                        "size_mismatch": True,
                        "diff_fraction": 1.0,
                        "pass": False,
                    }
                )
                continue
            delta = np.abs(a.astype(np.int16) - b.astype(np.int16)).max(axis=2)
            fraction = float((delta > PIXEL_DELTA).mean())
            pages.append(
                {
                    "page": index + 1,
                    "diff_fraction": round(fraction, 6),
                    "pass": fraction <= PAGE_TOLERANCE,
                }
            )
    passed = ref_pages == cand_pages and all(p["pass"] for p in pages)
    return {
        "pass": passed,
        "input_pages": ref_pages,
        "output_pages": cand_pages,
        "compared_pages": len(pages),
        "max_diff_fraction": max((p["diff_fraction"] for p in pages), default=None),
        "pages": pages,
    }


def thresholds() -> dict[str, Any]:
    return {
        "dpi": RENDER_DPI,
        "pixel_channel_delta": PIXEL_DELTA,
        "page_max_diff_fraction": PAGE_TOLERANCE,
        "max_compared_pages": MAX_PAGES,
        "file_rule": "equal page count and every compared page within page_max_diff_fraction",
    }


def reset_dir(path: Path) -> None:
    if path.exists():
        shutil.rmtree(path)
    path.mkdir(parents=True)
