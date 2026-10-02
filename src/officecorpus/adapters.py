"""Adapter registry and isolated environment provisioning.

Every adapter runs in a child process and applies the run's one edit (see edits.py)
before saving. Python adapters each get a private uv virtual environment under
`.envs/<adapter-id>`; Node adapters share the pinned `adapters/node` package; JVM, .NET,
and Go adapters run small helpers under `adapters/<runtime>/<library>/` in pinned Docker
images (see docker_helper.py); LibreOffice runs a Python-UNO script
(`adapters/libreoffice/`) inside the installed soffice with a fresh profile per file.
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import sys
import tarfile
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from . import docker_helper

PYTHON_VERSION = "3.12"
# Adapters installed from a wheel built out of a local checkout's committed HEAD.
LOCAL_SOURCES = {
    "wolfppt-wheel": Path("~/Projects/wolfppt").expanduser(),
    "wolfdocx-wheel": Path("~/Projects/wolfdocx").expanduser(),
}


@dataclass(frozen=True)
class AdapterSpec:
    id: str
    library: str  # officelibs catalog slug
    format: str
    kind: str  # python | node | docker | soffice | unsupported
    language: str
    install: tuple[str, ...] = ()
    mode: str = "default"
    unsupported_reason: str | None = None
    image: str | None = None  # docker: image tag
    helper: str | None = None  # docker: repo-relative helper source directory


POI_HELPER = "adapters/jvm/apache-poi"
POI_IMAGE = "officecorpus/apache-poi:5.5.1"
OPENXML_HELPER = "adapters/dotnet/open-xml-sdk"
OPENXML_IMAGE = "officecorpus/open-xml-sdk:3.5.1"
OPENXML_SAVE = (
    "input copied into an expandable MemoryStream, default OpenSettings; Save(), Dispose(), "
    "stream bytes written to the output"
)


ADAPTERS: tuple[AdapterSpec, ...] = (
    AdapterSpec(
        "openpyxl",
        "openpyxl",
        "xlsx",
        "python",
        "python",
        ("openpyxl==3.1.5",),
        mode="load_workbook(path); wb[sheet][cell] = value; save(output)",
    ),
    AdapterSpec(
        "wolfxl",
        "wolfxl",
        "xlsx",
        "python",
        "python",
        ("wolfxl",),
        mode=(
            "load_workbook(path, modify=True), the documented existing-workbook edit mode; "
            "wb[sheet][cell] = value; save(output); latest PyPI release"
        ),
    ),
    AdapterSpec(
        "exceljs",
        "exceljs",
        "xlsx",
        "node",
        "javascript",
        mode=(
            "new Workbook(); xlsx.readFile(path); getWorksheet(name).getCell(ref).value = string; "
            "xlsx.writeFile(output)"
        ),
    ),
    AdapterSpec(
        "sheetjs",
        "sheetjs",
        "xlsx",
        "node",
        "javascript",
        mode=(
            "XLSX.readFile(path); utils.sheet_add_aoa(Sheets[name], [[string]], {origin: ref}) "
            "(extends !ref); XLSX.writeFile(workbook, output); default options"
        ),
    ),
    AdapterSpec(
        "libreoffice",
        "libreoffice",
        "xlsx",
        "soffice",
        "c++",
        mode=(
            "Python-UNO script in headless soffice with a fresh profile per file "
            "(adapters/libreoffice/): loadComponentFromURL, Sheets.getByName(sheet).getCellRangeByName(cell).setString(value), storeToURL with the Calc MS Excel 2007 XML filter"
        ),
    ),
    AdapterSpec(
        "apache-poi",
        "apache-poi",
        "xlsx",
        "docker",
        "java",
        mode=(
            "new XSSFWorkbook(InputStream); getSheet(name).getRow/createRow, getCell/createCell, "
            "setCellValue(String); write(OutputStream)"
        ),
        image=POI_IMAGE,
        helper=POI_HELPER,
    ),
    AdapterSpec(
        "excelize",
        "excelize",
        "xlsx",
        "docker",
        "go",
        mode=(
            "excelize.OpenFile(path) with default Options; SetCellStr(sheet, ref, string); "
            "SaveAs(output) and Close()"
        ),
        image="officecorpus/excelize:2.11.0",
        helper="adapters/go/excelize",
    ),
    AdapterSpec(
        "aspose-cells-foss",
        "aspose-cells-foss",
        "xlsx",
        "python",
        "python",
        ("aspose-cells-foss==26.7.0",),
        mode=(
            "aspose.cells_foss.Workbook(path); get_worksheet_by_name(sheet).cells[cell]"
            ".put_value(value); save(output)"
        ),
    ),
    AdapterSpec(
        "python-pptx",
        "python-pptx",
        "pptx",
        "python",
        "python",
        ("python-pptx==1.0.2",),
        mode=(
            "Presentation(path); slides[i].shapes (shape_id); text_frame.paragraphs[p].runs[r]"
            ".text = value; save(output)"
        ),
    ),
    AdapterSpec(
        "wolfppt-wheel",
        "wolfppt",
        "pptx",
        "python",
        "python",
        mode=(
            "wheel built from the committed HEAD of the local wolfppt source into an isolated "
            "environment; python-pptx-compatible API: slides[i].shapes (shape_id), "
            "text_frame.paragraphs[p].runs[r].text = value; save(output)"
        ),
    ),
    AdapterSpec(
        "libreoffice",
        "libreoffice",
        "pptx",
        "soffice",
        "c++",
        mode=(
            "Python-UNO script in headless soffice with a fresh profile per file "
            "(adapters/libreoffice/): loadComponentFromURL, the shape at the target's position among the slide's top-level spTree shapes (UNO exposes no cNvPr id; the name is cross-checked), its text paragraph by enumeration, and the run's character range by text cursor setString, storeToURL with the Impress MS PowerPoint 2007 XML filter"
        ),
    ),
    AdapterSpec(
        "apache-poi",
        "apache-poi",
        "pptx",
        "docker",
        "java",
        mode=(
            "new XMLSlideShow(InputStream); getSlides() by index, XSLFTextShape of the top-level p:sp "
            "with the cNvPr id, getTextParagraphs() by index, XSLFTextRun of the n-th direct a:r, "
            "setText(String); write(OutputStream)"
        ),
        image=POI_IMAGE,
        helper=POI_HELPER,
    ),
    AdapterSpec(
        "open-xml-sdk",
        "open-xml-sdk",
        "pptx",
        "docker",
        "c#",
        mode=(
            "PresentationDocument.Open(stream, isEditable: true); SlideIdList order -> SlidePart, first "
            "direct ShapeTree Shape with NonVisualDrawingProperties.Id, TextBody paragraph/direct Run by "
            "index, run children except RunProperties removed, one A.Text appended; "
            + OPENXML_SAVE
        ),
        image=OPENXML_IMAGE,
        helper=OPENXML_HELPER,
    ),
    AdapterSpec(
        "pptx-automizer",
        "pptx-automizer",
        "pptx",
        "unsupported",
        "typescript",
        unsupported_reason=(
            "pptx-automizer cannot edit a slide of the root presentation in place (documented as limited "
            "to adding slides; modifying requires truncating the root and re-adding every slide), and its "
            "text modifiers (ModifyTextHelper.setText/replaceText) address a whole shape text body or tag "
            "patterns, not one run by paragraph/run index; addressing one a:r needs a custom XmlElement "
            "callback, i.e. raw XML patching"
        ),
    ),
    AdapterSpec(
        "python-docx",
        "python-docx",
        "docx",
        "python",
        "python",
        ("python-docx==1.2.0",),
        mode="Document(path); paragraphs[p].runs[r].text = value; save(output)",
    ),
    AdapterSpec(
        "wolfdocx-wheel",
        "wolfdocx",
        "docx",
        "python",
        "python",
        mode=(
            "wheel built from the committed HEAD of the local wolfdocx source into an isolated "
            "environment; documented CLI: `wolfdocx author report-targets <input> --part <main part>` "
            "for the target paragraph's structure fingerprint, then `wolfdocx author report <request>` "
            "with one text edit (paragraph = 1-based index of the target among the part's w:p, "
            "expected_old_text = the target run's text, which wolfdocx requires to occur once in the "
            "paragraph)"
        ),
    ),
    AdapterSpec(
        "libreoffice",
        "libreoffice",
        "docx",
        "soffice",
        "c++",
        mode=(
            "Python-UNO script in headless soffice with a fresh profile per file "
            "(adapters/libreoffice/): loadComponentFromURL, the target paragraph in the body Text "
            "enumeration (tables skipped; ordinal mapped through Writer's split at page/column breaks "
            "and its dropped empty section-break paragraphs, checked against the paragraph text with a "
            "nearest-equal-text fallback), the run's character range by text cursor setString with "
            "change tracking off, storeToURL with the MS Word 2007 XML filter"
        ),
    ),
    AdapterSpec(
        "pandoc",
        "pandoc",
        "docx",
        "unsupported",
        "haskell",
        unsupported_reason=(
            "pandoc's AST has no run or paragraph identity: the docx reader dissolves w:r runs into "
            "word-level Str/Space inlines wrapped in Strong/Emph by formatting (a run 'Department: ' "
            "becomes Strong [Str 'Department:'] followed by a Space outside it) and drops empty "
            "paragraphs, so a filter cannot address the target run; matching text would be the "
            "adapter's edit, not pandoc's"
        ),
    ),
    AdapterSpec(
        "apache-poi",
        "apache-poi",
        "docx",
        "docker",
        "java",
        mode=(
            "new XWPFDocument(InputStream); n-th XWPFParagraph of getBodyElements() (checked = n-th "
            "w:body/w:p), XWPFRun of the n-th direct w:r via getRun(CTR), run content except w:rPr "
            "cleared on getCTR(), setText(String, 0); write(OutputStream)"
        ),
        image=POI_IMAGE,
        helper=POI_HELPER,
    ),
    AdapterSpec(
        "docx4j",
        "docx4j",
        "docx",
        "docker",
        "java",
        mode=(
            "WordprocessingMLPackage.load(File); n-th P of MainDocumentPart.getContent() (unwrapped), "
            "n-th direct R of P.getContent(), R.getContent() cleared and one Text added (rPr kept); "
            "save(File)"
        ),
        image="officecorpus/docx4j:17.2.1",
        helper="adapters/jvm/docx4j",
    ),
    AdapterSpec(
        "open-xml-sdk",
        "open-xml-sdk",
        "docx",
        "docker",
        "c#",
        mode=(
            "WordprocessingDocument.Open(stream, isEditable: true); Body.Elements<Paragraph>() then "
            "direct Elements<Run>() by index, run children except RunProperties removed, one Text "
            "appended; " + OPENXML_SAVE
        ),
        image=OPENXML_IMAGE,
        helper=OPENXML_HELPER,
    ),
    AdapterSpec(
        "docxtpl",
        "docxtpl",
        "docx",
        "python",
        "python",
        ("docxtpl==0.20.2",),
        mode=(
            "DocxTemplate(path); get_docx(), docxtpl's documented handle on the base document, is a "
            "python-docx Document, so the edit is paragraphs[p].runs[r].text = value through python-docx "
            "(version in environment_packages); render({}) (save() without render() reloads the "
            "template and drops get_docx() edits); save(output)"
        ),
    ),
)


def adapters_for(fmt: str) -> list[AdapterSpec]:
    return [a for a in ADAPTERS if a.format == fmt]


def get_adapter(fmt: str, adapter_id: str) -> AdapterSpec:
    for spec in adapters_for(fmt):
        if spec.id == adapter_id:
            return spec
    known = ", ".join(a.id for a in adapters_for(fmt))
    raise ValueError(f"unknown {fmt} adapter {adapter_id!r}; known: {known}")


@dataclass
class Prepared:
    spec: AdapterSpec
    available: bool
    version: str | None = None
    reason: str | None = None
    receipt: dict[str, Any] = field(default_factory=dict)
    python: Path | None = None
    pool: docker_helper.SessionPool | None = None  # docker: persistent helper sessions

    def command(self, root: Path, source: Path, target: Path, edit: Path) -> list[str]:
        """Argv for one file; `edit` is the path of the run's edit.json for this file."""
        spec = self.spec
        if spec.kind == "python":
            assert self.python is not None
            return [
                str(self.python),
                str(root / "src/officecorpus/child_python.py"),
                spec.id,
                str(source),
                str(target),
                "--edit",
                str(edit),
            ]
        if spec.kind == "node":
            return [
                "node",
                str(root / "adapters/node/roundtrip.cjs"),
                spec.id,
                str(source),
                str(target),
                "--edit",
                str(edit),
            ]
        if spec.kind == "soffice":
            return [
                sys.executable,
                str(root / "adapters/libreoffice/officecorpus_lo.py"),
                spec.format,
                str(source),
                str(target),
                "--edit",
                str(edit),
            ]
        if spec.kind == "docker":
            # Files and the edit travel over the helper's stdin/stdout protocol.
            assert spec.image is not None
            return docker_helper.docker("run", *docker_helper.RUN_FLAGS, spec.image)
        raise RuntimeError(f"{spec.id} has no command")


def soffice_binary() -> str:
    found = shutil.which("soffice")
    if found:
        return found
    mac = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if mac.exists():
        return str(mac)
    raise FileNotFoundError("soffice not found")


def libreoffice_version() -> str | None:
    try:
        out = subprocess.run(
            [soffice_binary(), "--version"],
            capture_output=True,
            text=True,
            check=False,
            timeout=60,
        ).stdout
    except (OSError, subprocess.TimeoutExpired):
        return None
    return out.strip() or None


def _run(cmd: list[str], **kwargs: Any) -> subprocess.CompletedProcess[str]:
    return subprocess.run(cmd, capture_output=True, text=True, check=True, **kwargs)


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _child_version(python: Path, root: Path, adapter_id: str) -> str:
    out = _run(
        [
            str(python),
            str(root / "src/officecorpus/child_python.py"),
            "--version",
            adapter_id,
        ]
    ).stdout
    return json.loads(out.splitlines()[-1])["version"]


def _venv(root: Path, adapter_id: str, fresh: bool) -> Path:
    env = root / ".envs" / adapter_id
    if fresh and env.exists():
        shutil.rmtree(env)
    if not (env / "bin/python").exists():
        _run(["uv", "venv", "-q", "--python", PYTHON_VERSION, str(env)])
    return env / "bin/python"


def _freeze(python: Path) -> list[str]:
    out = _run(["uv", "pip", "freeze", "--python", str(python)]).stdout
    return sorted(
        line for line in out.splitlines() if line and not line.startswith("#")
    )


def prepare(root: Path, spec: AdapterSpec) -> Prepared:
    """Provision the adapter (fresh environment) and return its identity receipt."""
    try:
        if spec.kind == "unsupported":
            return Prepared(spec, False, reason=spec.unsupported_reason)
        if spec.kind == "python":
            python = _venv(root, spec.id, fresh=True)
            receipt: dict[str, Any] = {}
            if spec.id in LOCAL_SOURCES:
                receipt = _install_local_wheel(
                    root, python, spec.library, LOCAL_SOURCES[spec.id]
                )
            else:
                _run(
                    [
                        "uv",
                        "pip",
                        "install",
                        "-q",
                        "--python",
                        str(python),
                        *spec.install,
                    ]
                )
            receipt["environment_packages"] = _freeze(python)
            receipt["python"] = _run(
                [str(python), "-c", "import sys; print(sys.version.split()[0])"]
            ).stdout.strip()
            return Prepared(
                spec,
                True,
                _child_version(python, root, spec.id),
                receipt=receipt,
                python=python,
            )
        if spec.kind == "node":
            node_dir = root / "adapters/node"
            if not (node_dir / "node_modules").exists():
                _run(["npm", "ci", "--no-audit", "--no-fund"], cwd=node_dir)
            out = _run(
                ["node", str(node_dir / "roundtrip.cjs"), "--version", spec.id]
            ).stdout
            node_version = _run(["node", "--version"]).stdout.strip()
            return Prepared(
                spec,
                True,
                json.loads(out.splitlines()[-1])["version"],
                receipt={
                    "node": node_version,
                    "lockfile_sha256": _sha256(node_dir / "package-lock.json"),
                    "helper_source": "adapters/node/roundtrip.cjs",
                    "helper_source_sha256": _sha256(node_dir / "roundtrip.cjs"),
                },
            )
        if spec.kind == "soffice":
            version = libreoffice_version()
            if version is None:
                return Prepared(spec, False, reason="soffice is not installed")
            script_dir = root / "adapters/libreoffice"
            return Prepared(
                spec,
                True,
                version.split()[1],
                receipt={
                    "soffice_version": version,
                    "helper_source": {
                        f"adapters/libreoffice/{p.name}": _sha256(p)
                        for p in sorted(script_dir.glob("*.py"))
                    },
                },
            )
        if spec.kind == "docker":
            assert spec.image is not None and spec.helper is not None
            version, receipt = docker_helper.provision(root, spec.helper, spec.image)
            return Prepared(
                spec,
                True,
                version,
                receipt=receipt,
                pool=docker_helper.SessionPool(spec.image),
            )
    except (
        OSError,
        subprocess.CalledProcessError,
        subprocess.TimeoutExpired,
    ) as exc:
        detail = getattr(exc, "stderr", None) or str(exc)
        return Prepared(
            spec, False, reason=f"provisioning failed: {str(detail)[-1500:]}"
        )
    raise RuntimeError(spec.kind)


def _install_local_wheel(
    root: Path, python: Path, name: str, src: Path
) -> dict[str, Any]:
    """Build a wheel from the committed HEAD of a local checkout (out of tree) and install it."""
    commit = _run(["git", "-C", str(src), "rev-parse", "HEAD"]).stdout.strip()
    dirty = bool(
        _run(
            ["git", "-C", str(src), "status", "--porcelain", "--untracked-files=no"]
        ).stdout.strip()
    )
    target_dir = root / f".cache/{name}-cargo-target"
    with tempfile.TemporaryDirectory(prefix=f"officecorpus-{name}-") as tmp:
        tmp_path = Path(tmp)
        archive = tmp_path / "src.tar"
        with archive.open("wb") as handle:
            subprocess.run(
                ["git", "-C", str(src), "archive", "--format=tar", commit],
                stdout=handle,
                check=True,
            )
        build_src = tmp_path / "src"
        build_src.mkdir()
        with tarfile.open(archive) as handle:
            handle.extractall(build_src, filter="data")
        wheel_dir = tmp_path / "dist"
        env = {**os.environ, "CARGO_TARGET_DIR": str(target_dir)}
        _run(
            [
                "uv",
                "build",
                "--wheel",
                "--python",
                str(python),
                "--out-dir",
                str(wheel_dir),
                str(build_src),
            ],
            env=env,
        )
        wheels = sorted(wheel_dir.glob(f"{name}-*.whl"))
        if len(wheels) != 1:
            raise RuntimeError(
                f"expected one {name} wheel, found {[w.name for w in wheels]}"
            )
        wheel = wheels[0]
        receipt = {
            "source_repo": name,
            "source_revision": commit,
            "source_worktree_dirty": dirty,
            "build_mode": "uv build --wheel from git archive of HEAD (out of tree)",
            "artifact": {"name": wheel.name, "sha256": _sha256(wheel)},
        }
        _run(["uv", "pip", "install", "-q", "--python", str(python), str(wheel)])
    return receipt
