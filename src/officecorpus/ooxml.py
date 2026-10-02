"""OOXML package inspection.

Part inventory, feature detection, package validity, input-to-output part
mapping, and the canonical XML form used by the semantic-preservation check.
Everything here is pure: it reads bytes and returns data.
"""

from __future__ import annotations

import hashlib
import posixpath
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from urllib.parse import unquote

from lxml import etree

CT_PART = "[content_types].xml"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL_NAMESPACES = frozenset(
    {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "http://purl.oclc.org/ooxml/officeDocument/relationships",
    }
)
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
MC_PREFIX_LIST_ATTRS = frozenset(
    f"{{{MC_NS}}}{name}"
    for name in (
        "Ignorable",
        "ProcessContent",
        "MustUnderstand",
        "PreserveElements",
        "PreserveAttributes",
    )
)
OFFICE_DOCUMENT_REL_SUFFIXES = ("/officeDocument",)

# Metadata that any save is expected to refresh. Dropped before semantic comparison.
VOLATILE_ELEMENTS = {
    "docprops/core.xml": frozenset(
        {
            "{http://purl.org/dc/terms/}modified",
            "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}lastModifiedBy",
            "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}revision",
        }
    ),
    "docprops/app.xml": frozenset(
        {
            "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Application",
            "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}AppVersion",
            "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}TotalTime",
        }
    ),
}

_XML_PARSER = etree.XMLParser(
    resolve_entities=False,
    no_network=True,
    huge_tree=True,
    remove_comments=True,
    remove_pis=True,
)


class PackageError(Exception):
    """The file is not a readable OPC package."""


@dataclass(frozen=True)
class Relationship:
    rel_id: str
    rel_type: str
    target: (
        str  # resolved lowercase part key for internal targets, raw target for external
    )
    external: bool


@dataclass
class Package:
    """An opened OPC package. Part keys are lowercase names without a leading slash."""

    names: dict[str, str]  # key -> original zip member name
    data: dict[str, bytes]
    defaults: dict[str, str]  # extension (lowercase) -> content type
    overrides: dict[str, str]  # part key -> content type
    content_types_error: str | None
    _rels: dict[str, list[Relationship]] = field(default_factory=dict)

    def content_type(self, key: str) -> str | None:
        if key in self.overrides:
            return self.overrides[key]
        ext = key.rsplit(".", 1)[-1] if "." in posixpath.basename(key) else ""
        return self.defaults.get(ext)

    def rels_of(self, source: str) -> list[Relationship]:
        """Relationships whose source is `source` ('' is the package root)."""
        if source not in self._rels:
            self._rels[source] = _parse_rels(self, source)
        return self._rels[source]

    def sha256(self, key: str) -> str:
        return hashlib.sha256(self.data[key]).hexdigest()


def read_package(path: Path) -> Package:
    try:
        archive = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError, ValueError) as exc:
        raise PackageError(f"not a zip archive: {exc}") from exc
    names: dict[str, str] = {}
    data: dict[str, bytes] = {}
    with archive:
        for info in archive.infolist():
            if info.is_dir():
                continue
            key = info.filename.lstrip("/").lower()
            if key in names:
                raise PackageError(
                    f"duplicate part name (case-insensitive): {info.filename}"
                )
            try:
                data[key] = archive.read(info)
            except Exception as exc:  # zlib, CRC, encryption flag
                raise PackageError(
                    f"unreadable zip member {info.filename}: {exc}"
                ) from exc
            names[key] = info.filename
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    ct_error: str | None = None
    if CT_PART not in data:
        ct_error = "missing [Content_Types].xml"
    else:
        try:
            root = etree.fromstring(data[CT_PART], _XML_PARSER)
            for child in root:
                tag = etree.QName(child).localname if isinstance(child.tag, str) else ""
                if tag == "Default":
                    defaults[child.get("Extension", "").lower()] = child.get(
                        "ContentType", ""
                    )
                elif tag == "Override":
                    overrides[child.get("PartName", "").lstrip("/").lower()] = (
                        child.get("ContentType", "")
                    )
        except etree.XMLSyntaxError as exc:
            ct_error = f"[Content_Types].xml is not well-formed: {exc}"
    return Package(names, data, defaults, overrides, ct_error)


def rels_key_for(source: str) -> str:
    if source == "":
        return "_rels/.rels"
    directory, base = posixpath.split(source)
    return posixpath.join(directory, "_rels", base + ".rels")


def source_for_rels_key(key: str) -> str | None:
    """Inverse of rels_key_for; None if `key` is not a relationships part."""
    if key == "_rels/.rels":
        return ""
    directory, base = posixpath.split(key)
    if not base.endswith(".rels") or posixpath.basename(directory) != "_rels":
        return None
    return posixpath.join(posixpath.dirname(directory), base[: -len(".rels")])


def resolve_target(source: str, target: str) -> str:
    target = unquote(target.split("#", 1)[0])
    if target.startswith("/"):
        resolved = target.lstrip("/")
    else:
        resolved = posixpath.normpath(posixpath.join(posixpath.dirname(source), target))
    return resolved.lower()


def _parse_rels(pkg: Package, source: str) -> list[Relationship]:
    raw = pkg.data.get(rels_key_for(source))
    if raw is None:
        return []
    try:
        root = etree.fromstring(raw, _XML_PARSER)
    except etree.XMLSyntaxError:
        return []
    rels: list[Relationship] = []
    for child in root:
        if (
            not isinstance(child.tag, str)
            or etree.QName(child).localname != "Relationship"
        ):
            continue
        external = child.get("TargetMode", "Internal") == "External"
        target = child.get("Target", "")
        rels.append(
            Relationship(
                rel_id=child.get("Id", ""),
                rel_type=child.get("Type", ""),
                target=target if external else resolve_target(source, target),
                external=external,
            )
        )
    return rels


def is_rels_part(key: str) -> bool:
    return key.endswith(".rels")


def is_xml_part(pkg: Package, key: str) -> bool:
    if key == CT_PART or is_rels_part(key):
        return True
    ct = pkg.content_type(key) or ""
    # "+xml" content types also end in "xml".
    return ct.endswith("xml") or key.endswith((".xml", ".vml"))


# ---------------------------------------------------------------------------
# Package validity
# ---------------------------------------------------------------------------


def package_problems(pkg: Package) -> list[str]:
    """Structural problems: content types, root relationship, dangling internal targets."""
    problems: list[str] = []
    if pkg.content_types_error:
        problems.append(pkg.content_types_error)
    for key in sorted(pkg.data):
        if key == CT_PART or is_rels_part(key):
            continue
        if pkg.content_type(key) is None:
            problems.append(f"no content type for part {key}")
    for key in sorted(pkg.data):
        if not is_rels_part(key):
            continue
        try:
            etree.fromstring(pkg.data[key], _XML_PARSER)
        except etree.XMLSyntaxError:
            problems.append(f"relationships part {key} is not well-formed")
    root_rels = pkg.rels_of("")
    if "_rels/.rels" not in pkg.data:
        problems.append("missing _rels/.rels")
    elif not any(
        rel.rel_type.endswith(OFFICE_DOCUMENT_REL_SUFFIXES) and not rel.external
        for rel in root_rels
    ):
        problems.append("no officeDocument relationship in _rels/.rels")
    for key in sorted(pkg.data):
        source = source_for_rels_key(key)
        if source is None:
            continue
        for rel in pkg.rels_of(source):
            if not rel.external and rel.target not in pkg.data:
                problems.append(
                    f"dangling relationship {source or '/'} -> {rel.target}"
                )
    return problems


def main_part(pkg: Package) -> str | None:
    for rel in pkg.rels_of(""):
        if rel.rel_type.endswith(OFFICE_DOCUMENT_REL_SUFFIXES) and not rel.external:
            return rel.target
    return None


# ---------------------------------------------------------------------------
# Exclusion screening
# ---------------------------------------------------------------------------

OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
MACRO_CONTENT_TYPE = re.compile(r"macroEnabled|vbaProject|vbaData", re.IGNORECASE)


def screen(path: Path) -> tuple[str | None, Package | None]:
    """Return (exclusion reason, package). Reason None means eligible."""
    head = path.read_bytes()[:8]
    if head == OLE_MAGIC:
        return (
            "encrypted-or-legacy: OLE compound file (password-protected OOXML or binary format)",
            None,
        )
    try:
        pkg = read_package(path)
    except PackageError as exc:
        return f"corrupt: {exc}", None
    if "encryptioninfo" in pkg.data or "encryptedpackage" in pkg.data:
        return "encrypted: EncryptionInfo stream", None
    for key in pkg.data:
        if posixpath.basename(key) in ("vbaproject.bin", "vbadata.xml"):
            return f"macro-enabled: contains {key}", None
    for ct in list(pkg.overrides.values()) + list(pkg.defaults.values()):
        if MACRO_CONTENT_TYPE.search(ct):
            return f"macro-enabled: content type {ct}", None
    if pkg.content_types_error:
        return f"corrupt: {pkg.content_types_error}", None
    main = main_part(pkg)
    if main is None or main not in pkg.data:
        return "corrupt: main document part missing or unresolved", None
    try:
        etree.fromstring(pkg.data[main], _XML_PARSER)
    except etree.XMLSyntaxError as exc:
        return f"corrupt: main part not well-formed XML: {exc}", None
    return None, pkg


# ---------------------------------------------------------------------------
# Feature detection
# ---------------------------------------------------------------------------

_PART_FEATURES: dict[str, list[tuple[str, re.Pattern[str]]]] = {
    "common": [
        ("charts", re.compile(r"(^|/)charts/chart[^/]*\.xml$")),
        ("chart-styles", re.compile(r"(^|/)charts/(style|colors)[^/]*\.xml$")),
        ("smartart", re.compile(r"(^|/)diagrams/data[^/]*\.xml$")),
        ("embedded-objects", re.compile(r"(^|/)embeddings/")),
        (
            "images",
            re.compile(r"(^|/)media/[^/]+\.(png|jpe?g|gif|bmp|tiff?|emf|wmf|svg)$"),
        ),
        (
            "audio-video",
            re.compile(r"(^|/)media/[^/]+\.(mp3|wav|m4a|mp4|m4v|wmv|avi|mov)$"),
        ),
        ("custom-xml", re.compile(r"^customxml/")),
        ("theme", re.compile(r"(^|/)theme/theme[^/]*\.xml$")),
        ("thumbnail", re.compile(r"^docprops/thumbnail")),
        ("custom-properties", re.compile(r"^docprops/custom\.xml$")),
        ("ink", re.compile(r"(^|/)ink/")),
        ("activex", re.compile(r"(^|/)activex/")),
        ("digital-signature", re.compile(r"^_xmlsignatures/")),
        ("vml-drawings", re.compile(r"\.vml$")),
    ],
    "xlsx": [
        ("chartsheets", re.compile(r"^xl/chartsheets/")),
        ("pivots", re.compile(r"^xl/pivottables/")),
        ("pivot-caches", re.compile(r"^xl/pivotcache/")),
        ("comments", re.compile(r"^xl/comments[^/]*\.xml$")),
        ("threaded-comments", re.compile(r"^xl/threadedcomments/")),
        ("tables", re.compile(r"^xl/tables/")),
        ("drawings", re.compile(r"^xl/drawings/drawing[^/]*\.xml$")),
        ("external-links", re.compile(r"^xl/externallinks/")),
        ("shared-strings", re.compile(r"^xl/sharedstrings\.xml$")),
        ("calc-chain", re.compile(r"^xl/calcchain\.xml$")),
        ("query-tables", re.compile(r"^xl/querytables/")),
        ("connections", re.compile(r"^xl/connections\.xml$")),
        ("slicers", re.compile(r"^xl/slicers?(caches)?/")),
        ("timelines", re.compile(r"^xl/timeline")),
        ("rich-data", re.compile(r"^xl/richdata/")),
        ("printer-settings", re.compile(r"^xl/printersettings/")),
        ("persons", re.compile(r"^xl/persons/")),
    ],
    "pptx": [
        ("comments", re.compile(r"^ppt/comments/")),
        ("notes", re.compile(r"^ppt/notesslides/")),
        ("handout-master", re.compile(r"^ppt/handoutmasters/")),
        ("tags", re.compile(r"^ppt/tags/")),
        ("printer-settings", re.compile(r"^ppt/printersettings/")),
    ],
    "docx": [
        ("comments", re.compile(r"^word/comments\.xml$")),
        (
            "comments-extended",
            re.compile(r"^word/comments(extended|ids|extensible)\.xml$"),
        ),
        ("footnotes", re.compile(r"^word/footnotes\.xml$")),
        ("endnotes", re.compile(r"^word/endnotes\.xml$")),
        ("headers", re.compile(r"^word/header[^/]*\.xml$")),
        ("footers", re.compile(r"^word/footer[^/]*\.xml$")),
        ("numbering", re.compile(r"^word/numbering\.xml$")),
        ("glossary", re.compile(r"^word/glossary/")),
        ("embedded-fonts", re.compile(r"^word/fonts/")),
        ("people", re.compile(r"^word/people\.xml$")),
    ],
}

# (feature, part regex, byte pattern): the pattern must occur inside a matching part.
_CONTENT_FEATURES: dict[str, list[tuple[str, re.Pattern[str], re.Pattern[bytes]]]] = {
    "xlsx": [
        (
            "conditional-formatting",
            re.compile(r"^xl/worksheets/"),
            re.compile(rb"<(\w+:)?conditionalFormatting\b"),
        ),
        (
            "data-validation",
            re.compile(r"^xl/worksheets/"),
            re.compile(rb"<(\w+:)?dataValidations?\b"),
        ),
        ("formulas", re.compile(r"^xl/worksheets/"), re.compile(rb"<(\w+:)?f[\s>/]")),
        (
            "merged-cells",
            re.compile(r"^xl/worksheets/"),
            re.compile(rb"<(\w+:)?mergeCell\b"),
        ),
        (
            "hyperlinks",
            re.compile(r"^xl/worksheets/"),
            re.compile(rb"<(\w+:)?hyperlink\b"),
        ),
        (
            "sparklines",
            re.compile(r"^xl/worksheets/"),
            re.compile(rb"sparklineGroup\b"),
        ),
        (
            "autofilter",
            re.compile(r"^xl/worksheets/"),
            re.compile(rb"<(\w+:)?autoFilter\b"),
        ),
        (
            "frozen-panes",
            re.compile(r"^xl/worksheets/"),
            re.compile(rb"<(\w+:)?pane\b"),
        ),
        (
            "sheet-protection",
            re.compile(r"^xl/worksheets/"),
            re.compile(rb"<(\w+:)?sheetProtection\b"),
        ),
        (
            "outline-grouping",
            re.compile(r"^xl/worksheets/"),
            re.compile(rb"outlineLevel="),
        ),
        (
            "rich-text",
            re.compile(r"^xl/sharedstrings\.xml$"),
            re.compile(rb"<(\w+:)?rPr\b"),
        ),
        (
            "defined-names",
            re.compile(r"^xl/workbook\.xml$"),
            re.compile(rb"<(\w+:)?definedName\b"),
        ),
        ("array-formulas", re.compile(r"^xl/worksheets/"), re.compile(rb't="array"')),
        ("shared-formulas", re.compile(r"^xl/worksheets/"), re.compile(rb't="shared"')),
        (
            "number-formats",
            re.compile(r"^xl/styles\.xml$"),
            re.compile(rb"<(\w+:)?numFmt\b"),
        ),
        (
            "differential-styles",
            re.compile(r"^xl/styles\.xml$"),
            re.compile(rb"<(\w+:)?dxf\b"),
        ),
        ("shapes", re.compile(r"^xl/drawings/drawing"), re.compile(rb"<(\w+:)?sp\b")),
        (
            "print-setup",
            re.compile(r"^xl/worksheets/"),
            re.compile(rb"<(\w+:)?pageSetup\b"),
        ),
    ],
    "pptx": [
        ("tables", re.compile(r"^ppt/slides/"), re.compile(rb"<a:tbl\b")),
        ("animations", re.compile(r"^ppt/slides/"), re.compile(rb"<p:timing\b")),
        ("transitions", re.compile(r"^ppt/slides/"), re.compile(rb"<p:transition\b")),
        ("group-shapes", re.compile(r"^ppt/slides/"), re.compile(rb"<p:grpSp\b")),
        ("hyperlinks", re.compile(r"^ppt/slides/"), re.compile(rb"<a:hlinkClick\b")),
        ("ole-objects", re.compile(r"^ppt/slides/"), re.compile(rb"<p:oleObj\b")),
        ("connectors", re.compile(r"^ppt/slides/"), re.compile(rb"<p:cxnSp\b")),
        ("custom-geometry", re.compile(r"^ppt/slides/"), re.compile(rb"<a:custGeom\b")),
        (
            "gradient-fills",
            re.compile(r"^ppt/(slides|slidelayouts|slidemasters)/"),
            re.compile(rb"<a:gradFill\b"),
        ),
        (
            "sections",
            re.compile(r"^ppt/presentation\.xml$"),
            re.compile(rb"sectionLst\b"),
        ),
        (
            "alternate-content",
            re.compile(r"^ppt/slides/"),
            re.compile(rb"<mc:AlternateContent\b"),
        ),
        ("placeholders", re.compile(r"^ppt/slides/"), re.compile(rb"<p:ph\b")),
    ],
    "docx": [
        (
            "tracked-changes",
            re.compile(r"^word/document\.xml$"),
            re.compile(rb"<w:(ins|del|moveFrom|moveTo|rPrChange|pPrChange)\b"),
        ),
        ("tables", re.compile(r"^word/document\.xml$"), re.compile(rb"<w:tbl\b")),
        (
            "fields",
            re.compile(r"^word/(document|header\d*|footer\d*)\.xml$"),
            re.compile(rb"<w:(fldSimple|instrText|fldChar)\b"),
        ),
        (
            "content-controls",
            re.compile(r"^word/document\.xml$"),
            re.compile(rb"<w:sdt\b"),
        ),
        ("math", re.compile(r"^word/document\.xml$"), re.compile(rb"<m:oMath")),
        (
            "text-boxes",
            re.compile(r"^word/document\.xml$"),
            re.compile(rb"<w:txbxContent\b|<wps:txbx\b"),
        ),
        (
            "bookmarks",
            re.compile(r"^word/document\.xml$"),
            re.compile(rb"<w:bookmarkStart\b"),
        ),
        (
            "hyperlinks",
            re.compile(r"^word/document\.xml$"),
            re.compile(rb"<w:hyperlink\b"),
        ),
        (
            "columns",
            re.compile(r"^word/document\.xml$"),
            re.compile(rb'<w:cols\b[^>]*w:num="[2-9]'),
        ),
        ("drawings", re.compile(r"^word/document\.xml$"), re.compile(rb"<w:drawing\b")),
        ("legacy-vml", re.compile(r"^word/document\.xml$"), re.compile(rb"<w:pict\b")),
        (
            "comments-anchored",
            re.compile(r"^word/document\.xml$"),
            re.compile(rb"<w:commentRangeStart\b"),
        ),
        (
            "alternate-content",
            re.compile(r"^word/document\.xml$"),
            re.compile(rb"<mc:AlternateContent\b"),
        ),
        ("styles", re.compile(r"^word/styles\.xml$"), re.compile(rb"<w:style\b")),
        ("lists", re.compile(r"^word/document\.xml$"), re.compile(rb"<w:numPr\b")),
    ],
}


def detect_features(pkg: Package, fmt: str) -> list[str]:
    found: set[str] = set()
    keys = sorted(pkg.data)
    for name, pattern in _PART_FEATURES["common"] + _PART_FEATURES[fmt]:
        if any(pattern.search(key) for key in keys):
            found.add(name)
    for name, part_pattern, content_pattern in _CONTENT_FEATURES[fmt]:
        if name in found:
            continue
        for key in keys:
            if part_pattern.search(key) and content_pattern.search(pkg.data[key]):
                found.add(name)
                break
    if (
        fmt == "pptx"
        and sum(
            1 for k in keys if re.match(r"^ppt/slidemasters/slidemaster[^/]*\.xml$", k)
        )
        > 1
    ):
        found.add("multiple-masters")
    if (
        fmt == "xlsx"
        and sum(1 for k in keys if re.match(r"^xl/worksheets/sheet[^/]*\.xml$", k)) > 1
    ):
        found.add("multiple-sheets")
    if fmt == "docx" and pkg.data.get("word/document.xml", b"").count(b"<w:sectPr") > 1:
        found.add("sections")
    return sorted(found)


def part_inventory(pkg: Package) -> dict[str, Any]:
    return {
        "part_count": len(pkg.data),
        "xml_bytes": sum(len(v) for k, v in pkg.data.items() if is_xml_part(pkg, k)),
        "main_part": main_part(pkg),
    }


# ---------------------------------------------------------------------------
# Part mapping (input -> output)
# ---------------------------------------------------------------------------


def _natural(text: str) -> list[Any]:
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", text)]


def logical_keys(pkg: Package) -> dict[str, str]:
    """Map each reachable part to its relationship path from the package root.

    A logical key is the chain of relationship types (short form) with an ordinal
    among same-typed siblings, ordered by natural target-name order. Breadth-first,
    first reach wins.
    """
    keys: dict[str, str] = {}
    queue: list[tuple[str, str]] = [("", "")]
    seen_sources = {""}
    while queue:
        source, source_key = queue.pop(0)
        internal = [rel for rel in pkg.rels_of(source) if not rel.external]
        internal.sort(key=lambda rel: (rel.rel_type, _natural(rel.target)))
        counters: dict[str, int] = {}
        for rel in internal:
            short = rel.rel_type.rsplit("/", 1)[-1]
            counters[short] = counters.get(short, 0) + 1
            child_key = f"{source_key}/{short}#{counters[short]}"
            if rel.target not in keys:
                keys[rel.target] = child_key
            if rel.target not in seen_sources and rel.target in pkg.data:
                seen_sources.add(rel.target)
                queue.append((rel.target, keys[rel.target]))
    return keys


@dataclass
class PartMapping:
    mapping: dict[str, str]  # input key -> output key
    methods: dict[
        str, str
    ]  # input key -> name | relationship-path | identical-bytes | relationships-of-mapped-source
    missing: list[str]  # input parts with no counterpart
    added: list[str]  # output parts with no input counterpart


def map_parts(inp: Package, out: Package) -> PartMapping:
    mapping: dict[str, str] = {}
    methods: dict[str, str] = {}
    used: set[str] = set()
    non_rels = [k for k in sorted(inp.data) if not is_rels_part(k)]
    for key in non_rels:
        if key in out.data:
            mapping[key], methods[key] = key, "name"
            used.add(key)
    in_logical = logical_keys(inp)
    out_by_logical: dict[str, str] = {}
    for key, logical in logical_keys(out).items():
        out_by_logical.setdefault(logical, key)
    for key in non_rels:
        if key in mapping or key not in in_logical:
            continue
        candidate = out_by_logical.get(in_logical[key])
        if (
            candidate
            and candidate not in used
            and candidate in out.data
            and inp.content_type(key) == out.content_type(candidate)
        ):
            mapping[key], methods[key] = candidate, "relationship-path"
            used.add(candidate)
    out_by_hash: dict[str, list[str]] = {}
    for key in sorted(out.data):
        if key not in used and not is_rels_part(key):
            out_by_hash.setdefault(out.sha256(key), []).append(key)
    for key in non_rels:
        if key in mapping:
            continue
        pool = out_by_hash.get(inp.sha256(key))
        while pool:
            candidate = pool.pop(0)
            if candidate not in used:
                mapping[key], methods[key] = candidate, "identical-bytes"
                used.add(candidate)
                break
    for key in sorted(inp.data):
        if not is_rels_part(key):
            continue
        if key in out.data and key not in used:
            mapping[key], methods[key] = key, "name"
            used.add(key)
            continue
        source = source_for_rels_key(key)
        if source is not None and source in mapping:
            candidate = rels_key_for(mapping[source])
            if candidate in out.data and candidate not in used:
                mapping[key], methods[key] = candidate, "relationships-of-mapped-source"
                used.add(candidate)
    missing = [k for k in sorted(inp.data) if k not in mapping]
    added = [k for k in sorted(out.data) if k not in used]
    return PartMapping(mapping, methods, missing, added)


# ---------------------------------------------------------------------------
# Canonical XML
# ---------------------------------------------------------------------------


def _rel_token(
    pkg: Package, rel: Relationship, target_map: dict[str, str] | None
) -> str:
    if rel.external:
        return f"{rel.rel_type}|external|{rel.target}"
    target = rel.target
    if target_map is not None:
        target = target_map.get(target, f"unmapped:{target}")
    return f"{rel.rel_type}|internal|{target}"


def canonical_part(pkg: Package, key: str, target_map: dict[str, str] | None) -> Any:
    """Canonical comparable form of one part.

    `target_map` maps this package's part keys to the comparison namespace (the
    input package passes its input->output mapping; the output package passes None).
    """
    raw = pkg.data[key]
    source = source_for_rels_key(key)
    if source is not None:
        return (
            "rels",
            tuple(
                sorted(_rel_token(pkg, rel, target_map) for rel in pkg.rels_of(source))
            ),
        )
    if key == CT_PART:
        return (
            "content-types",
        )  # compared through effective content types of mapped parts
    if not is_xml_part(pkg, key):
        return ("bytes", hashlib.sha256(raw).hexdigest())
    try:
        root = etree.fromstring(raw, _XML_PARSER)
    except etree.XMLSyntaxError as exc:
        return ("malformed-xml", hashlib.sha256(raw).hexdigest(), str(exc)[:80])
    rel_ids = {rel.rel_id: _rel_token(pkg, rel, target_map) for rel in pkg.rels_of(key)}
    volatile = VOLATILE_ELEMENTS.get(key, frozenset())
    return ("xml", _canonical_element(root, rel_ids, volatile))


def _canonical_element(
    element: Any, rel_ids: dict[str, str], volatile: frozenset[str]
) -> Any:
    attrs = []
    for name, value in element.attrib.items():
        namespace = etree.QName(name).namespace if name.startswith("{") else None
        if namespace in OFFICE_REL_NAMESPACES and value in rel_ids:
            value = f"rel({rel_ids[value]})"
        elif name in MC_PREFIX_LIST_ATTRS:
            value = " ".join(
                sorted(_prefix_uri(element, token) for token in value.split())
            )
        attrs.append((name, value))
    attrs.sort()
    text = element.text if element.text and element.text.strip() else None
    children = []
    for child in element:
        if not isinstance(child.tag, str):
            continue
        if child.tag in volatile:
            continue
        children.append(_canonical_element(child, rel_ids, volatile))
        if child.tail and child.tail.strip():
            children.append(("tail", child.tail))
    return (element.tag, tuple(attrs), text, tuple(children))


def _prefix_uri(element: Any, token: str) -> str:
    # mc prefix lists may contain "prefix" or "prefix:local" tokens.
    prefix, _, local = token.partition(":")
    uri = element.nsmap.get(prefix, f"unbound:{prefix}")
    return f"{{{uri}}}{local}" if local else uri
