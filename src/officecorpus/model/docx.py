"""Reader-observable content model for WordprocessingML (.docx) packages.

`extract` turns a package into a deterministic, JSON-serializable model of what a
reader of the document can observe; `compare` diffs two models. A library that
re-serializes a document into an equivalent form produces the same model as one
that copies the original bytes, so both score as preserved.

Package and markup rules
  - Parts are found through relationships by type (the last segment of the
    relationship Type URI) and resolved target, never by relationship id or part
    name. Part names compare case-insensitively.
  - ISO/IEC 29500 Strict namespaces are mapped to their Transitional equivalents.
  - mc:AlternateContent contributes exactly one branch: the first mc:Choice, or
    mc:Fallback when there is no Choice.
  - ST_OnOff / xsd:boolean values compare by value ("1", "true", "on" are true; an
    on/off element without w:val is true). Element and attribute order never
    matters: everything is read by name.
  - Measurements are integers in twips; universal measures ("1in", "2.5cm", "12pt")
    are converted. Drawing extents (EMU) and VML sizes (CSS lengths) are converted
    to twips and rounded.
  - Colors are six-digit uppercase hex. An explicit w:val wins (Word writes the
    resolved theme color there); otherwise a theme color is resolved through the
    document theme with themeTint/themeShade applied linearly per channel. "auto"
    and missing colors are null.

Body (`body`)
  - Blocks in document order: paragraphs, tables, and a `section_break` marker
    after each paragraph that ends a section. Block-level content controls,
    custom XML and tracked-insertion wrappers are transparent (their content is
    inlined); the control itself is recorded under `content_controls`.
  - Paragraph style: style id resolved to the lowercased style name; a missing or
    unknown w:pStyle means the default paragraph style.
  - Effective paragraph properties (ECMA-376 17.7.2 hierarchy): docDefaults, the
    table style's basedOn chain for paragraphs in table cells (w:tblStyle, else the
    default table style; conditional formats excluded), the paragraph style's
    basedOn chain (base first), the numbering level's w:pPr, then direct w:pPr, merged
    attribute by attribute. Defaults are filled: alignment "left" (start == left,
    end == right, both == justify), indents 0, spacing before/after 0 (or "auto"
    for auto-spacing), line 240 with rule "auto", keep/page-break flags false,
    widow control true (ECMA-376 17.3.1.44: applied unless turned off).
  - Numbering: w:numPr (direct, else from the style chain) resolved to the list
    definition fingerprint (SHA-256 prefix of every level's format, level text and
    start after w:lvlOverride; numStyleLink followed), level, format and level text.
    numId and abstractNumId values never appear. Symbol-font bullet U+F0B7 is
    normalized to U+2022.
  - Runs: visible text with effective formatting (docDefaults, table style chain,
    paragraph style chain, character style chain, direct w:rPr; last value wins,
    toggle-property XOR between style levels is not modeled). Formatting fields:
    bold, italic, underline (null for none), strike ("single"/"double"/null), size
    in points (default 10), color, font (ascii slot, theme fonts resolved),
    highlight (null for none), vert_align (null for baseline). Adjacent runs with
    identical effective formatting are merged; empty runs vanish. Runs consisting
    only of page/column breaks and object marks (U+FFFC) have null formatting:
    nothing visible carries it.
  - Visible text: w:t (leading/trailing XML whitespace dropped unless
    xml:space="preserve"), tab/ptab "\\t", line break "\\n", page break "\\f",
    column break "\\v", w:cr "\\n", non-breaking hyphen U+2011, soft hyphen U+00AD,
    w:sym as its code point, math text (run format {"math": true}), and U+FFFC for
    each drawing, VML picture or embedded object. Field codes are never visible;
    field results are. Tracked insertions are visible; deleted / moved-from text
    is not (it appears under `tracked_changes`). Layout caches such as
    w:lastRenderedPageBreak are ignored.
  - Tables: style name, grid column widths, rows (repeat-header flag) and cells.
    Each cell has its starting grid column (w:gridBefore and w:gridSpan applied),
    colspan, rowspan (w:vMerge continuations folded into the origin cell;
    continuation cell content is not shown by readers and is dropped), shading
    fill, effective borders and nested blocks. Effective cell borders: the table
    style chain's and the table's w:tblBorders (outer sides on the table edge,
    insideH/insideV elsewhere) overridden by the cell's w:tcBorders, so borders
    written on the table or on every cell are equal. Sides: start == left,
    end == right; "nil"/"none" removes a side; "auto" border color is 000000.
    Table-style conditional formatting and border-conflict resolution between
    adjacent cells are not modeled.

Other features
  - `sections`: page size, orientation (derived from width > height), margins
    (header/footer distance null when the section shows no header/footer),
    columns (spacing ignored for a single equal-width column), page numbering
    format/start and break type. The first section's break type is null because
    no break precedes it. w:titlePg is observable only through `headers_footers`.
  - `headers_footers`: per section, the header and footer a reader sees on first,
    odd and even pages (first only with w:titlePg, even only with
    w:evenAndOddHeaders; unset references inherit from the previous section).
    Each is {text, fields, images} without field results, or null when absent or
    blank.
  - `footnotes` / `endnotes`: note text in body reference order; note ids and
    unreferenced separator notes do not appear.
  - `comments`: author, text and anchored text, in order of first anchor or
    reference in the body; unreferenced comments are not observable and omitted.
  - `tracked_changes`: insertion / deletion / move_from / move_to with author and
    text, in body order. Adjacent changes of the same kind and author with nothing
    visible between them are merged. Dates and revision ids are ignored.
  - `fields`: normalized instructions in body order (field name uppercased,
    whitespace collapsed, \\* MERGEFORMAT dropped).
  - `bookmarks`: sorted set of names across body, headers/footers, notes, comments.
  - `hyperlinks`: body hyperlinks (w:hyperlink and HYPERLINK fields alike) as
    text, URL and anchor. URLs are percent-decoded with scheme and host
    lowercased; an empty http(s) path is "/".
  - `content_controls`: tag, alias, kind and text, in body order.
  - `images`: body pictures in order: placement (inline / anchor; VML absolute
    position == anchor), extent, and media SHA-256 (or external link).
  - `shapes`: body drawings and VML shapes without pictures: placement and extent.
  - `text_boxes`: body text-box content as blocks.
  - `embedded_objects`: OLE objects (ProgID, link flag, payload digest, preview
    image), ActiveX controls (class id, persistence, binary digest), charts (plot
    types, title, series names and value formulas) and altChunk payload digests.
    ZIP payloads hash as their sorted (member, SHA-256) list so re-zipping is
    equivalent; other payloads hash their bytes.
  - `styles`: sorted set of "type:name" for every defined style.
  - `custom_properties`: name -> {type, value}; types are string, number, boolean,
    date (UTC ISO 8601) or the variant type name; values compare by value.
  - `core_properties`: title, subject, creator, keywords (empty == null). Save-time
    metadata (dates, last modified by, revision, app.xml statistics) is excluded.
  - `custom_xml`: sorted custom XML data parts: SHA-256 of the C14N 2.0 canonical
    form (prefixes rewritten, text stripped), datastore item id and schema refs.
  - `glossary`: building-block entries (gallery, category, name, text), sorted.

Not modeled: field result recomputation, list counter values, toggle-property XOR
in style chains, table-style conditional formatting, theme color scheme mappings
(w:clrSchemeMapping), comment replies and dates, tracked formatting changes,
hyperlinks/images in notes and comments, SmartArt text, header/footer shapes,
document settings, fonts and embedded font data.

`compare` walks both models per feature. Dicts compare by key; lists are aligned
with difflib on canonical JSON and unmatched items are reported as missing (index
in `before`) or added (index in `after`); paired replacements recurse. Lengths
(keys in _LENGTH_KEYS, twips) differing by at most 10 twips (0.5 pt) are equal:
unit conversion (e.g. through 1/100 mm) rounds them and the difference is below
what a reader can see. Paths are RFC 6901 JSON pointers starting at the feature.
"""

from __future__ import annotations

import difflib
import hashlib
import io
import json
import posixpath
import re
import zipfile
from collections.abc import Iterator
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from urllib.parse import unquote
from xml.etree import ElementTree as ET

FEATURES: tuple[str, ...] = (
    "body",
    "sections",
    "headers_footers",
    "footnotes",
    "endnotes",
    "comments",
    "tracked_changes",
    "fields",
    "bookmarks",
    "hyperlinks",
    "content_controls",
    "images",
    "shapes",
    "text_boxes",
    "embedded_objects",
    "styles",
    "custom_properties",
    "core_properties",
    "custom_xml",
    "glossary",
)

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
V = "{urn:schemas-microsoft-com:vml}"
O = "{urn:schemas-microsoft-com:office:office}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
AX = "{http://schemas.microsoft.com/office/2006/activeX}"
DC = "{http://purl.org/dc/elements/1.1/}"
CP = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
CUSTOM = "{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}"
DS = "{http://schemas.openxmlformats.org/officeDocument/2006/customXml}"
PKG_RELS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

_STRICT_NAMESPACES = {
    "http://purl.oclc.org/ooxml/wordprocessingml/main": W[1:-1],
    "http://purl.oclc.org/ooxml/officeDocument/relationships": R[1:-1],
    "http://purl.oclc.org/ooxml/drawingml/main": A[1:-1],
    "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing": WP[1:-1],
    "http://purl.oclc.org/ooxml/drawingml/picture": (
        "http://schemas.openxmlformats.org/drawingml/2006/picture"
    ),
    "http://purl.oclc.org/ooxml/drawingml/chart": C[1:-1],
    "http://purl.oclc.org/ooxml/officeDocument/math": M[1:-1],
    "http://purl.oclc.org/ooxml/officeDocument/customProperties": CUSTOM[1:-1],
    "http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes": (
        "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
    ),
    "http://purl.oclc.org/ooxml/officeDocument/customXml": DS[1:-1],
}

_TRUE = frozenset({"1", "true", "on"})
_ALIGNMENT = {
    "left": "left",
    "start": "left",
    "right": "right",
    "end": "right",
    "center": "center",
    "both": "justify",
}
_PARAGRAPH_FLAGS = (
    ("keepNext", "keep_next"),
    ("keepLines", "keep_lines"),
    ("pageBreakBefore", "page_break_before"),
    ("widowControl", "widow_control"),
)
_BORDER_SIDES = {
    "top": "top",
    "left": "left",
    "start": "left",
    "bottom": "bottom",
    "right": "right",
    "end": "right",
    "insideH": "inside_h",
    "insideV": "inside_v",
}
_THEME_COLORS = {
    "dark1": "dk1",
    "light1": "lt1",
    "dark2": "dk2",
    "light2": "lt2",
    "text1": "dk1",
    "background1": "lt1",
    "text2": "dk2",
    "background2": "lt2",
    "hyperlink": "hlink",
    "followedHyperlink": "folHlink",
}
_SDT_KINDS = (
    "text",
    "richText",
    "comboBox",
    "dropDownList",
    "date",
    "picture",
    "checkbox",
    "group",
    "citation",
    "bibliography",
    "equation",
    "docPartObj",
    "docPartList",
)
_SYMBOL_GLYPHS = {"\uf0b7": "\u2022"}
_FIELD_TOKEN = re.compile(r'"[^"]*"|\S+')
_MEASURE = re.compile(r"^\s*(-?\d+(?:\.\d+)?|-?\.\d+)\s*(mm|cm|in|pt|pc|pi|px)?\s*$")
_TWIPS_PER = {
    None: 1.0,
    "pt": 20.0,
    "pc": 240.0,
    "pi": 240.0,
    "in": 1440.0,
    "cm": 1440.0 / 2.54,
    "mm": 144.0 / 2.54,
    "px": 15.0,
}
_VML_SHAPES = frozenset(
    V + local
    for local in (
        "shape",
        "rect",
        "roundrect",
        "oval",
        "line",
        "polyline",
        "arc",
        "curve",
        "group",
        "image",
    )
)
_MATH_FORMAT = {"math": True}
_OBJECT_MARK = "\ufffc"
# Model keys holding lengths in twips, and the difference treated as equal (0.5 pt).
_LENGTH_KEYS = frozenset(
    {
        "grid",
        "widths",
        "width",
        "height",
        "left",
        "right",
        "first_line",
        "top",
        "bottom",
        "header",
        "footer",
        "gutter",
        "before",
        "after",
        "space",
    }
)
_LENGTH_TOLERANCE = 10


# ---------------------------------------------------------------------------
# Small value helpers
# ---------------------------------------------------------------------------


def _val(element: ET.Element | None, attr: str = "val") -> str | None:
    return element.get(W + attr) if element is not None else None


def _onoff(element: ET.Element | None) -> bool | None:
    if element is None:
        return None
    value = element.get(W + "val")
    return True if value is None else value.strip().lower() in _TRUE


def _onoff_attr(value: str | None) -> bool | None:
    return None if value is None else value.strip().lower() in _TRUE


def _twips(value: str | None) -> int | None:
    if value is None:
        return None
    match = _MEASURE.match(value)
    if match is None:
        return None
    return round(float(match.group(1)) * _TWIPS_PER[match.group(2)])


def _number(value: str | None) -> float | None:
    if value is None:
        return None
    try:
        return float(value)
    except ValueError:
        return None


def _emu_to_twips(value: str | None) -> int | None:
    number = _number(value)
    return None if number is None else round(number / 635)


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _payload_digest(data: bytes) -> str:
    """ZIP payloads hash as their sorted member digests; anything else by bytes."""
    if data[:4] == b"PK\x03\x04":
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                members = sorted(
                    (info.filename, _sha256(archive.read(info)))
                    for info in archive.infolist()
                    if not info.is_dir()
                )
        except (zipfile.BadZipFile, OSError, ValueError):
            return _sha256(data)
        return "zip:" + _sha256(json.dumps(members).encode())
    return _sha256(data)


def _chosen(alternate: ET.Element) -> list[ET.Element]:
    choice = alternate.find(MC + "Choice")
    if choice is not None:
        return list(choice)
    fallback = alternate.find(MC + "Fallback")
    return list(fallback) if fallback is not None else []


def _children(element: ET.Element) -> Iterator[ET.Element]:
    """Children with mc:AlternateContent replaced by its chosen branch."""
    for child in element:
        if child.tag == MC + "AlternateContent":
            yield from (
                grand
                for chosen in _chosen(child)
                for grand in _expand_alternate(chosen)
            )
        else:
            yield child


def _expand_alternate(element: ET.Element) -> Iterator[ET.Element]:
    if element.tag == MC + "AlternateContent":
        for chosen in _chosen(element):
            yield from _expand_alternate(chosen)
    else:
        yield element


def _descendants(
    element: ET.Element, stop: frozenset[str] = frozenset()
) -> Iterator[ET.Element]:
    """Descendants in document order following chosen branches; `stop` tags are
    yielded but not entered."""
    for child in _children(element):
        yield child
        if child.tag not in stop:
            yield from _descendants(child, stop)


def _normalize_instruction(value: str) -> str:
    tokens = _FIELD_TOKEN.findall(value)
    if not tokens:
        return ""
    kept: list[str] = [tokens[0].upper()]
    index = 1
    while index < len(tokens):
        if (
            tokens[index] == "\\*"
            and index + 1 < len(tokens)
            and tokens[index + 1].upper() == "MERGEFORMAT"
        ):
            index += 2
            continue
        kept.append(tokens[index])
        index += 1
    return " ".join(kept)


def _hyperlink_field(instruction: str) -> tuple[str | None, str | None] | None:
    tokens = _FIELD_TOKEN.findall(instruction)
    if not tokens or tokens[0].upper() != "HYPERLINK":
        return None
    url: str | None = None
    anchor: str | None = None
    index = 1
    while index < len(tokens):
        token = tokens[index]
        if token.startswith("\\"):
            argument = None
            if index + 1 < len(tokens) and not tokens[index + 1].startswith("\\"):
                argument = tokens[index + 1].strip('"')
                index += 1
            if token.lower() == "\\l":
                anchor = argument
        elif url is None:
            url = token.strip('"')
        index += 1
    return (_normalize_url(url) if url else None), anchor


def _normalize_url(url: str) -> str:
    """Percent-decode; lowercase scheme and host; an empty http(s) path is "/"."""
    url = unquote(url)
    match = re.match(r"^([A-Za-z][A-Za-z0-9+.-]*)://([^/?#]*)(.*)$", url, re.DOTALL)
    if match is None:
        return url
    scheme, host, rest = match.group(1).lower(), match.group(2).lower(), match.group(3)
    if scheme in ("http", "https") and not rest.startswith("/"):
        rest = "/" + rest
    return f"{scheme}://{host}{rest}"


def _css_lengths(style: str | None) -> dict[str, str]:
    result: dict[str, str] = {}
    for declaration in (style or "").split(";"):
        name, _, value = declaration.partition(":")
        if name.strip():
            result[name.strip().lower()] = value.strip()
    return result


def _css_twips(value: str | None) -> int | None:
    if not value:
        return None
    match = _MEASURE.match(value)
    if match is None:
        return None
    unit = match.group(2) or "px"
    return round(float(match.group(1)) * _TWIPS_PER[unit])


def _tint_shade(base: str, tint: str | None, shade: str | None) -> str:
    channels = [int(base[index : index + 2], 16) for index in (0, 2, 4)]
    if tint:
        factor = int(tint, 16) / 255
        channels = [round(value * factor + 255 * (1 - factor)) for value in channels]
    if shade:
        factor = int(shade, 16) / 255
        channels = [round(value * factor) for value in channels]
    return "".join(f"{value:02X}" for value in channels)


# ---------------------------------------------------------------------------
# Package access
# ---------------------------------------------------------------------------


def _destrict(root: ET.Element) -> None:
    def mapped(name: str) -> str:
        if name.startswith("{http://purl.oclc.org/"):
            uri, _, local = name[1:].partition("}")
            target = _STRICT_NAMESPACES.get(uri)
            if target is not None:
                return "{" + target + "}" + local
        return name

    for element in root.iter():
        if isinstance(element.tag, str):
            element.tag = mapped(element.tag)
        if any(key.startswith("{http://purl.oclc.org/") for key in element.attrib):
            element.attrib = {
                mapped(key): value for key, value in element.attrib.items()
            }


class _Package:
    def __init__(self, path: Path) -> None:
        try:
            with zipfile.ZipFile(path) as archive:
                self.data = {
                    info.filename.lstrip("/").lower(): archive.read(info)
                    for info in archive.infolist()
                    if not info.is_dir()
                }
        except (OSError, zipfile.BadZipFile, zipfile.LargeZipFile, ValueError) as exc:
            raise ValueError(f"unreadable package {path}: {exc}") from exc
        except Exception as exc:  # zlib.error, NotImplementedError (compression)
            raise ValueError(f"unreadable package {path}: {exc!r}") from exc
        self._xml: dict[str, ET.Element] = {}
        self._rels: dict[str, dict[str, tuple[str, str, bool]]] = {}

    def xml(self, name: str) -> ET.Element:
        if name not in self._xml:
            data = self.data.get(name)
            if data is None:
                raise ValueError(f"missing part {name}")
            try:
                root = ET.fromstring(data)
            except ET.ParseError as exc:
                raise ValueError(f"malformed XML in {name}: {exc}") from exc
            if b"purl.oclc.org/ooxml" in data:
                _destrict(root)
            self._xml[name] = root
        return self._xml[name]

    def rels(self, source: str) -> dict[str, tuple[str, str, bool]]:
        """Relationship id -> (type suffix, target, external) for a source part."""
        if source not in self._rels:
            directory, base = posixpath.split(source)
            key = posixpath.join(directory, "_rels", base + ".rels")
            result: dict[str, tuple[str, str, bool]] = {}
            if key in self.data:
                for item in self.xml(key):
                    if _local(item.tag) != "Relationship":
                        continue
                    kind = item.get("Type", "").rstrip("/").rsplit("/", 1)[-1]
                    target = item.get("Target", "")
                    external = item.get("TargetMode", "") == "External"
                    if not external:
                        target = self.resolve(source, target)
                    result[item.get("Id", "")] = (kind, target, external)
            self._rels[source] = result
        return self._rels[source]

    @staticmethod
    def resolve(source: str, target: str) -> str:
        target = unquote(target.split("#", 1)[0])
        if target.startswith("/"):
            resolved = target.lstrip("/")
        else:
            resolved = posixpath.normpath(
                posixpath.join(posixpath.dirname(source), target)
            )
        return resolved.lower()

    def related(self, source: str, kind: str) -> list[str]:
        return [
            target
            for rel_kind, target, external in self.rels(source).values()
            if rel_kind == kind and not external and target in self.data
        ]

    def first_related(self, source: str, kind: str) -> str | None:
        found = sorted(self.related(source, kind))
        return found[0] if found else None


# ---------------------------------------------------------------------------
# Styles, numbering, theme
# ---------------------------------------------------------------------------


class _Styles:
    def __init__(self, root: ET.Element | None) -> None:
        self.by_id: dict[str, ET.Element] = {}
        self.defaults: dict[str, str] = {}
        self.doc_rpr: ET.Element | None = None
        self.doc_ppr: ET.Element | None = None
        self._chains: dict[str | None, list[ET.Element]] = {}
        if root is None:
            return
        defaults = root.find(W + "docDefaults")
        if defaults is not None:
            self.doc_rpr = defaults.find(f"{W}rPrDefault/{W}rPr")
            self.doc_ppr = defaults.find(f"{W}pPrDefault/{W}pPr")
        for style in root.findall(W + "style"):
            identifier = style.get(W + "styleId")
            if identifier is None:
                continue
            self.by_id.setdefault(identifier, style)
            if _onoff_attr(style.get(W + "default")):
                self.defaults.setdefault(style.get(W + "type", "paragraph"), identifier)

    def name(self, identifier: str | None) -> str | None:
        style = self.by_id.get(identifier or "")
        if style is None:
            return None
        name = _val(style.find(W + "name"))
        return (name or identifier or "").lower()

    def chain(self, identifier: str | None) -> list[ET.Element]:
        """Style and its basedOn ancestors, base first (cycle-safe, memoized)."""
        if identifier in self._chains:
            return self._chains[identifier]
        key = identifier
        chain: list[ET.Element] = []
        seen: set[str] = set()
        while identifier and identifier not in seen and identifier in self.by_id:
            seen.add(identifier)
            style = self.by_id[identifier]
            chain.append(style)
            identifier = _val(style.find(W + "basedOn"))
        chain.reverse()
        self._chains[key] = chain
        return chain

    def paragraph_style(self, ppr: ET.Element | None) -> str | None:
        identifier = _val(ppr.find(W + "pStyle")) if ppr is not None else None
        if identifier in self.by_id:
            return identifier
        return self.defaults.get("paragraph")

    def defined(self) -> list[str]:
        return sorted(
            {
                f"{style.get(W + 'type', 'paragraph')}:{self.name(identifier)}"
                for identifier, style in self.by_id.items()
            }
        )


class _Theme:
    def __init__(self, root: ET.Element | None) -> None:
        self.colors: dict[str, str] = {}
        self.fonts: dict[str, str] = {}
        if root is None:
            return
        scheme = root.find(f"{A}themeElements/{A}clrScheme")
        if scheme is not None:
            for entry in scheme:
                for color in entry:
                    value = color.get("val") if _local(color.tag) == "srgbClr" else None
                    if _local(color.tag) == "sysClr":
                        value = color.get("lastClr")
                    if value:
                        self.colors[_local(entry.tag)] = value.upper()
        fonts = root.find(f"{A}themeElements/{A}fontScheme")
        if fonts is not None:
            for kind in ("major", "minor"):
                latin = fonts.find(f"{A}{kind}Font/{A}latin")
                if latin is not None and latin.get("typeface"):
                    self.fonts[kind] = latin.get("typeface", "")

    def color(self, spec: tuple[str | None, ...] | None) -> str | None:
        if spec is None:
            return None
        value, theme, tint, shade = spec
        if value and value.lower() != "auto":
            return value.upper()
        if theme:
            base = self.colors.get(_THEME_COLORS.get(theme, theme))
            if base:
                try:
                    return _tint_shade(base, tint, shade)
                except ValueError:
                    return base
        return None

    def font(self, spec: tuple[str | None, str | None] | None) -> str | None:
        if spec is None:
            return None
        name, theme = spec
        if theme:
            resolved = self.fonts.get("major" if theme.startswith("major") else "minor")
            if resolved:
                return resolved
        return name


def _apply_ppr(acc: dict[str, Any], ppr: ET.Element | None) -> None:
    if ppr is None:
        return
    jc = _val(ppr.find(W + "jc"))
    if jc is not None:
        acc["alignment"] = _ALIGNMENT.get(jc, jc)
    ind = ppr.find(W + "ind")
    if ind is not None:
        for attr, key in (
            ("left", "left"),
            ("start", "left"),
            ("right", "right"),
            ("end", "right"),
        ):
            value = _twips(ind.get(W + attr))
            if value is not None:
                acc[key] = value
        hanging = _twips(ind.get(W + "hanging"))
        first = _twips(ind.get(W + "firstLine"))
        if hanging is not None:
            acc["first_line"] = -hanging
        elif first is not None:
            acc["first_line"] = first
    spacing = ppr.find(W + "spacing")
    if spacing is not None:
        for attr in ("before", "after", "line"):
            value = _twips(spacing.get(W + attr))
            if value is not None:
                acc[attr] = value
        rule = spacing.get(W + "lineRule")
        if rule is not None:
            acc["line_rule"] = rule
        for attr in ("beforeAutospacing", "afterAutospacing"):
            flag = _onoff_attr(spacing.get(W + attr))
            if flag is not None:
                acc[attr] = flag
    for tag, key in _PARAGRAPH_FLAGS:
        flag = _onoff(ppr.find(W + tag))
        if flag is not None:
            acc[key] = flag


def _apply_rpr(acc: dict[str, Any], rpr: ET.Element | None) -> None:
    if rpr is None:
        return
    for tag in ("b", "i", "strike", "dstrike"):
        flag = _onoff(rpr.find(W + tag))
        if flag is not None:
            acc[tag] = flag
    underline = rpr.find(W + "u")
    if underline is not None:
        acc["u"] = underline.get(W + "val", "single")
    size = _number(_val(rpr.find(W + "sz")))
    if size is not None:
        acc["sz"] = size
    color = rpr.find(W + "color")
    if color is not None:
        acc["color"] = (
            color.get(W + "val"),
            color.get(W + "themeColor"),
            color.get(W + "themeTint"),
            color.get(W + "themeShade"),
        )
    fonts = rpr.find(W + "rFonts")
    if fonts is not None and (fonts.get(W + "ascii") or fonts.get(W + "asciiTheme")):
        acc["font"] = (fonts.get(W + "ascii"), fonts.get(W + "asciiTheme"))
    for tag in ("highlight", "vertAlign"):
        value = _val(rpr.find(W + tag))
        if value is not None:
            acc[tag] = value


def _level_row(level: ET.Element) -> list[Any]:
    text = _val(level.find(W + "lvlText"))
    if text is not None:
        text = "".join(_SYMBOL_GLYPHS.get(character, character) for character in text)
    start = _val(level.find(W + "start"))
    return [
        _val(level.find(W + "numFmt")) or "decimal",
        text,
        int(start) if start and start.lstrip("-").isdigit() else 0,
    ]


class _Numbering:
    def __init__(self, root: ET.Element | None, styles: _Styles) -> None:
        self.styles = styles
        self.abstract: dict[str, ET.Element] = {}
        self.nums: dict[str, ET.Element] = {}
        self._cache: dict[
            str, tuple[str, dict[int, list[Any]], dict[int, ET.Element]] | None
        ] = {}
        if root is None:
            return
        for item in root.findall(W + "abstractNum"):
            self.abstract.setdefault(item.get(W + "abstractNumId", ""), item)
        for item in root.findall(W + "num"):
            self.nums.setdefault(item.get(W + "numId", ""), item)

    def _abstract_for(self, num_id: str, depth: int = 0) -> ET.Element | None:
        num = self.nums.get(num_id)
        if num is None:
            return None
        abstract = self.abstract.get(_val(num.find(W + "abstractNumId")) or "")
        if abstract is not None and depth < 4:
            link = _val(abstract.find(W + "numStyleLink"))
            style = self.styles.by_id.get(link or "")
            linked_id = (
                _val(style.find(f"{W}pPr/{W}numPr/{W}numId"))
                if style is not None
                else None
            )
            if linked_id and linked_id != num_id:
                linked = self._abstract_for(linked_id, depth + 1)
                if linked is not None:
                    return linked
        return abstract

    def _definition(
        self, num_id: str
    ) -> tuple[str, dict[int, list[Any]], dict[int, ET.Element]] | None:
        if num_id not in self._cache:
            abstract = self._abstract_for(num_id)
            if abstract is None:
                self._cache[num_id] = None
                return None
            levels: dict[int, ET.Element] = {}
            for level in abstract.findall(W + "lvl"):
                levels[int(level.get(W + "ilvl", "0"))] = level
            rows = {index: _level_row(level) for index, level in levels.items()}
            for override in self.nums[num_id].findall(W + "lvlOverride"):
                index = int(override.get(W + "ilvl", "0"))
                level = override.find(W + "lvl")
                if level is not None:
                    levels[index] = level
                    rows[index] = _level_row(level)
                start = _val(override.find(W + "startOverride"))
                if start is not None and start.lstrip("-").isdigit() and index in rows:
                    rows[index][2] = int(start)
            fingerprint = _sha256(json.dumps(sorted(rows.items())).encode())[:16]
            self._cache[num_id] = (fingerprint, rows, levels)
        return self._cache[num_id]

    def record(self, num_id: str, level: int) -> dict[str, Any] | None:
        definition = self._definition(num_id)
        if definition is None:
            return None
        fingerprint, rows, _ = definition
        row = rows.get(level, [None, None, None])
        return {
            "definition": fingerprint,
            "level": level,
            "format": row[0],
            "text": row[1],
        }

    def level_ppr(self, num_id: str, level: int) -> ET.Element | None:
        definition = self._definition(num_id)
        if definition is None or level not in definition[2]:
            return None
        return definition[2][level].find(W + "pPr")


# ---------------------------------------------------------------------------
# Document
# ---------------------------------------------------------------------------


class _Document:
    def __init__(self, path: Path) -> None:
        self.pkg = _Package(path)
        main = self.pkg.first_related("", "officeDocument")
        if main is None:
            raise ValueError(f"{path}: no main document part")
        self.main = main
        self.root = self.pkg.xml(main)
        body = self.root.find(W + "body")
        if body is None:
            raise ValueError(f"{path}: main document has no w:body")
        self.body = body
        styles_part = self.pkg.first_related(main, "styles")
        self.styles = _Styles(self.pkg.xml(styles_part) if styles_part else None)
        theme_part = self.pkg.first_related(main, "theme")
        self.theme = _Theme(self.pkg.xml(theme_part) if theme_part else None)
        numbering_part = self.pkg.first_related(main, "numbering")
        self.numbering = _Numbering(
            self.pkg.xml(numbering_part) if numbering_part else None, self.styles
        )
        settings_part = self.pkg.first_related(main, "settings")
        settings = self.pkg.xml(settings_part) if settings_part else None
        self.even_and_odd = bool(
            settings is not None and _onoff(settings.find(W + "evenAndOddHeaders"))
        )
        self._para_base: dict[
            tuple[str | None, str | None], tuple[dict[str, Any], dict[str, Any]]
        ] = {}
        self.bookmarks: set[str] = set()

    # -- formatting -------------------------------------------------------

    def _base(
        self, style_id: str | None, table_style: str | None
    ) -> tuple[dict[str, Any], dict[str, Any]]:
        """docDefaults, then the table style chain, then the paragraph style chain."""
        key = (style_id, table_style)
        if key not in self._para_base:
            pacc: dict[str, Any] = {}
            racc: dict[str, Any] = {}
            _apply_ppr(pacc, self.styles.doc_ppr)
            _apply_rpr(racc, self.styles.doc_rpr)
            for style in self.styles.chain(table_style) + self.styles.chain(style_id):
                _apply_ppr(pacc, style.find(W + "pPr"))
                _apply_rpr(racc, style.find(W + "rPr"))
            self._para_base[key] = (pacc, racc)
        return self._para_base[key]

    def run_format(
        self,
        style_id: str | None,
        rpr: ET.Element | None,
        table_style: str | None = None,
    ) -> dict[str, Any]:
        acc = dict(self._base(style_id, table_style)[1])
        if rpr is not None:
            for style in self.styles.chain(_val(rpr.find(W + "rStyle"))):
                _apply_rpr(acc, style.find(W + "rPr"))
            _apply_rpr(acc, rpr)
        underline = acc.get("u")
        highlight = acc.get("highlight")
        vert_align = acc.get("vertAlign")
        return {
            "bold": bool(acc.get("b")),
            "italic": bool(acc.get("i")),
            "underline": None if underline in (None, "none") else underline,
            "strike": "double"
            if acc.get("dstrike")
            else "single"
            if acc.get("strike")
            else None,
            "size": acc.get("sz", 20.0) / 2,
            "color": self.theme.color(acc.get("color")),
            "font": self.theme.font(acc.get("font")),
            "highlight": None if highlight in (None, "none") else highlight,
            "vert_align": None if vert_align in (None, "baseline") else vert_align,
        }

    def _numbering_ref(
        self, style_id: str | None, ppr: ET.Element | None
    ) -> tuple[str, int] | None:
        num_pr = ppr.find(W + "numPr") if ppr is not None else None
        num_id = _val(num_pr.find(W + "numId")) if num_pr is not None else None
        level = _val(num_pr.find(W + "ilvl")) if num_pr is not None else None
        if num_id is None:
            for style in reversed(self.styles.chain(style_id)):
                style_num = style.find(f"{W}pPr/{W}numPr")
                if style_num is None:
                    continue
                num_id = num_id or _val(style_num.find(W + "numId"))
                level = level or _val(style_num.find(W + "ilvl"))
                if num_id is not None:
                    break
        if num_id is None or num_id == "0":
            return None
        try:
            return num_id, int(level or 0)
        except ValueError:
            return num_id, 0

    def paragraph_props(
        self,
        style_id: str | None,
        ppr: ET.Element | None,
        table_style: str | None = None,
    ) -> tuple[dict[str, Any], dict[str, Any] | None]:
        acc = dict(self._base(style_id, table_style)[0])
        reference = self._numbering_ref(style_id, ppr)
        numbering = None
        if reference is not None:
            numbering = self.numbering.record(*reference)
            _apply_ppr(acc, self.numbering.level_ppr(*reference))
        _apply_ppr(acc, ppr)
        props = {
            "alignment": acc.get("alignment", "left"),
            "indent": {
                "left": acc.get("left", 0),
                "right": acc.get("right", 0),
                "first_line": acc.get("first_line", 0),
            },
            "spacing": {
                "before": "auto"
                if acc.get("beforeAutospacing")
                else acc.get("before", 0),
                "after": "auto" if acc.get("afterAutospacing") else acc.get("after", 0),
                "line": acc.get("line", 240),
                "line_rule": acc.get("line_rule", "auto"),
            },
        }
        for _, key in _PARAGRAPH_FLAGS:
            props[key] = bool(acc.get(key, key == "widow_control"))
        return props, numbering

    def color(self, element: ET.Element | None, attr: str = "color") -> str | None:
        if element is None:
            return None
        theme_attr = "themeFill" if attr == "fill" else "themeColor"
        prefix = "themeFill" if attr == "fill" else "theme"
        return self.theme.color(
            (
                element.get(W + attr),
                element.get(W + theme_attr),
                element.get(W + prefix + "Tint"),
                element.get(W + prefix + "Shade"),
            )
        )

    # -- media ------------------------------------------------------------

    def target_digest(
        self, part: str, rel_id: str | None, *, payload: bool = False
    ) -> str | None:
        if not rel_id:
            return None
        rel = self.pkg.rels(part).get(rel_id)
        if rel is None:
            return "unresolved"
        _, target, external = rel
        if external:
            return "link:" + unquote(target)
        data = self.pkg.data.get(target)
        if data is None:
            return "missing"
        return _payload_digest(data) if payload else _sha256(data)


# ---------------------------------------------------------------------------
# Story walker
# ---------------------------------------------------------------------------


class _Story:
    """Walks one story (body, header, note, comment, text box) of a part."""

    def __init__(
        self, doc: _Document, part: str, *, field_results: bool = True
    ) -> None:
        self.doc = doc
        self.part = part
        self.field_results = field_results
        self.fields: list[str | None] = []
        self.hyperlinks: list[dict[str, Any] | None] = []
        self.content_controls: list[dict[str, Any]] = []
        self.images: list[dict[str, Any]] = []
        self.shapes: list[dict[str, Any]] = []
        self.objects: list[dict[str, Any]] = []
        self.text_boxes: list[list[dict[str, Any]]] = []
        self.changes: list[dict[str, Any]] = []
        self.note_refs: list[tuple[str, str]] = []
        self.comment_order: list[str] = []
        self.comment_anchor: dict[str, list[str]] = {}
        self.sections: list[ET.Element] = []
        self._field_stack: list[dict[str, Any]] = []
        self._simple_fields = 0
        self._captures: list[list[str]] = []
        self._deleted: list[list[str]] = []
        self._counter = 0
        self._last_change_end = -1
        self._runs: list[dict[str, Any]] | None = None
        self._style: str | None = None
        self._table_style: str | None = None

    # -- emission ---------------------------------------------------------

    def _hidden(self) -> bool:
        if any(field["phase"] == "code" for field in self._field_stack):
            return True
        return not self.field_results and bool(self._field_stack or self._simple_fields)

    def _emit(self, text: str, fmt: dict[str, Any] | None) -> None:
        if not text:
            return
        if self._deleted:
            for buffer in self._deleted:
                buffer.append(text)
            self._counter += 1
            return
        if self._hidden():
            return
        runs = self._runs
        if runs is not None:
            if not text.strip("\f\v" + _OBJECT_MARK):
                fmt = None  # breaks and object anchors carry no visible formatting
            if runs and runs[-1]["format"] == fmt:
                runs[-1]["text"] += text
            else:
                runs.append({"text": text, "format": fmt})
        for buffer in self._captures:
            buffer.append(text)
        self._counter += 1

    def _release(self, buffer: list[str]) -> str:
        for index in range(len(self._captures) - 1, -1, -1):
            if self._captures[index] is buffer:
                del self._captures[index]
                break
        return "".join(buffer).strip("\n")

    # -- blocks -----------------------------------------------------------

    def blocks(self, container: ET.Element) -> list[dict[str, Any]]:
        out: list[dict[str, Any]] = []
        for child in _children(container):
            self._block(child, out)
        return out

    def _block(self, element: ET.Element, out: list[dict[str, Any]]) -> None:
        tag = element.tag
        if tag == W + "p":
            out.append(self._paragraph(element))
            section = element.find(f"{W}pPr/{W}sectPr")
            if section is not None:
                self.sections.append(section)
                out.append({"type": "section_break"})
        elif tag == W + "tbl":
            out.append(self._table(element))
        elif tag == W + "sdt":
            self._content_control(
                element, lambda content: self._blocks_into(content, out)
            )
        elif tag in (W + "customXml", W + "ins", W + "moveTo", W + "smartTag"):
            self._blocks_into(element, out)
        elif tag in (W + "del", W + "moveFrom"):
            self._change(
                "deletion" if tag == W + "del" else "move_from",
                element,
                lambda: self._blocks_into(element, out),
            )
        elif tag == W + "sectPr":
            self.sections.append(element)
        elif tag == W + "altChunk":
            self.objects.append(
                {
                    "kind": "altChunk",
                    "data": self.doc.target_digest(
                        self.part, element.get(R + "id"), payload=True
                    ),
                }
            )
        else:
            self._markup(element)

    def _blocks_into(self, container: ET.Element, out: list[dict[str, Any]]) -> None:
        for child in _children(container):
            self._block(child, out)

    def _paragraph(self, paragraph: ET.Element) -> dict[str, Any]:
        ppr = paragraph.find(W + "pPr")
        style_id = self.doc.styles.paragraph_style(ppr)
        props, numbering = self.doc.paragraph_props(style_id, ppr, self._table_style)
        runs: list[dict[str, Any]] = []
        saved = (self._runs, self._style)
        self._runs, self._style = runs, style_id
        for child in _children(paragraph):
            self._inline(child)
        self._runs, self._style = saved
        for buffer in self._captures:
            buffer.append("\n")
        return {
            "type": "paragraph",
            "style": self.doc.styles.name(style_id),
            "props": props,
            "numbering": numbering,
            "runs": runs,
        }

    def _table(self, table: ET.Element) -> dict[str, Any]:
        tbl_pr = table.find(W + "tblPr")
        style = _val(tbl_pr.find(W + "tblStyle")) if tbl_pr is not None else None
        grid = [
            _twips(column.get(W + "w"))
            for column in table.findall(f"{W}tblGrid/{W}gridCol")
        ]
        rows: list[dict[str, Any]] = []
        origins: dict[int, dict[str, Any]] = {}
        for row in self._table_items(table, W + "tr"):
            tr_pr = row.find(W + "trPr")
            before = _val(tr_pr.find(W + "gridBefore")) if tr_pr is not None else None
            column = int(before) if before and before.isdigit() else 0
            cells: list[dict[str, Any]] = []
            next_origins: dict[int, dict[str, Any]] = {}
            for cell in self._table_items(row, W + "tc"):
                tc_pr = cell.find(W + "tcPr")
                span_value = (
                    _val(tc_pr.find(W + "gridSpan")) if tc_pr is not None else None
                )
                span = int(span_value) if span_value and span_value.isdigit() else 1
                merge = tc_pr.find(W + "vMerge") if tc_pr is not None else None
                continuing = (
                    merge is not None and merge.get(W + "val", "continue") != "restart"
                )
                origin = origins.get(column)
                if continuing and origin is not None and origin["colspan"] == span:
                    origin["rowspan"] += 1
                    next_origins[column] = origin
                    column += span
                    continue
                shading = tc_pr.find(W + "shd") if tc_pr is not None else None
                fill = None
                if shading is not None:
                    fill = (
                        self.doc.color(shading)
                        if shading.get(W + "val") == "solid"
                        else self.doc.color(shading, "fill")
                    )
                record = {
                    "column": column,
                    "colspan": span,
                    "rowspan": 1,
                    "shading": fill,
                    "borders": self._sides(
                        tc_pr.find(W + "tcBorders") if tc_pr is not None else None
                    ),
                    "blocks": self._cell_blocks(cell, style),
                }
                if merge is not None:
                    next_origins[column] = record
                cells.append(record)
                column += span
            origins = next_origins
            rows.append(
                {
                    "repeat_header": bool(
                        _onoff(tr_pr.find(W + "tblHeader"))
                        if tr_pr is not None
                        else False
                    ),
                    "cells": cells,
                }
            )
        table_borders: dict[str, Any] = {}
        for definition in self.doc.styles.chain(style):
            table_borders.update(
                self._sides(definition.find(f"{W}tblPr/{W}tblBorders"))
            )
        table_borders.update(
            self._sides(tbl_pr.find(W + "tblBorders") if tbl_pr is not None else None)
        )
        columns = max(
            [len(grid)]
            + [
                cell["column"] + cell["colspan"]
                for row in rows
                for cell in row["cells"]
            ]
        )
        for index, row in enumerate(rows):
            for cell in row["cells"]:
                edges = {
                    "top": table_borders.get("top" if index == 0 else "inside_h"),
                    "bottom": table_borders.get(
                        "bottom" if index + cell["rowspan"] >= len(rows) else "inside_h"
                    ),
                    "left": table_borders.get(
                        "left" if cell["column"] == 0 else "inside_v"
                    ),
                    "right": table_borders.get(
                        "right"
                        if cell["column"] + cell["colspan"] >= columns
                        else "inside_v"
                    ),
                }
                for side, value in cell["borders"].items():
                    if side in edges:
                        edges[side] = value
                cell["borders"] = {
                    side: value
                    for side, value in sorted(edges.items())
                    if value is not None
                }
        return {
            "type": "table",
            "style": self.doc.styles.name(style) if style else None,
            "grid": grid,
            "rows": rows,
        }

    def _table_items(self, container: ET.Element, tag: str) -> Iterator[ET.Element]:
        """Rows of a table or cells of a row, seen through sdt/customXml wrappers."""
        for child in _children(container):
            if child.tag == tag:
                yield child
            elif child.tag == W + "sdt":
                content = child.find(W + "sdtContent")
                if content is not None:
                    self._content_control(child, lambda _content: None)
                    yield from self._table_items(content, tag)
            elif child.tag in (W + "customXml", W + "ins", W + "moveTo"):
                yield from self._table_items(child, tag)
            else:
                self._markup(child)

    def _sides(self, borders: ET.Element | None) -> dict[str, Any]:
        """Explicit border sides; "nil"/"none" map to None so they override."""
        result: dict[str, Any] = {}
        if borders is None:
            return result
        for side in borders:
            name = _BORDER_SIDES.get(_local(side.tag))
            style = side.get(W + "val")
            if name is None or style is None:
                continue
            if style in ("nil", "none"):
                result[name] = None
                continue
            size = _number(side.get(W + "sz"))
            result[name] = {
                "style": style,
                "size": int(size) if size is not None else None,
                "color": self.doc.color(side) or "000000",
            }
        return result

    # -- inline -----------------------------------------------------------

    def _inline(self, element: ET.Element) -> None:
        tag = element.tag
        if tag == W + "r":
            self._run(element)
        elif tag == W + "hyperlink":
            url = None
            rel_id = element.get(R + "id")
            if rel_id:
                rel = self.doc.pkg.rels(self.part).get(rel_id)
                url = _normalize_url(rel[1]) if rel is not None else "unresolved"
            record: dict[str, Any] = {
                "text": "",
                "url": url,
                "anchor": element.get(W + "anchor"),
            }
            self.hyperlinks.append(record)
            buffer: list[str] = []
            self._captures.append(buffer)
            for child in _children(element):
                self._inline(child)
            record["text"] = self._release(buffer)
        elif tag in (W + "ins", W + "moveTo"):
            self._change(
                "insertion" if tag == W + "ins" else "move_to",
                element,
                lambda: self._inline_children(element),
            )
        elif tag in (W + "del", W + "moveFrom"):
            self._change(
                "deletion" if tag == W + "del" else "move_from",
                element,
                lambda: self._inline_children(element),
            )
        elif tag == W + "fldSimple":
            self._simple_field(element)
        elif tag == W + "sdt":
            self._content_control(element, self._inline_children)
        elif tag in (W + "smartTag", W + "customXml", W + "dir", W + "bdo"):
            self._inline_children(element)
        elif tag in (M + "oMathPara", M + "oMath"):
            text = "".join(node.text or "" for node in element.iter(M + "t"))
            self._emit(text, _MATH_FORMAT)
        elif tag == W + "pPr":
            return
        else:
            self._markup(element)

    def _inline_children(self, element: ET.Element) -> None:
        for child in _children(element):
            self._inline(child)

    def _markup(self, element: ET.Element) -> None:
        tag = element.tag
        if tag == W + "bookmarkStart":
            name = element.get(W + "name")
            if name:
                self.doc.bookmarks.add(name)
        elif tag == W + "commentRangeStart":
            identifier = element.get(W + "id", "")
            if identifier not in self.comment_anchor:
                self.comment_anchor[identifier] = []
                self._captures.append(self.comment_anchor[identifier])
            if identifier not in self.comment_order:
                self.comment_order.append(identifier)
        elif tag == W + "commentRangeEnd":
            buffer = self.comment_anchor.get(element.get(W + "id", ""))
            if buffer is not None:
                self._release(buffer)

    def _run(self, run: ET.Element) -> None:
        fmt = self.doc.run_format(self._style, run.find(W + "rPr"), self._table_style)
        for child in _children(run):
            self._run_child(child, fmt)

    def _cell_blocks(
        self, cell: ET.Element, table_style: str | None
    ) -> list[dict[str, Any]]:
        """Cell content formatted with the table style (or the default table style)."""
        styles = self.doc.styles
        effective = (
            table_style if table_style in styles.by_id else styles.defaults.get("table")
        )
        saved = self._table_style
        self._table_style = effective
        blocks = self.blocks(cell)
        self._table_style = saved
        return blocks

    def _run_child(self, child: ET.Element, fmt: dict[str, Any]) -> None:
        tag = child.tag
        if tag == W + "t":
            text = child.text or ""
            if child.get(XML_SPACE) != "preserve":
                text = text.strip(" \t\r\n")
            self._emit(text, fmt)
        elif tag == W + "delText":
            if self._deleted:
                self._emit(child.text or "", fmt)
        elif tag == W + "instrText":
            if self._field_stack and self._field_stack[-1]["phase"] == "code":
                self._field_stack[-1]["instr"].append(child.text or "")
        elif tag == W + "fldChar":
            self._field_char(child.get(W + "fldCharType"))
        elif tag in (W + "tab", W + "ptab"):
            self._emit("\t", fmt)
        elif tag == W + "br":
            kind = child.get(W + "type")
            self._emit(
                "\f" if kind == "page" else "\v" if kind == "column" else "\n", fmt
            )
        elif tag == W + "cr":
            self._emit("\n", fmt)
        elif tag == W + "noBreakHyphen":
            self._emit("\u2011", fmt)
        elif tag == W + "softHyphen":
            self._emit("\u00ad", fmt)
        elif tag == W + "sym":
            try:
                self._emit(chr(int(child.get(W + "char", ""), 16)), fmt)
            except ValueError:
                pass
        elif tag == W + "drawing":
            self._drawing(child)
            self._emit(_OBJECT_MARK, fmt)
        elif tag == W + "pict":
            self._pict(child)
            self._emit(_OBJECT_MARK, fmt)
        elif tag == W + "object":
            self._object(child)
            self._emit(_OBJECT_MARK, fmt)
        elif tag == W + "footnoteReference":
            self.note_refs.append(("footnote", child.get(W + "id", "")))
        elif tag == W + "endnoteReference":
            self.note_refs.append(("endnote", child.get(W + "id", "")))
        elif tag == W + "commentReference":
            identifier = child.get(W + "id", "")
            if identifier not in self.comment_order:
                self.comment_order.append(identifier)
        elif tag == W + "ruby":
            base = child.find(W + "rubyBase")
            if base is not None:
                self._inline_children(base)

    # -- fields -----------------------------------------------------------

    def _field_char(self, kind: str | None) -> None:
        if kind == "begin":
            self._field_stack.append(
                {
                    "phase": "code",
                    "instr": [],
                    "index": len(self.fields),
                    "link": len(self.hyperlinks),
                    "buffer": None,
                }
            )
            self.fields.append(None)
            self.hyperlinks.append(None)
        elif kind == "separate" and self._field_stack:
            field = self._field_stack[-1]
            if field["phase"] == "code":
                self._close_instruction(field)
                field["phase"] = "result"
                if self.hyperlinks[field["link"]] is not None:
                    field["buffer"] = []
                    self._captures.append(field["buffer"])
        elif kind == "end" and self._field_stack:
            field = self._field_stack.pop()
            if field["phase"] == "code":
                self._close_instruction(field)
            if field["buffer"] is not None:
                self.hyperlinks[field["link"]]["text"] = self._release(field["buffer"])

    def _close_instruction(self, field: dict[str, Any]) -> None:
        instruction = _normalize_instruction("".join(field["instr"]))
        self.fields[field["index"]] = instruction or None
        link = _hyperlink_field(instruction)
        if link is not None:
            self.hyperlinks[field["link"]] = {
                "text": "",
                "url": link[0],
                "anchor": link[1],
            }

    def _simple_field(self, element: ET.Element) -> None:
        instruction = _normalize_instruction(element.get(W + "instr", ""))
        self.fields.append(instruction or None)
        link = _hyperlink_field(instruction)
        record = None
        if link is not None:
            record = {"text": "", "url": link[0], "anchor": link[1]}
            self.hyperlinks.append(record)
        buffer: list[str] = []
        self._captures.append(buffer)
        self._simple_fields += 1
        self._inline_children(element)
        self._simple_fields -= 1
        text = self._release(buffer)
        if record is not None:
            record["text"] = text

    def finish(self) -> None:
        while self._field_stack:
            self._field_char("end")

    # -- tracked changes, content controls --------------------------------

    def _change(self, kind: str, element: ET.Element, walk: Any) -> None:
        buffer: list[str] = []
        start = self._counter
        deleting = kind in ("deletion", "move_from")
        (self._deleted if deleting else self._captures).append(buffer)
        walk()
        if deleting:
            self._deleted.pop()
            text = "".join(buffer)
        else:
            text = self._release(buffer)
        if not text:
            return
        author = element.get(W + "author")
        last = self.changes[-1] if self.changes else None
        if (
            last is not None
            and last["kind"] == kind
            and last["author"] == author
            and self._last_change_end == start
        ):
            last["text"] += text
        else:
            self.changes.append({"kind": kind, "author": author, "text": text})
        self._last_change_end = self._counter

    def _content_control(self, element: ET.Element, walk: Any) -> None:
        props = element.find(W + "sdtPr")
        kind = "richText"
        if props is not None:
            for name in _SDT_KINDS:
                if any(_local(child.tag) == name for child in props):
                    kind = name
                    break
        record = {
            "tag": _val(props.find(W + "tag")) if props is not None else None,
            "alias": _val(props.find(W + "alias")) if props is not None else None,
            "kind": kind,
            "text": "",
        }
        self.content_controls.append(record)
        content = element.find(W + "sdtContent")
        if content is None:
            return
        buffer: list[str] = []
        self._captures.append(buffer)
        walk(content)
        record["text"] = self._release(buffer)

    # -- drawings and objects --------------------------------------------

    def _blip(self, blip: ET.Element) -> str | None:
        embed = blip.get(R + "embed")
        if embed:
            return self.doc.target_digest(self.part, embed)
        return self.doc.target_digest(self.part, blip.get(R + "link"))

    def _text_box(self, content: ET.Element) -> None:
        saved = (self._runs, self._style, self._table_style)
        self._runs, self._table_style = None, None
        blocks = self.blocks(content)
        self._runs, self._style, self._table_style = saved
        self.text_boxes.append(blocks)

    def _drawing(self, drawing: ET.Element) -> None:
        for frame in _children(drawing):
            placement = "inline" if _local(frame.tag) == "inline" else "anchor"
            extent = frame.find(WP + "extent")
            width = _emu_to_twips(extent.get("cx")) if extent is not None else None
            height = _emu_to_twips(extent.get("cy")) if extent is not None else None
            stop = frozenset({W + "txbxContent"})
            nodes = list(_descendants(frame, stop))
            charts = [node for node in nodes if node.tag == C + "chart"]
            blips = [self._blip(node) for node in nodes if node.tag == A + "blip"]
            for chart in charts:
                self.objects.append(self._chart(chart.get(R + "id")))
            if blips:
                self.images.append(
                    {
                        "placement": placement,
                        "width": width,
                        "height": height,
                        "media": blips,
                    }
                )
            elif not charts:
                self.shapes.append(
                    {"placement": placement, "width": width, "height": height}
                )
            for node in nodes:
                if node.tag == W + "txbxContent":
                    self._text_box(node)

    def _vml_frame(self, shape: ET.Element) -> dict[str, Any]:
        style = _css_lengths(shape.get("style"))
        return {
            "placement": "anchor" if style.get("position") == "absolute" else "inline",
            "width": _css_twips(style.get("width")),
            "height": _css_twips(style.get("height")),
        }

    def _vml_images(self, shape: ET.Element) -> list[str | None]:
        stop = frozenset({W + "txbxContent"})
        media = []
        for node in _descendants(shape, stop):
            if node.tag == V + "imagedata":
                media.append(
                    self.doc.target_digest(
                        self.part, node.get(R + "id") or node.get(O + "relid")
                    )
                )
        if shape.tag == V + "image" and shape.get(R + "id"):
            media.append(self.doc.target_digest(self.part, shape.get(R + "id")))
        return media

    def _pict(self, pict: ET.Element) -> None:
        for shape in _children(pict):
            if shape.tag == W + "control":
                self.objects.append(self._control(shape))
                continue
            if shape.tag not in _VML_SHAPES:
                continue
            frame = self._vml_frame(shape)
            media = self._vml_images(shape)
            if media:
                self.images.append({**frame, "media": media})
            else:
                self.shapes.append(frame)
            for node in _descendants(shape, frozenset({W + "txbxContent"})):
                if node.tag == W + "txbxContent":
                    self._text_box(node)

    def _object(self, obj: ET.Element) -> None:
        preview: list[str | None] = []
        size: dict[str, Any] = {"width": None, "height": None}
        for shape in _children(obj):
            if shape.tag in _VML_SHAPES:
                frame = self._vml_frame(shape)
                size = {"width": frame["width"], "height": frame["height"]}
                preview.extend(self._vml_images(shape))
        control = obj.find(W + "control")
        if control is not None:
            record = self._control(control)
        else:
            ole = obj.find(O + "OLEObject")
            embed = obj.find(W + "objectEmbed")
            link = obj.find(W + "objectLink")
            if ole is not None:
                record = {
                    "kind": "ole",
                    "prog_id": ole.get("ProgID"),
                    "link": ole.get("Type") == "Link",
                    "data": self.doc.target_digest(
                        self.part, ole.get(R + "id"), payload=True
                    ),
                }
            elif embed is not None or link is not None:
                source = embed if embed is not None else link
                record = {
                    "kind": "ole",
                    "prog_id": source.get(W + "progId"),
                    "link": link is not None,
                    "data": self.doc.target_digest(
                        self.part, source.get(R + "id"), payload=True
                    ),
                }
            else:
                record = {"kind": "object"}
        record.update(size)
        record["preview"] = preview
        self.objects.append(record)

    def _control(self, control: ET.Element) -> dict[str, Any]:
        record: dict[str, Any] = {"kind": "control", "name": control.get(W + "name")}
        rel = self.doc.pkg.rels(self.part).get(control.get(R + "id") or "")
        if rel is None or rel[2] or rel[1] not in self.doc.pkg.data:
            record["class_id"] = None
            return record
        target = rel[1]
        data = self.doc.pkg.data[target]
        if target.endswith(".xml"):
            root = self.doc.pkg.xml(target)
            record["class_id"] = (root.get(AX + "classid") or "").upper() or None
            record["persistence"] = root.get(AX + "persistence")
            record["properties"] = sorted(
                [prop.get(AX + "name"), prop.get(AX + "value")]
                for prop in root.findall(AX + "ocxPr")
            )
            binary = root.get(R + "id")
            record["data"] = self.doc.target_digest(target, binary) if binary else None
        else:
            record["class_id"] = None
            record["data"] = _sha256(data)
        return record

    def _chart(self, rel_id: str | None) -> dict[str, Any]:
        rel = self.doc.pkg.rels(self.part).get(rel_id or "")
        if rel is None or rel[2] or rel[1] not in self.doc.pkg.data:
            return {"kind": "chart", "types": None, "title": None, "series": None}
        root = self.doc.pkg.xml(rel[1])
        plot = root.find(f"{C}chart/{C}plotArea")
        types = (
            [_local(child.tag) for child in plot if _local(child.tag).endswith("Chart")]
            if plot is not None
            else []
        )
        title = root.find(f"{C}chart/{C}title")
        series = []
        for item in root.iter(C + "ser"):
            name = item.find(C + "tx")
            values = item.find(C + "val")
            if values is None:
                values = item.find(C + "yVal")
            series.append(
                {
                    "name": "".join(name.itertext()).strip()
                    if name is not None
                    else None,
                    "values": (
                        "".join(node.text or "" for node in values.iter(C + "f"))
                        or None
                        if values is not None
                        else None
                    ),
                }
            )
        return {
            "kind": "chart",
            "types": types,
            "title": "".join(node.text or "" for node in title.iter(A + "t"))
            if title is not None
            else None,
            "series": series,
        }


def _blocks_text(blocks: list[dict[str, Any]]) -> str:
    lines: list[str] = []
    for block in blocks:
        if block["type"] == "paragraph":
            lines.append("".join(run["text"] for run in block["runs"]))
        elif block["type"] == "table":
            for row in block["rows"]:
                for cell in row["cells"]:
                    lines.append(_blocks_text(cell["blocks"]))
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Features
# ---------------------------------------------------------------------------


def _section(
    section: ET.Element, first: bool, chrome: dict[str, Any]
) -> dict[str, Any]:
    size = section.find(W + "pgSz")
    width = _twips(size.get(W + "w")) if size is not None else None
    height = _twips(size.get(W + "h")) if size is not None else None
    if width is not None and height is not None:
        orientation = "landscape" if width > height else "portrait"
    else:
        orientation = (
            size.get(W + "orient") if size is not None else None
        ) or "portrait"
    margins = section.find(W + "pgMar")
    columns = section.find(W + "cols")
    count = int(_number(columns.get(W + "num")) or 1) if columns is not None else 1
    equal = _onoff_attr(columns.get(W + "equalWidth")) if columns is not None else None
    equal = True if equal is None else equal
    explicit = (
        [
            [_twips(column.get(W + "w")), _twips(column.get(W + "space"))]
            for column in columns.findall(W + "col")
        ]
        if columns is not None and not equal
        else []
    )
    space = _twips(columns.get(W + "space")) if columns is not None else None
    numbering = section.find(W + "pgNumType")
    start = numbering.get(W + "start") if numbering is not None else None
    return {
        "type": None if first else (_val(section.find(W + "type")) or "nextPage"),
        "page": {"width": width, "height": height, "orientation": orientation},
        "margins": {
            # Header/footer distance has no effect on a section without one.
            side: None
            if margins is None or (side in chrome and not any(chrome[side].values()))
            else _twips(margins.get(W + side))
            for side in ("top", "right", "bottom", "left", "header", "footer", "gutter")
        },
        "columns": {
            "count": count,
            "space": (720 if space is None else space)
            if count > 1 or not equal
            else None,
            "equal_width": equal,
            "widths": explicit,
        },
        "page_numbering": {
            "format": (numbering.get(W + "fmt") if numbering is not None else None)
            or "decimal",
            "start": int(start) if start and start.lstrip("-").isdigit() else None,
        },
    }


def _headers_footers(
    doc: _Document, sections: list[ET.Element]
) -> list[dict[str, Any]]:
    rels = doc.pkg.rels(doc.main)
    cache: dict[str, dict[str, Any] | None] = {}

    def story(part: str | None) -> dict[str, Any] | None:
        if part is None:
            return None
        if part not in cache:
            walker = _Story(doc, part, field_results=False)
            blocks = walker.blocks(doc.pkg.xml(part))
            walker.finish()
            text = _blocks_text(blocks).strip()
            fields = [field for field in walker.fields if field]
            images = [image["media"] for image in walker.images]
            cache[part] = (
                {"text": text, "fields": fields, "images": images}
                if text or fields or images
                else None
            )
        return cache[part]

    current: dict[tuple[str, str], str | None] = {}
    result: list[dict[str, Any]] = []
    for section in sections:
        for kind in ("header", "footer"):
            for reference in section.findall(f"{W}{kind}Reference"):
                rel = rels.get(reference.get(R + "id", ""))
                target = (
                    rel[1]
                    if rel is not None and not rel[2] and rel[1] in doc.pkg.data
                    else None
                )
                current[(kind, reference.get(W + "type", "default"))] = target
        title_page = bool(_onoff(section.find(W + "titlePg")))
        entry: dict[str, Any] = {}
        for kind in ("header", "footer"):
            default = current.get((kind, "default"))
            entry[kind] = {
                "first": story(current.get((kind, "first")) if title_page else default),
                "odd": story(default),
                "even": story(
                    current.get((kind, "even")) if doc.even_and_odd else default
                ),
            }
        result.append(entry)
    return result


def _notes(doc: _Document, body: _Story, kind: str) -> list[dict[str, Any]]:
    part = doc.pkg.first_related(doc.main, kind + "s")
    notes: dict[str, ET.Element] = {}
    if part is not None:
        for note in doc.pkg.xml(part).findall(W + kind):
            notes.setdefault(note.get(W + "id", ""), note)
    result: list[dict[str, Any]] = []
    for note_kind, identifier in body.note_refs:
        if note_kind != kind:
            continue
        note = notes.get(identifier)
        if note is None or part is None:
            result.append({"text": None})
            continue
        walker = _Story(doc, part)
        blocks = walker.blocks(note)
        walker.finish()
        result.append({"text": _blocks_text(blocks).strip()})
    return result


def _comments(doc: _Document, body: _Story) -> list[dict[str, Any]]:
    part = doc.pkg.first_related(doc.main, "comments")
    comments: dict[str, ET.Element] = {}
    if part is not None:
        for comment in doc.pkg.xml(part).findall(W + "comment"):
            comments.setdefault(comment.get(W + "id", ""), comment)
    result: list[dict[str, Any]] = []
    for identifier in body.comment_order:
        comment = comments.get(identifier)
        anchor = "".join(body.comment_anchor.get(identifier, [])).strip("\n")
        if comment is None or part is None:
            result.append({"author": None, "text": None, "anchor": anchor})
            continue
        walker = _Story(doc, part)
        blocks = walker.blocks(comment)
        walker.finish()
        result.append(
            {
                "author": comment.get(W + "author"),
                "text": _blocks_text(blocks).strip(),
                "anchor": anchor,
            }
        )
    return result


def _core_properties(doc: _Document) -> dict[str, Any]:
    part = doc.pkg.first_related("", "core-properties")
    root = doc.pkg.xml(part) if part else None
    result: dict[str, Any] = {}
    for key, tag in (
        ("title", DC + "title"),
        ("subject", DC + "subject"),
        ("creator", DC + "creator"),
        ("keywords", CP + "keywords"),
    ):
        element = root.find(tag) if root is not None else None
        text = "".join(element.itertext()) if element is not None else ""
        result[key] = text or None
    return result


def _typed_value(value: ET.Element) -> dict[str, Any]:
    kind = _local(value.tag)
    text = (value.text or "").strip()
    if kind in ("lpwstr", "lpstr", "bstr"):
        return {"type": "string", "value": value.text or ""}
    if kind == "bool":
        return {"type": "boolean", "value": text.lower() in _TRUE}
    if kind in (
        "i1",
        "i2",
        "i4",
        "i8",
        "int",
        "ui1",
        "ui2",
        "ui4",
        "ui8",
        "uint",
        "r4",
        "r8",
        "decimal",
    ):
        try:
            number: int | float = int(text)
        except ValueError:
            try:
                number = float(text)
            except ValueError:
                return {"type": kind, "value": text}
            if number.is_integer():
                number = int(number)
        return {"type": "number", "value": number}
    if kind in ("filetime", "date"):
        try:
            moment = datetime.fromisoformat(text)
        except ValueError:
            return {"type": "date", "value": text}
        if moment.tzinfo is None:
            moment = moment.replace(tzinfo=UTC)
        return {
            "type": "date",
            "value": moment.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ"),
        }
    return {"type": kind, "value": "".join(value.itertext()).strip()}


def _custom_properties(doc: _Document) -> dict[str, Any]:
    part = doc.pkg.first_related("", "custom-properties")
    if part is None:
        return {}
    result: dict[str, Any] = {}
    for item in doc.pkg.xml(part).findall(CUSTOM + "property"):
        value = next(iter(item), None)
        record = (
            _typed_value(value) if value is not None else {"type": None, "value": None}
        )
        if item.get("linkTarget"):
            record["link"] = item.get("linkTarget")
        result[item.get("name", "")] = record
    return dict(sorted(result.items()))


def _custom_xml(doc: _Document) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for part in sorted(set(doc.pkg.related(doc.main, "customXml"))):
        data = doc.pkg.data[part]
        try:
            canonical = ET.canonicalize(
                data.decode("utf-8-sig"), strip_text=True, rewrite_prefixes=True
            )
            digest = _sha256(canonical.encode())
        except (ET.ParseError, UnicodeDecodeError, ValueError):
            digest = _sha256(data)
        item_id = None
        schemas: list[str] = []
        props = doc.pkg.first_related(part, "customXmlProps")
        if props is not None:
            root = doc.pkg.xml(props)
            item_id = (root.get(DS + "itemID") or "").upper() or None
            schemas = sorted(
                ref.get(DS + "uri", "") for ref in root.iter(DS + "schemaRef")
            )
        result.append({"sha256": digest, "item_id": item_id, "schemas": schemas})
    return sorted(result, key=lambda item: json.dumps(item, sort_keys=True))


def _glossary(doc: _Document) -> list[dict[str, Any]]:
    part = doc.pkg.first_related(doc.main, "glossaryDocument")
    if part is None:
        return []
    entries = []
    for entry in doc.pkg.xml(part).iter(W + "docPart"):
        props = entry.find(W + "docPartPr")
        body = entry.find(W + "docPartBody")
        walker = _Story(doc, part)
        blocks = walker.blocks(body) if body is not None else []
        walker.finish()
        entries.append(
            {
                "gallery": _val(props.find(f"{W}category/{W}gallery"))
                if props is not None
                else None,
                "category": _val(props.find(f"{W}category/{W}name"))
                if props is not None
                else None,
                "name": _val(props.find(W + "name")) if props is not None else None,
                "text": _blocks_text(blocks).strip(),
            }
        )
    return sorted(entries, key=lambda item: json.dumps(item, sort_keys=True))


def extract(path: Path) -> dict[str, Any]:
    """Return the reader-observable model of a .docx package (see module docstring)."""
    doc = _Document(Path(path))
    try:
        body = _Story(doc, doc.main)
        blocks = body.blocks(doc.body)
        body.finish()
        sections = list(body.sections)
        headers_footers = _headers_footers(doc, sections)
        model: dict[str, Any] = {
            "body": blocks,
            "sections": [
                _section(section, index == 0, headers_footers[index])
                for index, section in enumerate(sections)
            ],
            "headers_footers": headers_footers,
            "footnotes": _notes(doc, body, "footnote"),
            "endnotes": _notes(doc, body, "endnote"),
            "comments": _comments(doc, body),
            "tracked_changes": body.changes,
            "fields": [field for field in body.fields if field],
            "bookmarks": [],
            "hyperlinks": [link for link in body.hyperlinks if link is not None],
            "content_controls": body.content_controls,
            "images": body.images,
            "shapes": body.shapes,
            "text_boxes": body.text_boxes,
            "embedded_objects": body.objects,
            "styles": doc.styles.defined(),
            "custom_properties": _custom_properties(doc),
            "core_properties": _core_properties(doc),
            "custom_xml": _custom_xml(doc),
            "glossary": _glossary(doc),
        }
        model["bookmarks"] = sorted(doc.bookmarks)
    except (KeyError, IndexError, TypeError, AttributeError, RecursionError) as exc:
        raise ValueError(f"{path}: unreadable document structure: {exc!r}") from exc
    # Round-trip through JSON so the model is exactly what a stored copy would hold.
    return json.loads(json.dumps(model, ensure_ascii=False))


# ---------------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------------


def _pointer(path: str, key: str | int) -> str:
    return f"{path}/{str(key).replace('~', '~0').replace('/', '~1')}"


def _canonical(value: Any) -> str:
    return json.dumps(value, sort_keys=True, ensure_ascii=False)


def _within_tolerance(path: str, before: Any, after: Any) -> bool:
    """Lengths (twips) within _LENGTH_TOLERANCE are equal: unit-conversion rounding."""
    numbers = (int, float)
    if not (isinstance(before, numbers) and isinstance(after, numbers)):
        return False
    if isinstance(before, bool) or isinstance(after, bool):
        return False
    named = [
        segment for segment in path.split("/") if segment and not segment.isdigit()
    ]
    if not named or named[-1] not in _LENGTH_KEYS:
        return False
    return abs(before - after) <= _LENGTH_TOLERANCE


def _diff(
    feature: str, path: str, before: Any, after: Any, out: list[dict[str, Any]]
) -> None:
    if before == after and type(before) is type(after):
        return
    if isinstance(before, dict) and isinstance(after, dict):
        for key in sorted(set(before) | set(after), key=str):
            child = _pointer(path, key)
            if key not in after:
                out.append(_entry(feature, child, "missing", before[key], None))
            elif key not in before:
                out.append(_entry(feature, child, "added", None, after[key]))
            else:
                _diff(feature, child, before[key], after[key], out)
        return
    if isinstance(before, list) and isinstance(after, list):
        matcher = difflib.SequenceMatcher(
            None,
            [_canonical(item) for item in before],
            [_canonical(item) for item in after],
            autojunk=False,
        )
        for op, b_start, b_end, a_start, a_end in matcher.get_opcodes():
            if op == "equal":
                continue
            paired = min(b_end - b_start, a_end - a_start) if op == "replace" else 0
            for offset in range(paired):
                _diff(
                    feature,
                    _pointer(path, b_start + offset),
                    before[b_start + offset],
                    after[a_start + offset],
                    out,
                )
            for index in range(b_start + paired, b_end):
                out.append(
                    _entry(
                        feature, _pointer(path, index), "missing", before[index], None
                    )
                )
            for index in range(a_start + paired, a_end):
                out.append(
                    _entry(feature, _pointer(path, index), "added", None, after[index])
                )
        return
    if _within_tolerance(path, before, after):
        return
    out.append(_entry(feature, path, "changed", before, after))


def _entry(
    feature: str, path: str, kind: str, before: Any, after: Any
) -> dict[str, Any]:
    return {
        "feature": feature,
        "path": path,
        "kind": kind,
        "before": before,
        "after": after,
    }


def compare(before: dict[str, Any], after: dict[str, Any]) -> list[dict[str, Any]]:
    """Differences between two extracted models; an empty list means preserved."""
    out: list[dict[str, Any]] = []
    for feature in FEATURES:
        _diff(
            feature, _pointer("", feature), before.get(feature), after.get(feature), out
        )
    return out
