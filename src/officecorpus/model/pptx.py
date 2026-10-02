"""Viewer-observable content model for PowerPoint (.pptx) packages.

`extract` reduces a deck to what a person viewing it can observe; `compare` diffs two
such models. Two files with equal models are treated as the same deck, however their
XML is serialized. Only the standard library is used.

Shared cross-format rule (identical in the docx, xlsx and pptx models)
  * Presentation-only lengths compare equal when they differ by at most 0.5 pt
    (6350 EMU). Here that covers absolute shape offsets and extents, line widths,
    table column widths and row heights, slide size, and custom-geometry path points
    and arc radii. Font sizes are NOT lengths for this rule and compare exactly
    (10 pt vs 10.5 pt is a user-selectable difference).
  * Color modifiers (tint, shade, lumMod, lumOff, satMod, alpha, ...; 1/1000 %) compare
    by rendered effect: equal when |delta| <= 100000/510 (~196 units), i.e. they cannot
    move an 8-bit channel by a full level. This model resolves colors to 8-bit RGB, so
    the rule is applied to the resolved value: a color computed through modifiers is
    marked "~" (e.g. "7C7C7C~/a50"), and when either side carries the mark the colors
    are equal if every channel differs by at most one level, alpha by at most 100/510
    percent, and unknown modifiers match. Colors written without modifiers compare
    exactly.
  * Angles (shape rotation, group rotation folded into children, gradient linear
    angle, custom-geometry arc start and sweep angles) follow the 0.5 pt rule: they
    compare equal when |delta| <= 0.5 / max(slide cx, slide cy in pt) radians, so no
    point on the slide moves more than 0.5 pt (about 0.03 degree on a 960 pt-wide
    slide). The limit comes from the source deck's own slide size (the rewritten
    deck's if the source has none); rotations and start angles wrap mod 360 degrees,
    sweeps do not. Text rotation is not modeled (see Not modeled).
  * Other presentation-only non-length numbers compare equal within a relative 1e-9
    after unit normalization (crop and gradient stop fractions in 1/1000 %,
    transition durations, animation preset numbers). Every one of these is an
    integer in the model, so in practice they compare exactly.
  * Text, font sizes, slide order, z-order, media/image/object hashes and every other
    content value (chart caches, names, explicit colors, flags) compare exactly.
  * Custom-geometry paths are evaluated (shape guides and formulas resolved) and scaled
    to EMU offsets inside the placed shape box, so they are lengths like any other.

Rules
-----
Identity and structure
  * Part names, relationship ids, shape ids, slide ids, and XML layout never appear in
    the model. Relationships are followed by type (last segment of the Type URI, so
    transitional and strict namespaces agree) and resolved target. Slide order comes
    from p:sldIdLst; masters from p:sldMasterIdLst.
  * Elements are matched by local name, so namespace prefixes and transitional/strict
    namespace URIs are irrelevant. Attribute order, whitespace-only text, comments and
    processing instructions are irrelevant.
  * mc:AlternateContent is replaced by the content of its first mc:Choice (what a
    current PowerPoint renders), or its mc:Fallback when no Choice exists.
  * Z-order matters: shapes keep document order inside p:spTree and p:grpSp.

Values
  * Geometry is absolute slide geometry: group child coordinate spaces (chOff/chExt)
    and group rotation/mirroring are applied to every descendant, so a group written
    with any child space places its children identically. flipV is folded into
    flipH + 180 degrees (same drawing, same upside-down text); rot is normalized to
    [0, 360) degrees. Coordinates are rounded to whole EMU after placement.
  * A group's own frame is not drawn, so a group has no geometry in the model; only
    its descendants' absolute geometry counts (refitting or un-rotating a group
    while keeping every child in place is not a change).
  * A straight line (preset line/straightConnector1) is modeled by its absolute start
    and end points instead of box, rotation and flips: a rotated line and an
    unrotated, flipped line through the same two points draw identically. Endpoint
    order is kept because line ends (arrowheads) attach to start and end.
  * Preset geometry: the preset name (presets with identical ECMA-376 definitions are
    one name: straightConnector1 == line) plus adjust values evaluated from their
    formulas. Only sp, cxnSp and pic carry geometry; a pic without one is a rect.
  * Rotation (60000ths of a degree), crop (1000ths of a percent), font size (100ths of
    a point) and line width (EMU) are integers in their native units; the shared rule
    above sets their comparison tolerance.
  * Booleans accept "1"/"true"/"on" and "0"/"false"/"off".
  * Absent attributes take their schema default (rot 0, flips false, bold false,
    underline "none", algn "l", gridSpan/rowSpan 1, show true, ...).
  * Colors resolve to uppercase RRGGBB: schemeClr through the effective color map
    (slide/layout clrMapOvr, else master clrMap) and the master's theme; sysClr through
    lastClr (or the documented system defaults); prstClr, hslClr and scrgbClr are
    converted. lumMod/lumOff/satMod/satOff/hueMod/hueOff apply in HSL, tint/shade in
    linear RGB, gamma/invGamma convert linear->sRGB / sRGB->linear; results are rounded
    to 8-bit channels; alpha below 100% is appended as "/aN" (exact percent, e.g.
    "/a50.4"). Unknown modifiers are appended verbatim.
  * Numeric chart cache values compare numerically ("1", "1.0", "1E0" are equal);
    text values compare exactly.
  * External link targets are percent-decoded; file: URLs use forward slashes.
  * Slide size is cx/cy only (the sldSz type label is page-setup metadata).

Inheritance (attribute defaults filled)
  * A placeholder without its own xfrm, geometry, fill or line takes them from the
    matching layout placeholder, then the matching master placeholder (layout match:
    idx, then type; master match: type, with ctrTitle -> title and content types -> body).
  * Shape fill/line without explicit spPr values resolve p:style fillRef/lnRef through
    the theme format scheme with the reference color as phClr.
  * Paragraph alignment and bullets and run bold/italic/underline/size/color/latin font
    resolve through: own pPr/rPr, the shape lstStyle, p:style fontRef (font, color), the
    layout and master placeholder lstStyles, the master title/body/other text style,
    and the presentation defaultTextStyle. Theme font references (+mj-lt, +mn-lt)
    resolve to the theme's major/minor latin typeface.
  * Slide background resolves slide -> layout -> master.
  * A slide's master name is the master cSld name, else its theme name.
  * Lines: an absent line fill means no line; an absent width is 0; a line whose fill
    is none records only its fill (width, dash and arrowheads are not drawn).
  * Transitions: no p:transition equals a transition without an effect element; speed
    and duration are recorded only when there is an effect.
  * Animations are present when the timing tree holds an effect preset or a behavior
    (anim*, set, cmd, audio, video); an empty root time node is no animation.

Text
  * Run boundaries are invisible: adjacent runs with identical effective formatting
    (and identical hyperlink) merge; empty runs drop. a:br is a "\\n" run, a:fld is a run
    carrying its field type. Empty paragraphs are kept (they occupy a line) with bullet
    "none" (a bullet is not drawn without text).
  * Bullet characters in the symbol-font private-use range U+F020-U+F0FF equal their
    8-bit codes (U+F06C == "l").
  * Hyperlinks list (text, target) in reading order: run links carry the merged run
    text, shape click links carry the shape's text. Internal slide jumps resolve to the
    target slide's position ("slide:N").
  * Notes: the text of the notes body placeholder(s), paragraphs joined by "\\n"; empty
    notes equal absent notes.

Binary content
  * Images, media, embedded objects and chart workbooks compare by SHA-256 of bytes.
  * An embedded object's preview image is taken from either mc:AlternateContent
    branch, else from the legacy VML shape with the object's spid.
  * A media clip referenced by both a:videoFile/a:audioFile and p14:media counts once.
  * Custom XML parts compare by SHA-256 of a canonical tree (expanded names, sorted
    attributes, whitespace-only text dropped) so re-serialization is not a change.

Comparison
  * `compare` walks both models: dict keys compare as sets (missing/added), lists are
    aligned with difflib on canonical JSON so one inserted or deleted item reports once
    instead of shifting every later item; aligned pairs recurse; scalars report
    "changed". Paths are RFC 6901 JSON Pointers into the model ("before" index for
    missing and paired items, "after" index for added items).
  * Each difference carries the FEATURES entry of the innermost model key that names a
    feature; containers such as tables, charts and fill-line keep their feature for all
    nested keys.

Model coverage per slide: hidden flag, layout and master name, background, shape tree
(kind, name, hidden, alt text, geometry, preset/custom geometry, fill, line,
placeholder, text, tables, pictures, groups, charts, SmartArt, embedded objects),
hyperlinks, notes, comments, transition, animation presence/effects, media. Package:
slide size, themes (name, color scheme, fonts), masters with their non-placeholder
shapes and layouts (name, show-master-shapes flag, non-placeholder shapes; layouts are
a set sorted by name), custom XML, sections.

Not modeled: text body properties (insets, anchoring, autofit, columns), paragraph
spacing/indents, character spacing/strike/baseline/east-asian and complex-script fonts,
shape effects (shadow, glow, 3-D), image effects, table-style definitions and table
text inheritance from table styles, chart formatting beyond chart type/title/legend/
data, chartex (cx:) plots beyond series layout and data dimensions, SmartArt styling,
transition effect option defaults (options compare as written), animation targets and
timing details, custom-geometry text rectangles and connection sites, rendered output
of SmartArt drawings and ink, exact agreement with applications that bake modified
theme colors into fixed sRGB (their rounding can differ by 1-2 per channel), document
properties, VBA.
"""

from __future__ import annotations

import colorsys
import difflib
import hashlib
import json
import math
import pathlib
import posixpath
import re
import zipfile
from typing import Any
from urllib.parse import unquote
from xml.etree import ElementTree as ET

FEATURES: tuple[str, ...] = (
    "package",
    "slide-size",
    "slides",
    "slide-hidden",
    "slide-layout",
    "background",
    "shapes",
    "geometry",
    "fill-line",
    "placeholders",
    "text",
    "text-format",
    "tables",
    "pictures",
    "groups",
    "charts",
    "smartart",
    "embedded-objects",
    "hyperlinks",
    "notes",
    "comments",
    "transitions",
    "animations",
    "media",
    "theme",
    "masters-layouts",
    "custom-xml",
    "sections",
)

_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_CT_PART = "[content_types].xml"
_TRUE = {"1", "true", "on"}

# ---------------------------------------------------------------------------
# XML helpers (local-name based, mc:AlternateContent aware)
# ---------------------------------------------------------------------------


def _ln(tag: Any) -> str:
    if not isinstance(tag, str):
        return ""
    return tag[tag.rfind("}") + 1 :]


def _ns(tag: str) -> str:
    return tag[1:].split("}", 1)[0] if tag.startswith("{") else ""


_AC_TAG = "{" + _MC + "}AlternateContent"


def _kids(el: ET.Element | None) -> list[ET.Element]:
    """Children of `el` with mc:AlternateContent replaced by its chosen branch."""
    if el is None:
        return []
    out: list[ET.Element] = []
    for child in el:
        if child.tag != _AC_TAG:
            out.append(child)
            continue
        branch = next((c for c in child if _ln(c.tag) == "Choice"), None)
        if branch is None:
            branch = next((c for c in child if _ln(c.tag) == "Fallback"), None)
        out.extend(_kids(branch))
    return out


def _child(el: ET.Element | None, *names: str) -> ET.Element | None:
    if el is None:
        return None
    for child in el:
        if child.tag == _AC_TAG:
            for sub in _kids(el):
                if _ln(sub.tag) in names:
                    return sub
            return None
    for child in el:
        if _ln(child.tag) in names:
            return child
    return None


def _path(el: ET.Element | None, *names: str) -> ET.Element | None:
    for name in names:
        el = _child(el, name)
        if el is None:
            return None
    return el


def _children(el: ET.Element | None, name: str) -> list[ET.Element]:
    return [c for c in _kids(el) if _ln(c.tag) == name]


def _walk(el: ET.Element | None):
    """Descendants (document order), expanding mc:AlternateContent."""
    for child in _kids(el):
        yield child
        yield from _walk(child)


def _attr(el: ET.Element | None, name: str) -> str | None:
    if el is None:
        return None
    if name in el.attrib:
        return el.attrib[name]
    for key, value in el.attrib.items():
        if _ln(key) == name and not key.startswith(
            "{http://schemas.openxmlformats.org/officeDocument"
        ):
            return value
    return None


def _rattr(el: ET.Element | None, name: str) -> str | None:
    """Attribute in a relationships namespace (r:id, r:embed, r:link, ...)."""
    if el is None:
        return None
    for key, value in el.attrib.items():
        if key.startswith("{") and _ln(key) == name and "relationships" in _ns(key):
            return value
    return None


def _int(value: str | None, default: int | None = None) -> int | None:
    if value is None:
        return default
    try:
        return int(value)
    except ValueError:
        try:
            return int(float(value))
        except ValueError:
            return default


def _bool(value: str | None, default: bool) -> bool:
    if value is None:
        return default
    return value.strip().lower() in _TRUE


def _text_of(el: ET.Element | None) -> str:
    return "".join((t.text or "") for t in _walk(el) if _ln(t.tag) == "t")


def _sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _canonical(el: ET.Element) -> Any:
    text = el.text if (el.text or "").strip() else None
    return [
        el.tag,
        sorted(el.attrib.items()),
        text,
        [_canonical(c) for c in el if isinstance(c.tag, str)],
    ]


def _num(value: str | None) -> str | None:
    if value is None:
        return None
    try:
        number = float(value)
    except ValueError:
        return value
    if math.isnan(number) or math.isinf(number):
        return value.strip()
    if number.is_integer() and abs(number) < 1e15:
        return str(int(number))
    return repr(number)


def _external(target: str) -> str:
    """External link target: percent-decoded; file: URLs use forward slashes."""
    target = unquote(target.strip())
    if target[:5].lower() == "file:":
        target = target.replace("\\", "/")
    return target


def _symbol_char(text: str) -> str:
    """Symbol-font private-use code points U+F020-U+F0FF equal their 8-bit codes."""
    return "".join(
        chr(ord(c) - 0xF000) if 0xF020 <= ord(c) <= 0xF0FF else c for c in text
    )


_ANIM_BEHAVIORS = {
    "anim",
    "animClr",
    "animEffect",
    "animMotion",
    "animRot",
    "animScale",
    "set",
    "cmd",
    "audio",
    "video",
}


# ---------------------------------------------------------------------------
# Package access
# ---------------------------------------------------------------------------


class _Pkg:
    def __init__(self, path: pathlib.Path) -> None:
        try:
            self.zip = zipfile.ZipFile(path)
            infos = self.zip.infolist()
        except (OSError, zipfile.BadZipFile, zipfile.LargeZipFile) as exc:
            raise ValueError(f"not a readable zip package: {exc}") from exc
        self.names: dict[str, str] = {}
        for info in infos:
            if not info.is_dir():
                self.names.setdefault(info.filename.lstrip("/").lower(), info.filename)
        self._xml: dict[str, ET.Element | None] = {}
        self._rels: dict[str, dict[str, tuple[str, str, bool]]] = {}
        ct = self.xml(_CT_PART)
        if ct is None:
            raise ValueError("package has no [Content_Types].xml")
        self.ct_default: dict[str, str] = {}
        self.ct_override: dict[str, str] = {}
        for item in ct:
            if _ln(item.tag) == "Default":
                ext = (item.get("Extension") or "").lower()
                self.ct_default[ext] = (item.get("ContentType") or "").lower()
            elif _ln(item.tag) == "Override":
                part = (item.get("PartName") or "").lstrip("/").lower()
                self.ct_override[unquote(part)] = (
                    item.get("ContentType") or ""
                ).lower()

    def data(self, part: str | None) -> bytes | None:
        if part is None:
            return None
        name = self.names.get(part)
        if name is None:
            return None
        try:
            return self.zip.read(name)
        except (OSError, zipfile.BadZipFile, KeyError, RuntimeError, ValueError) as exc:
            raise ValueError(f"cannot read part {part}: {exc}") from exc

    def xml(self, part: str | None) -> ET.Element | None:
        if part is None:
            return None
        if part not in self._xml:
            raw = self.data(part)
            if raw is None:
                self._xml[part] = None
            else:
                try:
                    self._xml[part] = ET.fromstring(raw)
                except ET.ParseError as exc:
                    raise ValueError(f"malformed XML in {part}: {exc}") from exc
        return self._xml[part]

    def content_type(self, part: str) -> str | None:
        if part in self.ct_override:
            return self.ct_override[part]
        ext = posixpath.splitext(part)[1].lstrip(".")
        return self.ct_default.get(ext)

    def rels(self, source: str) -> dict[str, tuple[str, str, bool]]:
        """rId -> (type suffix, resolved target, external)."""
        if source in self._rels:
            return self._rels[source]
        directory, base = posixpath.split(source)
        root = self.xml(posixpath.join(directory, "_rels", base + ".rels"))
        result: dict[str, tuple[str, str, bool]] = {}
        for rel in root if root is not None else []:
            if _ln(rel.tag) != "Relationship":
                continue
            rtype = (rel.get("Type") or "").rstrip("/").rsplit("/", 1)[-1]
            target = rel.get("Target") or ""
            external = (rel.get("TargetMode") or "").lower() == "external"
            if not external:
                target = _resolve(source, target)
            result[rel.get("Id") or ""] = (rtype, target, external)
        self._rels[source] = result
        return result

    def rel(self, source: str, rid: str | None) -> tuple[str, str, bool] | None:
        if rid is None:
            return None
        return self.rels(source).get(rid)

    def internal(self, source: str, rid: str | None) -> str | None:
        found = self.rel(source, rid)
        if found is None or found[2]:
            return None
        return found[1]

    def by_type(self, source: str, rtype: str) -> list[tuple[str, bool]]:
        return sorted(
            (target, external)
            for kind, target, external in self.rels(source).values()
            if kind == rtype
        )

    def first(self, source: str, rtype: str) -> str | None:
        for target, external in self.by_type(source, rtype):
            if not external:
                return target
        return None

    def blob_sha(self, part: str | None) -> str | None:
        raw = self.data(part)
        return None if raw is None else _sha(raw)


def _resolve(source: str, target: str) -> str:
    target = unquote(target.split("#", 1)[0]).replace("\\", "/")
    if target.startswith("/"):
        resolved = posixpath.normpath(target.lstrip("/"))
    else:
        resolved = posixpath.normpath(posixpath.join(posixpath.dirname(source), target))
    return resolved.lstrip("/").lower()


# ---------------------------------------------------------------------------
# Theme, color and fill resolution
# ---------------------------------------------------------------------------

_SYS = {
    "windowText": "000000",
    "window": "FFFFFF",
    "btnFace": "F0F0F0",
    "btnText": "000000",
    "highlight": "0078D7",
    "highlightText": "FFFFFF",
    "grayText": "6D6D6D",
    "menu": "F0F0F0",
    "menuText": "000000",
    "background": "000000",
    "captionText": "000000",
    "infoText": "000000",
    "infoBk": "FFFFE1",
    "3dDkShadow": "696969",
    "3dLight": "E3E3E3",
    "btnShadow": "A0A0A0",
    "btnHighlight": "FFFFFF",
    "windowFrame": "646464",
}
_PRESET = {
    "black": "000000",
    "white": "FFFFFF",
    "red": "FF0000",
    "green": "008000",
    "blue": "0000FF",
    "yellow": "FFFF00",
    "cyan": "00FFFF",
    "magenta": "FF00FF",
    "gray": "808080",
    "grey": "808080",
    "darkGray": "A9A9A9",
    "lightGray": "D3D3D3",
    "orange": "FFA500",
    "purple": "800080",
    "navy": "000080",
    "maroon": "800000",
    "olive": "808000",
    "teal": "008080",
    "silver": "C0C0C0",
    "lime": "00FF00",
    "aqua": "00FFFF",
    "fuchsia": "FF00FF",
    "brown": "A52A2A",
    "pink": "FFC0CB",
    "gold": "FFD700",
}
_COLOR_TAGS = {"srgbClr", "schemeClr", "sysClr", "prstClr", "hslClr", "scrgbClr"}
_FILL_TAGS = {"noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"}
_DEFAULT_CLRMAP = {
    "bg1": "lt1",
    "tx1": "dk1",
    "bg2": "lt2",
    "tx2": "dk2",
    **{f"accent{i}": f"accent{i}" for i in range(1, 7)},
    "hlink": "hlink",
    "folHlink": "folHlink",
}


class _Theme:
    def __init__(self, pkg: _Pkg, part: str | None) -> None:
        self.part = part
        root = pkg.xml(part)
        self.name = _attr(root, "name") or ""
        elements = _child(root, "themeElements")
        scheme = _child(elements, "clrScheme")
        self.scheme_name = _attr(scheme, "name") or ""
        self.colors: dict[str, str] = {}
        for slot in _kids(scheme):
            color = next((c for c in _kids(slot) if _ln(c.tag) in _COLOR_TAGS), None)
            if color is not None:
                self.colors[_ln(slot.tag)] = _base_color(color, None) or "000000"
        fonts = _child(elements, "fontScheme")
        self.major = _attr(_path(fonts, "majorFont", "latin"), "typeface") or ""
        self.minor = _attr(_path(fonts, "minorFont", "latin"), "typeface") or ""
        fmt = _child(elements, "fmtScheme")
        self.fills = _kids(_child(fmt, "fillStyleLst"))
        self.bg_fills = _kids(_child(fmt, "bgFillStyleLst"))
        self.lines = _kids(_child(fmt, "lnStyleLst"))


class _Ctx:
    """Resolution context: package, source part (for relationships), theme, color map."""

    def __init__(
        self,
        pkg: _Pkg,
        part: str,
        theme: _Theme,
        clrmap: dict[str, str],
        deck: _Deck | None,
    ) -> None:
        self.pkg = pkg
        self.part = part
        self.theme = theme
        self.clrmap = clrmap
        self.deck = deck

    def at(self, part: str) -> _Ctx:
        return _Ctx(self.pkg, part, self.theme, self.clrmap, self.deck)


def _clrmap(el: ET.Element | None) -> dict[str, str] | None:
    if el is None:
        return None
    mapping = dict(_DEFAULT_CLRMAP)
    for key, value in el.attrib.items():
        mapping[_ln(key)] = value
    return mapping


def _hex(value: str | None) -> str | None:
    if value is None:
        return None
    value = value.strip().upper()
    return value if re.fullmatch(r"[0-9A-F]{6}", value) else None


def _base_color(el: ET.Element, ctx: _Ctx | None, ph: str | None = None) -> str | None:
    tag = _ln(el.tag)
    if tag == "srgbClr":
        return _hex(_attr(el, "val"))
    if tag == "sysClr":
        return _hex(_attr(el, "lastClr")) or _SYS.get(_attr(el, "val") or "", "000000")
    if tag == "prstClr":
        return _PRESET.get(_attr(el, "val") or "")
    if tag == "hslClr":
        h = (_int(_attr(el, "hue"), 0) or 0) / 21600000
        s = (_int(_attr(el, "sat"), 0) or 0) / 100000
        lum = (_int(_attr(el, "lum"), 0) or 0) / 100000
        return _rgb_hex(colorsys.hls_to_rgb(h % 1.0, lum, s))
    if tag == "scrgbClr":
        rgb = [(_int(_attr(el, k), 0) or 0) / 100000 for k in ("r", "g", "b")]
        return _rgb_hex(tuple(_to_srgb(c) for c in rgb))
    if tag == "schemeClr":
        val = _attr(el, "val") or ""
        if val == "phClr":
            return ph
        if ctx is None:
            return None
        slot = ctx.clrmap.get(val, val)
        return ctx.theme.colors.get(slot)
    return None


def _clamp(x: float) -> float:
    return 0.0 if x < 0 else 1.0 if x > 1 else x


def _to_linear(c: float) -> float:
    return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4


def _to_srgb(c: float) -> float:
    c = _clamp(c)
    return c * 12.92 if c <= 0.0031308 else 1.055 * c ** (1 / 2.4) - 0.055


def _rgb_hex(rgb: tuple[float, ...]) -> str:
    return "".join(f"{round(_clamp(c) * 255):02X}" for c in rgb)


def _color(
    el: ET.Element | None, ctx: _Ctx | None, ph: str | None = None
) -> str | None:
    """Resolve a color element (srgbClr, schemeClr, ...) to RRGGBB[~][/aN][+mod=val]."""
    if el is None:
        return None
    base = _base_color(el, ctx, ph)
    if base is None:
        return None
    alpha = 100000
    suffix = re.search(r"/a([\d.]+)", base)
    if suffix:
        alpha = round(float(suffix.group(1)) * 1000)
    r, g, b = (int(base[i : i + 2], 16) / 255 for i in (0, 2, 4))
    extra: list[str] = []
    for mod in _kids(el):
        name = _ln(mod.tag)
        val = _int(_attr(mod, "val"), 0) or 0
        if name in {
            "lumMod",
            "lumOff",
            "satMod",
            "satOff",
            "hueMod",
            "hueOff",
            "hue",
            "sat",
            "lum",
        }:
            h, light, s = colorsys.rgb_to_hls(r, g, b)
            if name == "lumMod":
                light *= val / 100000
            elif name == "lumOff":
                light += val / 100000
            elif name == "lum":
                light = val / 100000
            elif name == "satMod":
                s *= val / 100000
            elif name == "satOff":
                s += val / 100000
            elif name == "sat":
                s = val / 100000
            elif name == "hueMod":
                h = (h * val / 100000) % 1.0
            elif name == "hueOff":
                h = (h + val / 21600000) % 1.0
            elif name == "hue":
                h = (val / 21600000) % 1.0
            r, g, b = colorsys.hls_to_rgb(h, _clamp(light), _clamp(s))
        elif name == "tint":
            t = val / 100000
            r, g, b = (_to_srgb(1 - (1 - _to_linear(c)) * t) for c in (r, g, b))
        elif name == "shade":
            t = val / 100000
            r, g, b = (_to_srgb(_to_linear(c) * t) for c in (r, g, b))
        elif name == "alpha":
            alpha = val
        elif name == "alphaMod":
            alpha = alpha * val // 100000
        elif name == "alphaOff":
            alpha += val
        elif name == "inv":
            r, g, b = 1 - r, 1 - g, 1 - b
        elif name == "comp":
            h, light, s = colorsys.rgb_to_hls(r, g, b)
            r, g, b = colorsys.hls_to_rgb((h + 0.5) % 1.0, light, s)
        elif name == "gray":
            y = 0.3 * r + 0.59 * g + 0.11 * b
            r = g = b = y
        elif name == "gamma":
            r, g, b = (_to_srgb(c) for c in (r, g, b))
        elif name == "invGamma":
            r, g, b = (_to_linear(c) for c in (r, g, b))
        else:
            extra.append(f"+{name}={_attr(mod, 'val') or ''}")
    out = _rgb_hex((r, g, b))
    # "~" marks a color computed through modifiers (compared by rendered effect).
    if _kids(el) or "~" in base:
        out += "~"
    alpha = max(0, min(100000, alpha))
    if alpha < 100000:
        out += f"/a{alpha / 1000:g}"
    return out + "".join(extra)


def _color_in(
    container: ET.Element | None, ctx: _Ctx | None, ph: str | None = None
) -> str | None:
    for child in _kids(container):
        if _ln(child.tag) in _COLOR_TAGS:
            return _color(child, ctx, ph)
    return None


def _fill_el(container: ET.Element | None) -> ET.Element | None:
    for child in _kids(container):
        if _ln(child.tag) in _FILL_TAGS:
            return child
    return None


def _fill(
    el: ET.Element | None, ctx: _Ctx, ph: str | None = None
) -> dict[str, Any] | None:
    """Model of a fill element (noFill/solidFill/gradFill/blipFill/pattFill/grpFill)."""
    if el is None:
        return None
    tag = _ln(el.tag)
    if tag == "noFill":
        return {"type": "none"}
    if tag == "grpFill":
        return {"type": "group"}
    if tag == "solidFill":
        return {"type": "solid", "color": _color_in(el, ctx, ph)}
    if tag == "pattFill":
        return {
            "type": "pattern",
            "preset": _attr(el, "prst") or "pct5",
            "fg": _color_in(_child(el, "fgClr"), ctx, ph),
            "bg": _color_in(_child(el, "bgClr"), ctx, ph),
        }
    if tag == "gradFill":
        stops = [
            [_int(_attr(gs, "pos"), 0), _color_in(gs, ctx, ph)]
            for gs in _children(_child(el, "gsLst"), "gs")
        ]
        lin = _child(el, "lin")
        path = _child(el, "path")
        shade: Any = None
        if lin is not None:
            shade = {
                "linear": _int(_attr(lin, "ang"), 0),
                "scaled": _bool(_attr(lin, "scaled"), False),
            }
        elif path is not None:
            shade = {"path": _attr(path, "path") or "circle"}
        return {"type": "gradient", "stops": stops, "shade": shade}
    if tag == "blipFill":
        return {"type": "picture", **_blip(el, ctx)}
    return None


def _blip(blip_fill: ET.Element, ctx: _Ctx) -> dict[str, Any]:
    blip = _child(blip_fill, "blip")
    embed = ctx.pkg.internal(ctx.part, _rattr(blip, "embed"))
    link = ctx.pkg.rel(ctx.part, _rattr(blip, "link"))
    src = _child(blip_fill, "srcRect")
    crop = [_int(_attr(src, k), 0) for k in ("l", "t", "r", "b")]
    tile = _child(blip_fill, "tile")
    return {
        "image": ctx.pkg.blob_sha(embed),
        "link": _external(link[1]) if link and link[2] else None,
        "crop": crop,
        "mode": "tile" if tile is not None else "stretch",
    }


def _style_fill(ref: ET.Element | None, ctx: _Ctx) -> dict[str, Any] | None:
    if ref is None:
        return None
    idx = _int(_attr(ref, "idx"), 0) or 0
    color = _color_in(ref, ctx)
    if idx == 0:
        return {"type": "none"}
    styles, pos = (
        (ctx.theme.bg_fills, idx - 1001) if idx >= 1001 else (ctx.theme.fills, idx - 1)
    )
    if not 0 <= pos < len(styles):
        return None
    return _fill(styles[pos], ctx.at(ctx.theme.part or ctx.part), color)


def _line_parts(
    ln: ET.Element | None, ctx: _Ctx, ph: str | None = None
) -> dict[str, Any]:
    if ln is None:
        return {}
    out: dict[str, Any] = {}
    if _attr(ln, "w") is not None:
        out["width"] = _int(_attr(ln, "w"), 0)
    fill = _fill(_fill_el(ln), ctx, ph)
    if fill is not None:
        out["fill"] = fill
    dash = _child(ln, "prstDash")
    if dash is not None:
        out["dash"] = _attr(dash, "val") or "solid"
    elif _child(ln, "custDash") is not None:
        out["dash"] = "custom"
    for end in ("headEnd", "tailEnd"):
        node = _child(ln, end)
        if node is not None:
            out[end] = _attr(node, "type") or "none"
    return out


def _style_line(ref: ET.Element | None, ctx: _Ctx) -> dict[str, Any]:
    if ref is None:
        return {}
    idx = _int(_attr(ref, "idx"), 0) or 0
    if idx == 0:
        return {"fill": {"type": "none"}}
    if not 0 < idx <= len(ctx.theme.lines):
        return {}
    return _line_parts(
        ctx.theme.lines[idx - 1],
        ctx.at(ctx.theme.part or ctx.part),
        _color_in(ref, ctx),
    )


# ---------------------------------------------------------------------------
# Deck context: masters, layouts, placeholders, text styles
# ---------------------------------------------------------------------------


class _Master:
    def __init__(self, pkg: _Pkg, part: str) -> None:
        self.part = part
        self.root = pkg.xml(part)
        self.theme = _Theme(pkg, pkg.first(part, "theme"))
        self.clrmap = _clrmap(_child(self.root, "clrMap")) or dict(_DEFAULT_CLRMAP)
        csld = _child(self.root, "cSld")
        self.name = _attr(csld, "name") or self.theme.name
        self.placeholders = _placeholders(_child(csld, "spTree"))
        styles = _child(self.root, "txStyles")
        self.title_style = _child(styles, "titleStyle")
        self.body_style = _child(styles, "bodyStyle")
        self.other_style = _child(styles, "otherStyle")


class _Layout:
    def __init__(self, pkg: _Pkg, part: str, master: _Master | None) -> None:
        self.part = part
        self.root = pkg.xml(part)
        self.master = master
        csld = _child(self.root, "cSld")
        self.name = _attr(csld, "name") or ""
        self.placeholders = _placeholders(_child(csld, "spTree"))
        override = _path(self.root, "clrMapOvr", "overrideClrMapping")
        self.clrmap = _clrmap(override) or (
            master.clrmap if master else dict(_DEFAULT_CLRMAP)
        )


class _Deck:
    def __init__(self, pkg: _Pkg, main: str, root: ET.Element) -> None:
        self.pkg = pkg
        self.main = main
        self.root = root
        self.default_style = _child(root, "defaultTextStyle")
        self.masters: list[_Master] = []
        self.master_by_part: dict[str, _Master] = {}
        for sm in _children(_child(root, "sldMasterIdLst"), "sldMasterId"):
            part = pkg.internal(main, _rattr(sm, "id"))
            if part and part not in self.master_by_part and pkg.xml(part) is not None:
                master = _Master(pkg, part)
                self.masters.append(master)
                self.master_by_part[part] = master
        self.layouts: dict[str, _Layout] = {}
        self.slides: list[str] = []
        for sid in _children(_child(root, "sldIdLst"), "sldId"):
            part = pkg.internal(main, _rattr(sid, "id"))
            if part and pkg.xml(part) is not None:
                self.slides.append(part)
        self.slide_index = {part: i for i, part in enumerate(self.slides)}
        self.slide_ids = {
            _attr(sid, "id"): i
            for i, sid in enumerate(
                s
                for s in _children(_child(root, "sldIdLst"), "sldId")
                if pkg.internal(main, _rattr(s, "id")) in self.slide_index
            )
        }

    def master(self, part: str | None) -> _Master | None:
        if part is None or self.pkg.xml(part) is None:
            return None
        if part not in self.master_by_part:
            self.master_by_part[part] = _Master(self.pkg, part)
        return self.master_by_part[part]

    def layout(self, part: str | None) -> _Layout | None:
        if part is None or self.pkg.xml(part) is None:
            return None
        if part not in self.layouts:
            master = self.master(self.pkg.first(part, "slideMaster"))
            self.layouts[part] = _Layout(self.pkg, part, master)
        return self.layouts[part]


_TITLE_TYPES = {"title", "ctrTitle"}
_OTHER_TYPES = {"dt", "ftr", "sldNum", "hdr"}


def _ph_info(shape: ET.Element) -> tuple[str, int] | None:
    nv = next((c for c in _kids(shape) if _ln(c.tag).startswith("nv")), None)
    ph = _path(nv, "nvPr", "ph")
    if ph is None:
        return None
    return (_attr(ph, "type") or "obj", _int(_attr(ph, "idx"), 0) or 0)


def _placeholders(tree: ET.Element | None) -> list[tuple[str, int, ET.Element]]:
    out = []
    for shape in _kids(tree):
        info = _ph_info(shape)
        if info is not None:
            out.append((info[0], info[1], shape))
    return out


def _master_type(ph_type: str) -> str:
    if ph_type in _TITLE_TYPES:
        return "title"
    if ph_type in _OTHER_TYPES:
        return ph_type
    return "body"


def _match_layout(
    info: tuple[str, int], cands: list[tuple[str, int, ET.Element]]
) -> ET.Element | None:
    ph_type, idx = info
    if idx:
        for t, i, el in cands:
            if i == idx:
                return el
    for t, i, el in cands:
        if t == ph_type:
            return el
    if ph_type in _TITLE_TYPES:
        for t, i, el in cands:
            if t in _TITLE_TYPES:
                return el
    return None


def _match_master(
    ph_type: str, cands: list[tuple[str, int, ET.Element]]
) -> ET.Element | None:
    want = _master_type(ph_type)
    for t, i, el in cands:
        if _master_type(t) == want:
            return el
    return None


# ---------------------------------------------------------------------------
# Text
# ---------------------------------------------------------------------------


class _TextChain:
    """Ordered sources of paragraph/run defaults for one text body."""

    def __init__(self, sources: list[tuple[str, Any]]) -> None:
        # ("lst", element with lvlNpPr/defPPr) or ("font", {"font": ..., "color": ...})
        self.sources = sources

    def ppr(self, level: int) -> list[Any]:
        out: list[Any] = []
        for kind, src in self.sources:
            if kind == "font":
                out.append(src)
                continue
            lvl = _child(src, f"lvl{level + 1}pPr")
            if lvl is not None:
                out.append(lvl)
            default = _child(src, "defPPr")
            if default is not None:
                out.append(default)
        return out


def _bullet(ppr: ET.Element | None) -> str | None:
    for child in _kids(ppr):
        tag = _ln(child.tag)
        if tag == "buNone":
            return "none"
        if tag == "buChar":
            return "char:" + _symbol_char(_attr(child, "char") or "")
        if tag == "buAutoNum":
            return f"autonum:{_attr(child, 'type') or ''}:{_int(_attr(child, 'startAt'), 1)}"
        if tag == "buBlip":
            return "picture"
    return None


def _resolve_font(face: str | None, ctx: _Ctx) -> str | None:
    if face is None:
        return None
    if face.startswith("+mj"):
        return ctx.theme.major
    if face.startswith("+mn"):
        return ctx.theme.minor
    return face


def _run_props(rpr: ET.Element | None, chain: list[Any], ctx: _Ctx) -> dict[str, Any]:
    layers: list[Any] = [rpr] + [
        src if isinstance(src, dict) else _child(src, "defRPr") for src in chain
    ]

    def attr(name: str) -> str | None:
        for layer in layers:
            if layer is None or isinstance(layer, dict):
                continue
            value = _attr(layer, name)
            if value is not None:
                return value
        return None

    color = font = None
    for layer in layers:
        if layer is None:
            continue
        if isinstance(layer, dict):
            if color is None and layer.get("color") is not None:
                color = layer["color"]
            if font is None and layer.get("font") is not None:
                font = layer["font"]
            continue
        if color is None:
            fill = _fill_el(layer)
            if fill is not None:
                model = _fill(fill, ctx)
                color = (
                    model.get("color")
                    if model and model["type"] == "solid"
                    else (model or {}).get("type")
                )
        if font is None:
            latin = _child(layer, "latin")
            if latin is not None and _attr(latin, "typeface"):
                font = _resolve_font(_attr(latin, "typeface"), ctx)
        if color is not None and font is not None:
            break
    return {
        "bold": _bool(attr("b"), False),
        "italic": _bool(attr("i"), False),
        "underline": attr("u") or "none",
        "size": _int(attr("sz"), 1800),
        "color": color or ctx.theme.colors.get(ctx.clrmap.get("tx1", "dk1")),
        "font": font or ctx.theme.minor,
    }


def _link_target(hlink: ET.Element | None, ctx: _Ctx) -> str | None:
    if hlink is None:
        return None
    action = _attr(hlink, "action") or ""
    rel = ctx.pkg.rel(ctx.part, _rattr(hlink, "id"))
    target = ""
    if rel is not None:
        kind, dest, external = rel
        if external:
            target = _external(dest)
        elif ctx.deck is not None and dest in ctx.deck.slide_index:
            target = f"slide:{ctx.deck.slide_index[dest] + 1}"
        else:
            target = f"{kind}:{ctx.pkg.blob_sha(dest) or 'missing'}"
    if not action and not target:
        return None
    return f"{action}|{target}" if action else target


def _paragraphs(
    body: ET.Element | None, chain: _TextChain, ctx: _Ctx, links: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    paragraphs = []
    for para in _children(body, "p"):
        ppr = _child(para, "pPr")
        level = _int(_attr(ppr, "lvl"), 0) or 0
        inherited = chain.ppr(level)
        layers = [ppr] + [p for p in inherited if not isinstance(p, dict)]
        align = next(
            (_attr(p, "algn") for p in layers if _attr(p, "algn") is not None), None
        )
        bullet = next(
            (b for b in (_bullet(p) for p in layers) if b is not None), "none"
        )
        runs: list[dict[str, Any]] = []
        keys: list[Any] = []
        for node in _kids(para):
            tag = _ln(node.tag)
            if tag not in {"r", "br", "fld"}:
                continue
            rpr = _child(node, "rPr")
            text = (
                "\n"
                if tag == "br"
                else (_child(node, "t").text or "")
                if _child(node, "t") is not None
                else ""
            )
            if text == "":
                continue
            run = {"text": text, **_run_props(rpr, inherited, ctx)}
            if tag == "fld":
                run["field"] = _attr(node, "type") or ""
            link = _link_target(_child(rpr, "hlinkClick"), ctx)
            key = (
                json.dumps(
                    {k: v for k, v in run.items() if k != "text"}, sort_keys=True
                ),
                link,
            )
            if runs and keys[-1] == key and "field" not in run:
                runs[-1]["text"] += text
            else:
                runs.append(run)
                keys.append(key)
            if link is not None:
                if links and links[-1]["_run"] is runs[-1]:
                    links[-1]["text"] = runs[-1]["text"]
                else:
                    links.append(
                        {"text": runs[-1]["text"], "target": link, "_run": runs[-1]}
                    )
        # A bullet is not drawn on a paragraph without text.
        paragraphs.append(
            {
                "level": level,
                "bullet": bullet if runs else "none",
                "align": align or "l",
                "runs": runs,
            }
        )
    return paragraphs


def _plain(paragraphs: list[dict[str, Any]]) -> str:
    return "\n".join("".join(r["text"] for r in p["runs"]) for p in paragraphs)


# ---------------------------------------------------------------------------
# Shapes
# ---------------------------------------------------------------------------

_SHAPE_TAGS = {"sp", "grpSp", "graphicFrame", "cxnSp", "pic", "contentPart"}
_GRAPHIC_KINDS = {
    "table": "table",
    "chart": "chart",
    "chartex": "chart",
    "diagram": "smartart",
    "ole": "ole",
}


class _Frame:
    """A group's child coordinate space mapped onto the slide."""

    def __init__(
        self,
        box: tuple[float, float, float, float],
        rot: int,
        flip: bool,
        child: tuple[int, int, int, int],
    ) -> None:
        self.box = box  # absolute (x, y, cx, cy) of the group
        self.rot = rot  # absolute rotation, 60000ths of a degree
        self.flip = flip  # absolute horizontal mirror (vertical flips folded into rot)
        self.child = child  # (x, y, cx, cy) child space mapped onto `box`

    def point(self, px: float, py: float) -> tuple[float, float]:
        """Map a point in this group's child space onto the slide."""
        bx, by, bcx, bcy = self.box
        chx, chy, chcx, chcy = self.child
        sx = bcx / chcx if chcx else 1.0
        sy = bcy / chcy if chcy else 1.0
        dx, dy = (px - chx) * sx - bcx / 2, (py - chy) * sy - bcy / 2
        if self.flip:
            dx = -dx
        theta = math.radians(self.rot / 60000)
        cos, sin = math.cos(theta), math.sin(theta)
        return bx + bcx / 2 + dx * cos - dy * sin, by + bcy / 2 + dx * sin + dy * cos


class _ShapeScope:
    """Where shapes live: slide (placeholders inherit) or master/layout (no inheritance)."""

    def __init__(
        self,
        ctx: _Ctx,
        layout: _Layout | None,
        master: _Master | None,
        links: list[dict[str, Any]],
        media: list[dict[str, Any]],
        frame: _Frame | None = None,
    ) -> None:
        self.ctx = ctx
        self.layout = layout
        self.master = master
        self.links = links
        self.media = media
        self.frame = frame

    def inside(self, frame: _Frame) -> _ShapeScope:
        return _ShapeScope(
            self.ctx, self.layout, self.master, self.links, self.media, frame
        )


def _inherited(shape: ET.Element, scope: _ShapeScope) -> list[ET.Element]:
    """Placeholder ancestors of `shape` (nearest first)."""
    info = _ph_info(shape)
    if info is None:
        return []
    chain = []
    layout_ph = _match_layout(info, scope.layout.placeholders) if scope.layout else None
    if layout_ph is not None:
        chain.append(layout_ph)
        info = _ph_info(layout_ph) or info
    if scope.master is not None:
        master_ph = _match_master(info[0], scope.master.placeholders)
        if master_ph is not None:
            chain.append(master_ph)
    return chain


def _sppr(shape: ET.Element) -> ET.Element | None:
    return _child(shape, "spPr", "grpSpPr")


def _xfrm(shape: ET.Element) -> ET.Element | None:
    if _ln(shape.tag) == "graphicFrame":
        return _child(shape, "xfrm")
    return _child(_sppr(shape), "xfrm")


_Raw = tuple[int, int, int, int, int, bool, bool]
_Placed = tuple[float, float, float, float, int, bool]


def _raw_xfrm(xfrm: ET.Element | None) -> _Raw | None:
    if xfrm is None:
        return None
    off, ext = _child(xfrm, "off"), _child(xfrm, "ext")
    return (
        _int(_attr(off, "x"), 0) or 0,
        _int(_attr(off, "y"), 0) or 0,
        _int(_attr(ext, "cx"), 0) or 0,
        _int(_attr(ext, "cy"), 0) or 0,
        _int(_attr(xfrm, "rot"), 0) or 0,
        _bool(_attr(xfrm, "flipH"), False),
        _bool(_attr(xfrm, "flipV"), False),
    )


def _place(raw: _Raw, frame: _Frame | None) -> _Placed:
    """Absolute box, rotation and mirror of a shape; flipV == flipH + 180 degrees."""
    x, y, cx, cy, rot, flip_h, flip_v = raw
    rot += 10800000 if flip_v else 0
    flip = flip_h != flip_v
    px, py, w, h = x + cx / 2, y + cy / 2, float(cx), float(cy)
    if frame is not None:
        chcx, chcy = frame.child[2], frame.child[3]
        sx = frame.box[2] / chcx if chcx else 1.0
        sy = frame.box[3] / chcy if chcy else 1.0
        # A child turned 45-135 or 225-315 degrees scales along swapped axes.
        if 2700000 <= rot % 10800000 < 8100000:
            w, h = w * sy, h * sx
        else:
            w, h = w * sx, h * sy
        px, py = frame.point(px, py)
        if frame.flip:
            rot = -rot
        rot += frame.rot
        flip = flip != frame.flip
    return (px - w / 2, py - h / 2, w, h, rot % 21600000, flip)


def _line_ends(raw: _Raw, frame: _Frame | None) -> dict[str, Any]:
    """Absolute start and end points of a straight line: top-left to bottom-right of
    its unflipped box, then flipped and rotated about the box center."""
    x, y, cx, cy, rot, flip_h, flip_v = raw
    mx, my = x + cx / 2, y + cy / 2
    theta = math.radians(rot / 60000)
    cos, sin = math.cos(theta), math.sin(theta)
    ends = []
    for sign in (-1, 1):
        dx = sign * cx / 2 * (-1 if flip_h else 1)
        dy = sign * cy / 2 * (-1 if flip_v else 1)
        px, py = mx + dx * cos - dy * sin, my + dx * sin + dy * cos
        if frame is not None:
            px, py = frame.point(px, py)
        ends.append([round(px), round(py)])
    return {"start": ends[0], "end": ends[1]}


def _geometry(placed: _Placed | None) -> dict[str, Any] | None:
    if placed is None:
        return None
    x, y, cx, cy, rot, flip = placed
    return {
        "x": round(x),
        "y": round(y),
        "cx": round(cx),
        "cy": round(cy),
        "rot": rot,
        "flip_h": flip,
    }


def _group_frame(xfrm: ET.Element | None, raw: _Raw, placed: _Placed) -> _Frame:
    ch_off, ch_ext = _child(xfrm, "chOff"), _child(xfrm, "chExt")
    child = (
        (_int(_attr(ch_off, "x"), 0) or 0) if ch_off is not None else raw[0],
        (_int(_attr(ch_off, "y"), 0) or 0) if ch_off is not None else raw[1],
        (_int(_attr(ch_ext, "cx"), 0) or 0) if ch_ext is not None else raw[2],
        (_int(_attr(ch_ext, "cy"), 0) or 0) if ch_ext is not None else raw[3],
    )
    return _Frame(placed[:4], placed[4], placed[5], child)


# Presets whose ECMA-376 presetShapeDefinitions geometry is identical.
_PRESET_SYNONYMS = {"straightConnector1": "line"}


def _guide_env(w: float, h: float) -> dict[str, float]:
    env: dict[str, float] = {
        "w": w,
        "h": h,
        "l": 0.0,
        "t": 0.0,
        "r": w,
        "b": h,
        "hc": w / 2,
        "vc": h / 2,
        "ss": min(w, h),
        "ls": max(w, h),
        "cd2": 10800000.0,
        "cd4": 5400000.0,
        "cd8": 2700000.0,
        "3cd4": 16200000.0,
        "3cd8": 8100000.0,
        "5cd8": 13500000.0,
        "7cd8": 18900000.0,
    }
    for n in (2, 3, 4, 5, 6, 8, 10, 12, 16, 32):
        env[f"wd{n}"] = w / n
        env[f"hd{n}"] = h / n
        env[f"ssd{n}"] = min(w, h) / n
    return env


def _operand(token: str | None, env: dict[str, float]) -> float:
    if token is None:
        return 0.0
    try:
        return float(token)
    except ValueError:
        return env.get(token, 0.0)


def _formula(fmla: str, env: dict[str, float]) -> float:
    """Evaluate a DrawingML shape-guide formula (angles in 60000ths of a degree)."""
    op, *args = fmla.split() or [""]
    x, y, z = ([_operand(a, env) for a in args] + [0.0, 0.0, 0.0])[:3]

    def rad(v: float) -> float:
        return math.radians(v / 60000)

    if op == "*/":
        return x * y / z if z else 0.0
    if op == "+-":
        return x + y - z
    if op == "+/":
        return (x + y) / z if z else 0.0
    if op == "?:":
        return y if x > 0 else z
    if op == "abs":
        return abs(x)
    if op == "at2":
        return math.degrees(math.atan2(y, x)) * 60000
    if op == "cat2":
        return x * math.cos(math.atan2(z, y))
    if op == "sat2":
        return x * math.sin(math.atan2(z, y))
    if op == "cos":
        return x * math.cos(rad(y))
    if op == "sin":
        return x * math.sin(rad(y))
    if op == "tan":
        return x * math.tan(rad(y))
    if op == "max":
        return max(x, y)
    if op == "min":
        return min(x, y)
    if op == "mod":
        return math.sqrt(x * x + y * y + z * z)
    if op == "pin":
        return x if y < x else min(y, z)
    if op == "sqrt":
        return math.sqrt(x) if x > 0 else 0.0
    if op == "val":
        return x
    return 0.0


def _eval_guides(container: ET.Element | None, env: dict[str, float]) -> None:
    for lst in ("avLst", "gdLst"):
        for gd in _children(_child(container, lst), "gd"):
            env[_attr(gd, "name") or ""] = _formula(_attr(gd, "fmla") or "", env)


def _preset(sppr: ET.Element | None, raw: _Raw | None, placed: _Placed | None) -> Any:
    """Preset name + evaluated adjust values, or custom paths in EMU of the placed box."""
    w, h = (float(raw[2]), float(raw[3])) if raw else (0.0, 0.0)
    prst = _child(sppr, "prstGeom")
    if prst is not None:
        env = _guide_env(w, h)
        adjust = sorted(
            [_attr(g, "name") or "", round(_formula(_attr(g, "fmla") or "", env))]
            for g in _children(_child(prst, "avLst"), "gd")
        )
        name = _attr(prst, "prst") or ""
        return {"preset": _PRESET_SYNONYMS.get(name, name), "adjust": adjust}
    cust = _child(sppr, "custGeom")
    if cust is None:
        return None
    env = _guide_env(w, h)
    _eval_guides(cust, env)
    abs_w, abs_h = (placed[2], placed[3]) if placed else (w, h)
    paths = []
    for path in _children(_child(cust, "pathLst"), "path"):
        pw = _int(_attr(path, "w"), 0) or w
        ph = _int(_attr(path, "h"), 0) or h
        sx = abs_w / pw if pw else 1.0
        sy = abs_h / ph if ph else 1.0
        commands = []
        for cmd in _kids(path):
            entry: dict[str, Any] = {"op": _ln(cmd.tag)}
            points = [
                [
                    round(_operand(_attr(pt, "x"), env) * sx),
                    round(_operand(_attr(pt, "y"), env) * sy),
                ]
                for pt in _children(cmd, "pt")
            ]
            if points:
                entry["points"] = points
            if entry["op"] == "arcTo":
                entry["wR"] = round(_operand(_attr(cmd, "wR"), env) * sx)
                entry["hR"] = round(_operand(_attr(cmd, "hR"), env) * sy)
                entry["stAng"] = round(_operand(_attr(cmd, "stAng"), env))
                entry["swAng"] = round(_operand(_attr(cmd, "swAng"), env))
            commands.append(entry)
        paths.append(
            {
                "fill": _attr(path, "fill") or "norm",
                "stroke": _bool(_attr(path, "stroke"), True),
                "commands": commands,
            }
        )
    return {"custom": paths}


def _shape(shape: ET.Element, scope: _ShapeScope) -> dict[str, Any]:
    ctx = scope.ctx
    tag = _ln(shape.tag)
    nv = next((c for c in _kids(shape) if _ln(c.tag).startswith("nv")), None)
    cnv = _child(nv, "cNvPr")
    ancestors = _inherited(shape, scope)
    out: dict[str, Any] = {
        "kind": tag,
        "name": _attr(cnv, "name") or "",
        "hidden": _bool(_attr(cnv, "hidden"), False),
        "alt_text": _attr(cnv, "descr") or "",
    }
    info = _ph_info(shape)
    out["placeholder"] = {"type": info[0], "idx": info[1]} if info else None

    xfrm = _xfrm(shape)
    for anc in ancestors:
        if xfrm is not None:
            break
        xfrm = _xfrm(anc)
    raw = _raw_xfrm(xfrm)
    placed = _place(raw, scope.frame) if raw is not None else None
    # A group's own frame is never drawn; its children carry absolute geometry.
    out["geometry"] = None if tag == "grpSp" else _geometry(placed)

    sppr = _sppr(shape)
    preset = None
    if tag in {"sp", "cxnSp", "pic"}:
        preset = _preset(sppr, raw, placed)
        for anc in ancestors:
            if preset is not None:
                break
            preset = _preset(_sppr(anc), raw, placed)
        if preset is None and tag == "pic":
            preset = {"preset": "rect", "adjust": []}
    out["preset"] = preset
    if raw is not None and isinstance(preset, dict) and preset.get("preset") == "line":
        # A straight line is fully described by its endpoints.
        out["geometry"] = _line_ends(raw, scope.frame)

    if tag in {"sp", "cxnSp", "pic", "grpSp"}:
        style = _child(shape, "style")
        fill = _fill(_fill_el(sppr), ctx)
        for anc in ancestors:
            if fill is not None:
                break
            fill = _fill(_fill_el(_sppr(anc)), ctx)
        if fill is None:
            fill = _style_fill(_child(style, "fillRef"), ctx)
        out["fill"] = fill or {"type": "none"}
        line: dict[str, Any] = dict(_style_line(_child(style, "lnRef"), ctx))
        for anc in reversed(ancestors):
            line.update(_line_parts(_child(_sppr(anc), "ln"), ctx))
        line.update(_line_parts(_child(sppr, "ln"), ctx))
        line_fill = line.get("fill", {"type": "none"})
        visible = line_fill["type"] != "none"
        out["line"] = {
            "width": line.get("width", 0) if visible else None,
            "fill": line_fill,
            "dash": line.get("dash", "solid") if visible else None,
            "head": line.get("headEnd", "none") if visible else None,
            "tail": line.get("tailEnd", "none") if visible else None,
        }

    body = _child(shape, "txBody")
    if body is not None:
        out["paragraphs"] = _paragraphs(
            body, _text_chain(shape, body, ancestors, info, scope), ctx, scope.links
        )

    link = _link_target(_child(cnv, "hlinkClick"), ctx)
    if link is not None:
        scope.links.append(
            {"text": _plain(out.get("paragraphs", [])), "target": link, "_run": None}
        )

    if tag == "pic":
        out["picture"] = (
            _blip(_child(shape, "blipFill"), ctx)
            if _child(shape, "blipFill") is not None
            else None
        )
        _collect_media(_path(nv, "nvPr"), ctx, scope.media)
    elif tag == "grpSp":
        inner = scope
        if raw is not None and placed is not None:
            inner = scope.inside(_group_frame(xfrm, raw, placed))
        out["children"] = [
            _shape(c, inner) for c in _kids(shape) if _ln(c.tag) in _SHAPE_TAGS
        ]
    elif tag == "graphicFrame":
        _graphic_frame(shape, out, scope)
    elif tag == "contentPart":
        part = ctx.pkg.internal(ctx.part, _rattr(shape, "id"))
        out["object"] = {
            "content_type": ctx.pkg.content_type(part) if part else None,
            "sha256": ctx.pkg.blob_sha(part),
        }
    return out


def _text_chain(
    shape: ET.Element,
    body: ET.Element,
    ancestors: list[ET.Element],
    info: tuple[str, int] | None,
    scope: _ShapeScope,
) -> _TextChain:
    sources: list[tuple[str, Any]] = []
    own = _child(body, "lstStyle")
    if own is not None:
        sources.append(("lst", own))
    font_ref = _path(shape, "style", "fontRef")
    if font_ref is not None:
        idx = _attr(font_ref, "idx") or "minor"
        sources.append(
            (
                "font",
                {
                    "font": scope.ctx.theme.major
                    if idx == "major"
                    else scope.ctx.theme.minor
                    if idx == "minor"
                    else None,
                    "color": _color_in(font_ref, scope.ctx),
                },
            )
        )
    for anc in ancestors:
        lst = _path(anc, "txBody", "lstStyle")
        if lst is not None:
            sources.append(("lst", lst))
    master = scope.master
    if info is not None and master is not None:
        style = (
            master.title_style
            if info[0] in _TITLE_TYPES
            else master.other_style
            if info[0] in _OTHER_TYPES
            else master.body_style
        )
        if style is not None:
            sources.append(("lst", style))
    deck = scope.ctx.deck
    if deck is not None and deck.default_style is not None:
        sources.append(("lst", deck.default_style))
    return _TextChain(sources)


def _collect_media(
    nvpr: ET.Element | None, ctx: _Ctx, media: list[dict[str, Any]]
) -> None:
    found: dict[tuple[str | None, str | None], str] = {}
    for node in _walk(nvpr):
        tag = _ln(node.tag)
        if tag not in {
            "videoFile",
            "audioFile",
            "quickTimeFile",
            "media",
            "wavAudioFile",
            "audioCd",
        }:
            continue
        kind = {"videoFile": "video", "quickTimeFile": "video", "media": "media"}.get(
            tag, "audio"
        )
        for rid in (_rattr(node, "embed"), _rattr(node, "link")):
            rel = ctx.pkg.rel(ctx.part, rid)
            if rel is None:
                continue
            _, target, external = rel
            key = (
                None if external else ctx.pkg.blob_sha(target),
                _external(target) if external else None,
            )
            # p14:media is an extension of the same clip; the typed element names its kind.
            if key not in found or found[key] == "media":
                found[key] = kind
    for (sha, link), kind in found.items():
        media.append({"kind": kind, "sha256": sha, "link": link})


def _graphic_frame(shape: ET.Element, out: dict[str, Any], scope: _ShapeScope) -> None:
    ctx = scope.ctx
    data = _path(shape, "graphic", "graphicData")
    uri = (_attr(data, "uri") or "").rstrip("/").rsplit("/", 1)[-1]
    kind = _GRAPHIC_KINDS.get(uri, "graphicFrame")
    out["kind"] = kind
    if kind == "table":
        out["table"] = _table(_child(data, "tbl"), scope)
    elif kind == "chart":
        ref = next((c for c in _kids(data) if _ln(c.tag) == "chart"), None)
        part = ctx.pkg.internal(ctx.part, _rattr(ref, "id"))
        out["chart"] = (
            _chart(ctx.pkg, part, ctx) if uri == "chart" else _chartex(ctx.pkg, part)
        )
    elif kind == "smartart":
        out["smartart"] = _smartart(ctx.pkg, ctx.part, _child(data, "relIds"))
    elif kind == "ole":
        out["object"] = _ole(_child(data, "oleObj"), data, ctx)
    else:
        out["object"] = {"uri": _attr(data, "uri") or ""}


def _table(tbl: ET.Element | None, scope: _ShapeScope) -> dict[str, Any] | None:
    if tbl is None:
        return None
    ctx = scope.ctx
    tblpr = _child(tbl, "tblPr")
    flags = sorted(
        k
        for k in (
            "firstRow",
            "firstCol",
            "lastRow",
            "lastCol",
            "bandRow",
            "bandCol",
            "rtl",
        )
        if _bool(_attr(tblpr, k), False)
    )
    style = _child(tblpr, "tableStyleId")
    deck = ctx.deck
    chain = _TextChain(
        [("lst", deck.default_style)] if deck and deck.default_style is not None else []
    )
    rows = []
    for tr in _children(tbl, "tr"):
        cells = []
        for tc in _children(tr, "tc"):
            body = _child(tc, "txBody")
            own = _child(body, "lstStyle")
            cell_chain = _TextChain(
                ([("lst", own)] if own is not None else []) + chain.sources
            )
            merged = "".join(
                m
                for m, k in (("h", "hMerge"), ("v", "vMerge"))
                if _bool(_attr(tc, k), False)
            )
            cells.append(
                {
                    "span": [
                        _int(_attr(tc, "gridSpan"), 1),
                        _int(_attr(tc, "rowSpan"), 1),
                    ],
                    "merged": merged or None,
                    "paragraphs": _paragraphs(body, cell_chain, ctx, scope.links)
                    if body is not None
                    else [],
                    "fill": _fill(_fill_el(_child(tc, "tcPr")), ctx),
                }
            )
        rows.append({"height": _int(_attr(tr, "h"), 0), "cells": cells})
    return {
        "style": (style.text or "").strip().upper() if style is not None else None,
        "flags": flags,
        "columns": [
            _int(_attr(c, "w"), 0) for c in _children(_child(tbl, "tblGrid"), "gridCol")
        ],
        "rows": rows,
    }


def _ole(
    obj: ET.Element | None, frame: ET.Element | None, ctx: _Ctx
) -> dict[str, Any] | None:
    if obj is None:
        return None
    rel = ctx.pkg.rel(ctx.part, _rattr(obj, "id"))
    content_type = sha = link = None
    if rel is not None:
        _, target, external = rel
        if external:
            link = _external(target)
        else:
            content_type = ctx.pkg.content_type(target)
            sha = ctx.pkg.blob_sha(target)
    # The preview image may sit in either mc:AlternateContent branch; any branch counts.
    preview = None
    for pic in frame.iter() if frame is not None else ():
        blip_fill = _child(pic, "blipFill") if _ln(pic.tag) == "pic" else None
        if blip_fill is not None:
            preview = _blip(blip_fill, ctx)["image"]
            break
    if preview is None and _attr(obj, "spid"):
        preview = _vml_preview(ctx, _attr(obj, "spid") or "")
    return {
        "prog_id": _attr(obj, "progId") or "",
        "content_type": content_type,
        "sha256": sha,
        "link": link,
        "preview": preview,
    }


def _vml_preview(ctx: _Ctx, spid: str) -> str | None:
    """Image of the legacy VML shape `spid` (where pre-2010 OLE previews live)."""
    for target, external in ctx.pkg.by_type(ctx.part, "vmlDrawing"):
        raw = None if external else ctx.pkg.data(target)
        if raw is None:
            continue
        try:
            root = ET.fromstring(raw)
        except ET.ParseError:
            continue
        for shape in root.iter():
            ids = {v for k, v in shape.attrib.items() if _ln(k) in {"id", "spid"}}
            if spid not in ids:
                continue
            for node in shape.iter():
                if _ln(node.tag) != "imagedata":
                    continue
                rid = next(
                    (v for k, v in node.attrib.items() if _ln(k) in {"relid", "id"}),
                    None,
                )
                image = ctx.pkg.internal(target, rid)
                if image is not None:
                    return ctx.pkg.blob_sha(image)
    return None


# ---------------------------------------------------------------------------
# Charts and SmartArt
# ---------------------------------------------------------------------------


def _cache(container: ET.Element | None) -> dict[str, Any] | None:
    """Values of a c:tx/c:cat/c:val/... container: formula, format, points (idx-ordered)."""
    if container is None:
        return None
    ref = next(
        (
            c
            for c in _kids(container)
            if _ln(c.tag) in {"strRef", "numRef", "strLit", "numLit", "multiLvlStrRef"}
        ),
        None,
    )
    if ref is None:
        direct = _child(container, "v")
        return (
            {"ref": None, "format": None, "values": [direct.text or ""]}
            if direct is not None
            else None
        )
    numeric = _ln(ref.tag).startswith("num")
    formula = _child(ref, "f")
    cache = next(
        (
            c
            for c in _kids(ref)
            if _ln(c.tag) in {"strCache", "numCache", "multiLvlStrCache"}
        ),
        ref if _ln(ref.tag).endswith("Lit") else None,
    )

    def points(node: ET.Element | None) -> list[Any]:
        count = _int(_attr(_child(node, "ptCount"), "val"), None)
        found: dict[int, str] = {}
        for pt in _children(node, "pt"):
            value = _child(pt, "v")
            found[_int(_attr(pt, "idx"), 0) or 0] = (
                value.text if value is not None and value.text else ""
            )
        size = count if count is not None else (max(found) + 1 if found else 0)
        return [
            (_num(found[i]) if numeric else found[i]) if i in found else None
            for i in range(size)
        ]

    if cache is not None and _ln(cache.tag) == "multiLvlStrCache":
        values: Any = [points(lvl) for lvl in _children(cache, "lvl")]
    else:
        values = points(cache)
    return {
        "ref": (formula.text or "").strip() if formula is not None else None,
        "format": (_child(cache, "formatCode").text or "")
        if _child(cache, "formatCode") is not None
        else None,
        "values": values,
    }


def _chart(pkg: _Pkg, part: str | None, ctx: _Ctx) -> dict[str, Any] | None:
    root = pkg.xml(part)
    if root is None or part is None:
        return None
    chart = _child(root, "chart")
    title = _child(chart, "title")
    title_text = _text_of(_child(title, "tx")) if title is not None else None
    plots = []
    for plot in _kids(_child(chart, "plotArea")):
        tag = _ln(plot.tag)
        if not tag.endswith("Chart"):
            continue
        series = []
        for ser in _children(plot, "ser"):
            name = _cache(_child(ser, "tx"))
            series.append(
                {
                    "name": name["values"][0] if name and name["values"] else None,
                    **{
                        key: _cache(_child(ser, src))
                        for key, src in (
                            ("categories", "cat"),
                            ("values", "val"),
                            ("x", "xVal"),
                            ("y", "yVal"),
                            ("sizes", "bubbleSize"),
                        )
                        if _child(ser, src) is not None
                    },
                }
            )
        plots.append(
            {
                "type": tag,
                "bar_dir": _attr(_child(plot, "barDir"), "val"),
                "grouping": _attr(_child(plot, "grouping"), "val"),
                "series": series,
            }
        )
    legend = _child(chart, "legend")
    workbook = None
    for target, external in pkg.by_type(part, "package"):
        if not external:
            workbook = pkg.blob_sha(target)
    return {
        "title": title_text,
        "auto_title_deleted": _bool(
            _attr(_child(chart, "autoTitleDeleted"), "val"), False
        ),
        "plots": plots,
        "legend": (_attr(_child(legend, "legendPos"), "val") or "r")
        if legend is not None
        else None,
        "workbook": workbook,
    }


def _chartex(pkg: _Pkg, part: str | None) -> dict[str, Any] | None:
    root = pkg.xml(part)
    if root is None or part is None:
        return None
    data = []
    for node in _walk(_child(root, "chartData")):
        if _ln(node.tag) in {"numDim", "strDim"}:
            numeric = _ln(node.tag) == "numDim"
            levels = []
            for lvl in _children(node, "lvl"):
                levels.append(
                    [
                        _num(pt.text) if numeric else (pt.text or "")
                        for pt in _children(lvl, "pt")
                    ]
                )
            f = _child(node, "f")
            data.append(
                {
                    "dim": _attr(node, "type") or "",
                    "ref": (f.text or "").strip() if f is not None else None,
                    "levels": levels,
                }
            )
    series = [
        {"layout": _attr(s, "layoutId") or "", "name": _text_of(_child(s, "tx"))}
        for s in _walk(_child(root, "chart"))
        if _ln(s.tag) == "series"
    ]
    title = _path(root, "chart", "title")
    return {
        "title": _text_of(title) if title is not None else None,
        "series": series,
        "data": data,
    }


def _smartart(
    pkg: _Pkg, source: str, rel_ids: ET.Element | None
) -> dict[str, Any] | None:
    if rel_ids is None:
        return None
    data_part = pkg.internal(source, _rattr(rel_ids, "dm"))
    layout_part = pkg.internal(source, _rattr(rel_ids, "lo"))
    layout = _attr(pkg.xml(layout_part), "uniqueId") if layout_part else None
    root = pkg.xml(data_part)
    points: dict[str, tuple[str, str]] = {}
    order: list[str] = []
    for pt in _children(_child(root, "ptLst"), "pt"):
        pid = _attr(pt, "modelId") or ""
        points[pid] = (
            _attr(pt, "type") or "node",
            "\n".join(_text_of(p) for p in _walk(_child(pt, "t")) if _ln(p.tag) == "p"),
        )
        order.append(pid)
    children: dict[str, list[tuple[int, str]]] = {}
    has_parent: set[str] = set()
    for cxn in _children(_child(root, "cxnLst"), "cxn"):
        if (_attr(cxn, "type") or "parOf") != "parOf":
            continue
        src, dst = _attr(cxn, "srcId") or "", _attr(cxn, "destId") or ""
        children.setdefault(src, []).append((_int(_attr(cxn, "srcOrd"), 0) or 0, dst))
        has_parent.add(dst)

    def node(pid: str, seen: frozenset[str]) -> list[Any]:
        kids = [
            node(c, seen | {pid})
            for _, c in sorted(children.get(pid, []))
            if c in points and c not in seen and points[c][0] in {"node", "asst"}
        ]
        return [points[pid][1], kids]

    roots = [pid for pid in order if points[pid][0] == "doc"]
    if roots:
        tree = node(roots[0], frozenset())[1]
    else:
        tree = [
            node(pid, frozenset())
            for pid in order
            if pid not in has_parent and points[pid][0] in {"node", "asst"}
        ]
    return {"layout": layout, "tree": tree}


# ---------------------------------------------------------------------------
# Slides
# ---------------------------------------------------------------------------


def _background(roots: list[tuple[ET.Element | None, _Ctx]]) -> dict[str, Any] | None:
    for root, ctx in roots:
        bg = _path(root, "cSld", "bg")
        if bg is None:
            continue
        bgpr = _child(bg, "bgPr")
        if bgpr is not None:
            return _fill(_fill_el(bgpr), ctx)
        ref = _child(bg, "bgRef")
        if ref is not None:
            return _style_fill(ref, ctx)
    return None


_NO_TRANSITION = {
    "effect": None,
    "options": [],
    "speed": None,
    "duration": None,
    "advance_on_click": True,
    "advance_after": None,
    "sound": None,
    "stop_sound": False,
}


def _transition(root: ET.Element, ctx: _Ctx) -> dict[str, Any]:
    tr = _child(root, "transition")
    if tr is None:
        return dict(_NO_TRANSITION)
    effect = next((c for c in _kids(tr) if _ln(c.tag) not in {"sndAc", "extLst"}), None)
    sound = None
    snd = next((n for n in _walk(_child(tr, "sndAc")) if _ln(n.tag) == "snd"), None)
    if snd is not None:
        sound = ctx.pkg.blob_sha(ctx.pkg.internal(ctx.part, _rattr(snd, "embed")))
    return {
        "effect": _ln(effect.tag) if effect is not None else None,
        "options": sorted([_ln(k), v] for k, v in effect.attrib.items())
        if effect is not None
        else [],
        # Speed and duration only describe an effect; without one they are unobservable.
        "speed": (_attr(tr, "spd") or "fast") if effect is not None else None,
        "duration": _int(_attr(tr, "dur"), None) if effect is not None else None,
        "advance_on_click": _bool(_attr(tr, "advClick"), True),
        "advance_after": _int(_attr(tr, "advTm"), None),
        "sound": sound,
        "stop_sound": _child(_child(tr, "sndAc"), "endSnd") is not None,
    }


def _animation(root: ET.Element) -> dict[str, Any]:
    timing = _child(root, "timing")
    effects = [
        [
            _attr(n, "presetClass"),
            _int(_attr(n, "presetID"), 0),
            _int(_attr(n, "presetSubtype"), 0),
        ]
        for n in _walk(timing)
        if _ln(n.tag) == "cTn" and _attr(n, "presetClass") is not None
    ]
    present = bool(effects) or any(
        _ln(n.tag) in _ANIM_BEHAVIORS for n in _walk(_child(timing, "tnLst"))
    )
    return {"present": present, "effects": effects}


def _comment_authors(pkg: _Pkg, main: str) -> dict[str, str]:
    authors: dict[str, str] = {}
    for kind, target, external in pkg.rels(main).values():
        if external or kind not in {"commentAuthors", "authors"}:
            continue
        for author in _walk(pkg.xml(target)):
            if _ln(author.tag) in {"cmAuthor", "author"}:
                authors[_attr(author, "id") or ""] = _attr(author, "name") or ""
    return authors


def _comments(pkg: _Pkg, slide: str, authors: dict[str, str]) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for target, external in pkg.by_type(slide, "comments"):
        root = None if external else pkg.xml(target)
        for cm in _children(root, "cm"):
            body = _child(cm, "text")
            text = (
                (body.text or "")
                if body is not None
                else "\n".join(
                    _text_of(p)
                    for p in _walk(_child(cm, "txBody"))
                    if _ln(p.tag) == "p"
                )
            )
            out.append(
                {
                    "author": authors.get(_attr(cm, "authorId") or "", ""),
                    "text": text,
                    "reply": False,
                }
            )
            for reply in _children(_child(cm, "replyLst"), "reply"):
                rtext = "\n".join(
                    _text_of(p)
                    for p in _walk(_child(reply, "txBody"))
                    if _ln(p.tag) == "p"
                )
                out.append(
                    {
                        "author": authors.get(_attr(reply, "authorId") or "", ""),
                        "text": rtext,
                        "reply": True,
                    }
                )
    return out


def _notes(pkg: _Pkg, deck: _Deck, slide: str) -> str | None:
    part = pkg.first(slide, "notesSlide")
    root = pkg.xml(part)
    if root is None or part is None:
        return None
    texts = []
    for shape in _walk(_path(root, "cSld", "spTree")):
        info = _ph_info(shape) if _ln(shape.tag) == "sp" else None
        if info is not None and info[0] == "body":
            body = _child(shape, "txBody")
            texts.append(
                "\n".join(
                    "".join(
                        "\n" if _ln(n.tag) == "br" else _text_of(n)
                        for n in _kids(p)
                        if _ln(n.tag) in {"r", "br", "fld"}
                    )
                    for p in _children(body, "p")
                )
            )
    text = "\n".join(texts)
    return text if text.strip() else None


def _strip_link_refs(links: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [{"text": entry["text"], "target": entry["target"]} for entry in links]


def _slide(deck: _Deck, part: str, authors: dict[str, str]) -> dict[str, Any]:
    pkg = deck.pkg
    root = pkg.xml(part)
    assert root is not None
    layout = deck.layout(pkg.first(part, "slideLayout"))
    master = layout.master if layout else None
    theme = master.theme if master else _Theme(pkg, None)
    override = _clrmap(_path(root, "clrMapOvr", "overrideClrMapping"))
    clrmap = override or (layout.clrmap if layout else dict(_DEFAULT_CLRMAP))
    ctx = _Ctx(pkg, part, theme, clrmap, deck)
    links: list[dict[str, Any]] = []
    media: list[dict[str, Any]] = []
    scope = _ShapeScope(ctx, layout, master, links, media)
    tree = _path(root, "cSld", "spTree")
    shapes = [_shape(s, scope) for s in _kids(tree) if _ln(s.tag) in _SHAPE_TAGS]
    bg_sources: list[tuple[ET.Element | None, _Ctx]] = [(root, ctx)]
    if layout:
        bg_sources.append(
            (layout.root, _Ctx(pkg, layout.part, theme, layout.clrmap, deck))
        )
    if master:
        bg_sources.append(
            (master.root, _Ctx(pkg, master.part, theme, master.clrmap, deck))
        )
    return {
        "hidden": not _bool(_attr(root, "show"), True),
        "layout": layout.name if layout else None,
        "master": master.name if master else None,
        "background": _background(bg_sources),
        "shapes": shapes,
        "hyperlinks": _strip_link_refs(links),
        "notes": _notes(pkg, deck, part),
        "comments": _comments(pkg, part, authors),
        "transition": _transition(root, ctx),
        "animation": _animation(root),
        "media": media,
    }


def _template_shapes(
    deck: _Deck, root: ET.Element | None, ctx: _Ctx, master: _Master | None
) -> list[dict[str, Any]]:
    """Visible non-placeholder shapes of a master or layout."""
    scope = _ShapeScope(ctx, None, master, [], [])
    tree = _path(root, "cSld", "spTree")
    return [
        _shape(s, scope)
        for s in _kids(tree)
        if _ln(s.tag) in _SHAPE_TAGS and _ph_info(s) is None
    ]


def _masters(deck: _Deck) -> list[dict[str, Any]]:
    pkg = deck.pkg
    out = []
    for master in deck.masters:
        ctx = _Ctx(pkg, master.part, master.theme, master.clrmap, deck)
        layouts = []
        for lid in _children(_child(master.root, "sldLayoutIdLst"), "sldLayoutId"):
            layout = deck.layout(pkg.internal(master.part, _rattr(lid, "id")))
            if layout is None:
                continue
            lctx = _Ctx(pkg, layout.part, master.theme, layout.clrmap, deck)
            layouts.append(
                {
                    "name": layout.name,
                    "show_master_shapes": _bool(
                        _attr(layout.root, "showMasterSp"), True
                    ),
                    "shapes": _template_shapes(deck, layout.root, lctx, master),
                }
            )
        layouts.sort(key=lambda item: (item["name"], json.dumps(item, sort_keys=True)))
        out.append(
            {
                "name": master.name,
                "shapes": _template_shapes(deck, master.root, ctx, master),
                "layouts": layouts,
            }
        )
    return out


def _themes(deck: _Deck) -> list[dict[str, Any]]:
    out = []
    for master in deck.masters:
        theme = master.theme
        out.append(
            {
                "name": theme.name,
                "color_scheme": theme.scheme_name,
                "colors": dict(sorted(theme.colors.items())),
                "fonts": {"major": theme.major, "minor": theme.minor},
            }
        )
    return out


def _custom_xml(pkg: _Pkg, main: str) -> list[str]:
    hashes = []
    seen: set[str] = set()
    for source in ("", main):
        for kind, target, external in pkg.rels(source).values():
            if external or kind != "customXml" or target in seen:
                continue
            seen.add(target)
            raw = pkg.data(target)
            if raw is None:
                continue
            try:
                canonical = json.dumps(
                    _canonical(ET.fromstring(raw)), ensure_ascii=False
                )
                hashes.append(_sha(canonical.encode()))
            except ET.ParseError:
                hashes.append(_sha(raw))
    return sorted(hashes)


def _sections(deck: _Deck) -> list[dict[str, Any]]:
    out = []
    for section in _walk(_child(deck.root, "extLst")):
        if _ln(section.tag) != "section":
            continue
        slides = [
            deck.slide_ids[_attr(s, "id")]
            for s in _children(_child(section, "sldIdLst"), "sldId")
            if _attr(s, "id") in deck.slide_ids
        ]
        out.append({"name": _attr(section, "name") or "", "slides": slides})
    return out


def extract(path: pathlib.Path) -> dict[str, Any]:
    """Return the viewer-observable model of the .pptx at `path`."""
    pkg = _Pkg(pathlib.Path(path))
    main = None
    for kind, target, external in pkg.rels("").values():
        if kind == "officeDocument" and not external:
            main = target
    root = pkg.xml(main)
    if main is None or root is None or _ln(root.tag) != "presentation":
        raise ValueError("package has no presentation part")
    deck = _Deck(pkg, main, root)
    size = _child(root, "sldSz")
    authors = _comment_authors(pkg, main)
    return {
        "slide_size": {
            "cx": _int(_attr(size, "cx"), None),
            "cy": _int(_attr(size, "cy"), None),
        },
        "slides": [_slide(deck, part, authors) for part in deck.slides],
        "themes": _themes(deck),
        "masters": _masters(deck),
        "custom_xml": _custom_xml(pkg, main),
        "sections": _sections(deck),
    }


# ---------------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------------

_KEY_FEATURE = {
    "slide_size": "slide-size",
    "slides": "slides",
    "layout": "slide-layout",
    "master": "slide-layout",
    "background": "background",
    "shapes": "shapes",
    "kind": "shapes",
    "name": "shapes",
    "alt_text": "shapes",
    "geometry": "geometry",
    "preset": "geometry",
    "fill": "fill-line",
    "line": "fill-line",
    "placeholder": "placeholders",
    "paragraphs": "text",
    "runs": "text",
    "text": "text",
    "field": "text",
    "level": "text-format",
    "bullet": "text-format",
    "align": "text-format",
    "bold": "text-format",
    "italic": "text-format",
    "underline": "text-format",
    "size": "text-format",
    "color": "text-format",
    "font": "text-format",
    "table": "tables",
    "picture": "pictures",
    "children": "groups",
    "chart": "charts",
    "smartart": "smartart",
    "object": "embedded-objects",
    "hyperlinks": "hyperlinks",
    "notes": "notes",
    "comments": "comments",
    "transition": "transitions",
    "animation": "animations",
    "media": "media",
    "themes": "theme",
    "masters": "masters-layouts",
    "custom_xml": "custom-xml",
    "sections": "sections",
}
# Once inside these, nested keys keep the container's feature.
_STICKY = {
    "slide-size",
    "background",
    "geometry",
    "fill-line",
    "placeholders",
    "tables",
    "pictures",
    "charts",
    "smartart",
    "embedded-objects",
    "hyperlinks",
    "notes",
    "comments",
    "transitions",
    "animations",
    "media",
    "theme",
    "masters-layouts",
    "custom-xml",
    "sections",
}


def _feature_for(current: str, key: str) -> str:
    if current in _STICKY:
        return current
    if key == "hidden":
        return "slide-hidden" if current == "slides" else "shapes"
    return _KEY_FEATURE.get(key, current)


def _signature(value: Any) -> str:
    return json.dumps(value, sort_keys=True, ensure_ascii=False)


def _pointer(key: str) -> str:
    return str(key).replace("~", "~0").replace("/", "~1")


# Model keys whose numbers are presentation lengths (EMU; 12700 per point).
_LENGTH_TOLERANCE = dict.fromkeys(
    ("x", "y", "cx", "cy", "start", "end", "width", "columns", "height", "points")
    + ("wR", "hR"),
    6350,
)
_RELATIVE_TOLERANCE = 1e-9
# Model keys whose numbers are angles (60000ths of a degree); the first set wraps.
_ANGLE_KEYS = {"rot", "linear", "stAng"}
_SWEEP_KEYS = {"swAng"}
_FULL_TURN = 21600000


def _angle_tolerance(model: dict[str, Any]) -> float:
    """Largest rotation (60000ths of a degree) that moves no point of the deck's slide
    by more than 0.5 pt: 0.5 / max(slide width, height in pt) radians."""
    size = model.get("slide_size") or {}
    extent = max(size.get("cx") or 0, size.get("cy") or 0) / 12700
    if extent <= 0:
        return 0.0
    return math.degrees(0.5 / extent) * 60000


# Model keys holding resolved colors (gradient "stops" are [position, color] pairs).
_COLOR_KEYS = {"color", "fg", "bg", "stops"}
_COLOR = re.compile(r"([0-9A-F]{6})(~?)(?:/a([\d.]+))?(.*)", re.DOTALL)
# A modifier delta of at most 100000/510 (1/1000 %) moves an 8-bit channel by at most
# half a level, so after rounding the two resolved channels differ by at most one level.
_COLOR_LEVELS = 1
_ALPHA_PERCENT = 100 / 510


def _number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _equal_color(before: str, after: str) -> bool:
    a, b = _COLOR.fullmatch(before), _COLOR.fullmatch(after)
    if a is None or b is None:
        return before == after
    if not (a.group(2) or b.group(2)):
        return before == after
    if a.group(4) != b.group(4):
        return False
    alpha_a, alpha_b = float(a.group(3) or 100), float(b.group(3) or 100)
    if abs(alpha_a - alpha_b) > _ALPHA_PERCENT:
        return False
    return all(
        abs(int(a.group(1)[i : i + 2], 16) - int(b.group(1)[i : i + 2], 16))
        <= _COLOR_LEVELS
        for i in (0, 2, 4)
    )


def _equal_scalar(before: Any, after: Any, key: str, angle: float) -> bool:
    if _number(before) and _number(after):
        limit = _LENGTH_TOLERANCE.get(key)
        if limit is not None:
            return abs(before - after) <= limit
        if key in _ANGLE_KEYS:
            delta = abs(before - after) % _FULL_TURN
            return min(delta, _FULL_TURN - delta) <= angle
        if key in _SWEEP_KEYS:
            return abs(before - after) <= angle
        return abs(before - after) <= _RELATIVE_TOLERANCE * max(abs(before), abs(after))
    if key in _COLOR_KEYS and isinstance(before, str) and isinstance(after, str):
        return _equal_color(before, after)
    return type(before) is type(after) and before == after


def _diff(
    before: Any,
    after: Any,
    path: str,
    feature: str,
    out: list[dict[str, Any]],
    key: str = "",
    angle: float = 0.0,
) -> None:
    if isinstance(before, dict) and isinstance(after, dict):
        for name in sorted(set(before) | set(after)):
            sub = f"{path}/{_pointer(name)}"
            feat = _feature_for(feature, name)
            if name not in after:
                out.append(
                    {
                        "feature": feat,
                        "path": sub,
                        "kind": "missing",
                        "before": before[name],
                        "after": None,
                    }
                )
            elif name not in before:
                out.append(
                    {
                        "feature": feat,
                        "path": sub,
                        "kind": "added",
                        "before": None,
                        "after": after[name],
                    }
                )
            else:
                _diff(before[name], after[name], sub, feat, out, name, angle)
        return
    if isinstance(before, list) and isinstance(after, list):
        if before == after:
            return
        a = [_signature(v) for v in before]
        b = [_signature(v) for v in after]
        matcher = difflib.SequenceMatcher(None, a, b, autojunk=False)
        for op, i1, i2, j1, j2 in matcher.get_opcodes():
            if op == "equal":
                continue
            paired = min(i2 - i1, j2 - j1)
            for k in range(paired):
                _diff(
                    before[i1 + k],
                    after[j1 + k],
                    f"{path}/{i1 + k}",
                    feature,
                    out,
                    key,
                    angle,
                )
            for k in range(i1 + paired, i2):
                out.append(
                    {
                        "feature": feature,
                        "path": f"{path}/{k}",
                        "kind": "missing",
                        "before": before[k],
                        "after": None,
                    }
                )
            for k in range(j1 + paired, j2):
                out.append(
                    {
                        "feature": feature,
                        "path": f"{path}/{k}",
                        "kind": "added",
                        "before": None,
                        "after": after[k],
                    }
                )
        return
    if not _equal_scalar(before, after, key, angle):
        out.append(
            {
                "feature": feature,
                "path": path,
                "kind": "changed",
                "before": before,
                "after": after,
            }
        )


def compare(before: dict[str, Any], after: dict[str, Any]) -> list[dict[str, Any]]:
    """Differences between two models from `extract`; an empty list means preserved."""
    out: list[dict[str, Any]] = []
    angle = _angle_tolerance(before) or _angle_tolerance(after)
    _diff(before, after, "", "package", out, "", angle)
    return out
