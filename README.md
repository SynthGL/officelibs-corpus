# officelibs-corpus

A real-world Office Open XML corpus (`.xlsx`, `.pptx`, `.docx`) and the OfficeCorpus round-trip lane behind the "real-world round trip" numbers on officelibs.com.

The lane asks one question per library and file: if you open this file, make one small edit, and save it, is the edit there and does everything else survive? The edit is the same for every library, computed in advance from the input alone, so a library cannot pass by copying the file through: it has to parse the document, change it, and write it back. Every library is run the same way, in its own process, against the same frozen holdout set, and scored by the same checks.

Maintained by SynthGL, which also makes WolfXL, WolfPPT, and WolfDocx. Those engines are benchmarked here like every other library, and the holdout split exists so that none of them is scored on files they were developed against.

## Layout

| Path | Contents |
|---|---|
| `corpus/<format>/<source>/` | The corpus files, unmodified upstream bytes |
| `manifest.json` | Every file: SHA-256, size, format, source repo and pinned revision, upstream path, SPDX license, split, detected features. Also every excluded candidate with its reason, the engine scan receipt, and the frozen holdout digest |
| `holdout.txt` | Frozen holdout list (`<sha256>  <path>`) |
| `sources.json` | Pinned upstream sources and the selection quotas |
| `NOTICE`, `LICENSES/` | Attribution and upstream license texts |
| `src/officecorpus/` | Corpus builder and the round-trip runner (`officecorpus` CLI) |
| `src/officecorpus/model/` | Content models (`xlsx.py`, `pptx.py`, `docx.py`) used by the scored `edit-applied` and `content-preserved` checks |
| `src/officecorpus/edits.py` | The edit each file receives and the content model it implies |
| `adapters/node/` | Pinned Node.js adapters (ExcelJS, SheetJS, pptx-automizer) |
| `adapters/jvm/`, `adapters/dotnet/`, `adapters/go/` | Docker helper sources for Apache POI, docx4j, Open XML SDK, and Excelize |
| `adapters/libreoffice/` | Python-UNO script and launcher for the LibreOffice adapter |
| `evidence/<date>/<format>-edit/run.json` | One-edit lane output. `evidence/2026-10-02/<format>/run.json` (no `-edit` suffix) is the earlier no-edit round trip, kept as history and superseded by the one-edit lane |

## Sources and licenses

| Source | Pinned at | License | Formats |
|---|---|---|---|
| [Govdocs1](https://digitalcorpora.org/corpora/file-corpora/files/) (Digital Corpora) | per-format `by_type` archive SHA-256 in `sources.json` | `LicenseRef-Govdocs1-Public` (see `LICENSES/`) | xlsx, pptx, docx |
| [Apache POI](https://github.com/apache/poi) `test-data/` | `9958a2a216c186064e931af57126ef4410562c02` | Apache-2.0 | xlsx, pptx, docx |
| [LibreOffice core](https://github.com/LibreOffice/core) `sc/qa`, `sd/qa`, `sw/qa` | `434f3ee8a5e8a7789de93a567cec274c4f2de6e9` | MPL-2.0 | xlsx, pptx, docx |
| [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl) test data | changeset `c7b9026dab21c7fae28c45c2366c0fb3edc0319f` | MIT | xlsx |
| [python-pptx](https://github.com/scanny/python-pptx) test files | `278b47b1dedd5b46ee84c286e77cdfb0bf4594be` | MIT | pptx |
| [python-docx](https://github.com/python-openxml/python-docx) test files | `e45454602b53e8e572b179ccf1c91093ec9f4ed7` | MIT | docx |

Govdocs1 is the only source of documents written by ordinary users for ordinary purposes; the other sources are test suites, chosen because they exercise features (charts, pivots, SmartArt, tracked changes, embedded objects) that real documents use and simple generators do not. Each file keeps its upstream license. `NOTICE` carries the attributions, including the Apache POI NOTICE.

## Selection

Per format and source, `officecorpus build-corpus`:

1. collects every file with the format's extension (and its macro-enabled variants, so they can be recorded as excluded);
2. excludes, with the reason recorded in `manifest.json` `excluded[]`: encrypted or OLE-wrapped files, macro-enabled files (`.xlsm`/`.pptm`/`.docm` or any `vbaProject` part), files whose upstream name marks them as deliberately damaged or hostile (`corrupt`, `invalid`, `fuzz`, `crash`, ...), corrupt packages (unreadable zip, missing content types or main part, any XML part not well-formed), files over 10 MB, and byte-identical duplicates;
3. greedily picks the eligible file whose detected features add the most rarity-weighted coverage, until the source's quota is met (quotas in `sources.json`, 100 files per format).

Features are detected from the part inventory and part content (for example `xl/pivotTables/`, `ppt/diagrams/data*.xml`, `<w:ins>` in `word/document.xml`). The detector is a heuristic for selection and labeling, not a conformance claim.

Corpus as built (files per source, holdout / dev):

| Format | Govdocs1 | Apache POI | LibreOffice | Library tests | Total |
|---|---|---|---|---|---|
| xlsx | 25 / 0 | 30 / 0 | 30 / 0 | openpyxl 2 / 13 | 87 / 13 |
| pptx | 30 / 0 | 20 / 0 | 30 / 0 | python-pptx 20 / 0 | 100 / 0 |
| docx | 30 / 0 | 20 / 0 | 30 / 0 | python-docx 15 / 5 | 95 / 5 |

## Split rule

A file is `dev` when its SHA-256 matches any file in the working tree, or any blob reachable from any ref in the git history, of the `wolfxl`, `wolfppt`, or `wolfdocx` engine repositories. Everything else is `holdout`. The scan is read-only and hashes only files whose size equals a corpus file's size. `manifest.json` records the scanned engine commits, how many files and blobs were hashed, and for each dev file which repository and where (working tree or history) it was found.

The holdout list is frozen by `manifest.holdout.digest`: SHA-256 over the sorted `<sha256>  <path>` lines in `holdout.txt`. The runner refuses to start if that digest or any file's bytes no longer match the manifest.

## Round-trip lane

```
uv run officecorpus roundtrip --format <xlsx|pptx|docx> --adapter <id> [--adapter <id> ...] --split holdout --output <new dir> [--limit N]
```

`--adapter all` runs every adapter of the format. Each (adapter, file) runs in a fresh child process group with a private copy of the input and a timeout (default 120 s); the group is killed on timeout. Docker adapters are the exception: see below.

### The edit

Every file gets exactly one edit, computed once by `officecorpus.edits.target` from the input package (relationships followed by type, never by part name) and handed to every adapter unchanged. The marker text is `officelibs edit 7f3a`.

| Format | Edit | Target |
|---|---|---|
| xlsx | `{"kind": "set-cell", "sheet", "cell": "A<n>", "value"}`: a new plain string cell | The first worksheet in workbook order (chartsheets, dialogsheets and macrosheets skipped). `n` is the largest row number in its `sheetData` plus 2 (2 when empty), so the cell never existed |
| pptx | `{"kind": "set-run-text", "slide", "shape_id", "paragraph", "run", "text"}`: the run's text becomes the marker, its formatting stays | Slides in presentation order; the first top-level `p:sp` of `p:spTree` (groups are not entered) with a paragraph that has a direct `a:r` whose `a:t` is non-empty; the first such paragraph and run. `slide` is 1-based, `shape_id` is the `p:cNvPr` id, `paragraph` and `run` are 0-based indexes among direct children |
| docx | `{"kind": "set-run-text", "paragraph", "run", "text"}`: the run keeps its `w:rPr` and holds one `w:t` with the marker | The first top-level `w:body/w:p` with a direct `w:r` holding a non-empty `w:t`; the first such run. Indexes count direct children only |

A file without a valid target is unscored for every adapter (outcome `unscored`, reason `no-edit-target`), no adapter runs on it, and the count is reported per adapter as `no_edit_target`.

LibreOffice's UNO model carries no OOXML ids, so its adapter finds the pptx and docx target by position. A wrong guess must never count as a LibreOffice preservation failure. The result is unscored for LibreOffice (reason `locator mismatch`) when its lookup fails, when the pptx shape at the target position does not carry the target shape's name, or when the located paragraph's text differs from the input model's target paragraph text. That comparison ignores object placeholders (U+FFFC) and treats `\v`, `\r`, and U+2028 as `\n` and `\f` as nothing, because the models and UNO spell breaks differently. The summary reports `locator_mismatch` and `locator_mismatch_files`, and these files are left out of LibreOffice's denominators. Python, Node, and LibreOffice adapters receive the spec as a JSON file (`--edit <path>`); Docker helpers receive it base64-encoded on their request line. Each adapter makes the edit through the library's public API for that operation (set a cell value, set a run's text). A library that cannot express the edit is listed as unsupported with the concrete reason; no adapter patches XML on a library's behalf.

For xlsx, LibreOffice's adapter finds the target sheet by name, in this order: the exact name; else the one sheet whose name matches case-insensitively (Excel sheet names are unique case-insensitively); else the sheet at the target's position in workbook order, used only when UNO holds as many sheets as the workbook. The position step exists because LibreOffice's importer can rename a sheet (in this corpus it reads `sheet1` as `sheet1_2`), and a user editing that workbook in LibreOffice edits that sheet. The rename itself is LibreOffice's behavior and stays a scored difference in its output; it is not a locator mismatch. Several case-insensitive matches or a different sheet count is an adapter error. Each LibreOffice result records how the sheet was found in `samples[0].details.locator.sheet_method` (`name`, `name-case-insensitive`, or `position`).

### Adapters

| Format | Adapter id | Library | Call |
|---|---|---|---|
| xlsx | `openpyxl` | openpyxl 3.1.5 | `load_workbook(path)`, `wb[sheet][cell] = value`, `save(out)` |
| xlsx | `wolfxl` | WolfXL, latest PyPI release | `load_workbook(path, modify=True)` (the documented mode for editing existing workbooks), `wb[sheet][cell] = value`, `save(out)` |
| xlsx | `exceljs` | ExcelJS 4.4.0 | `workbook.xlsx.readFile`, `getWorksheet(name).getCell(ref).value = string`, `workbook.xlsx.writeFile` |
| xlsx | `sheetjs` | SheetJS CE 0.20.3 | `XLSX.readFile`, `utils.sheet_add_aoa(sheet, [[string]], {origin: ref})`, `XLSX.writeFile`, default options |
| xlsx | `apache-poi` | Apache POI 5.5.1 (Docker) | `new XSSFWorkbook(InputStream)`, `getSheet(name)`, `getRow`/`createRow`, `getCell`/`createCell`, `setCellValue(String)`, `write(OutputStream)` |
| xlsx | `excelize` | Excelize 2.11.0 (Docker) | `excelize.OpenFile(path)` with default options, `SetCellStr(sheet, ref, string)`, `SaveAs(out)`, `Close()` |
| xlsx | `aspose-cells-foss` | Aspose.Cells FOSS for Python 26.7.0 | `Workbook(path)`, `get_worksheet_by_name(sheet).cells[cell].put_value(value)`, `save(out)` |
| xlsx | `libreoffice` | LibreOffice (installed) | Python-UNO script in headless `soffice` with a fresh profile per file: the target sheet from `Sheets` (lookup order above), `getCellRangeByName(cell).setString(value)`, store with the `Calc MS Excel 2007 XML` filter |
| pptx | `python-pptx` | python-pptx 1.0.2 | `Presentation(path)`, the shape with `shape_id` in `slides[i].shapes`, `text_frame.paragraphs[p].runs[r].text = value`, `save(out)` |
| pptx | `wolfppt-wheel` | WolfPPT, latest PyPI release | same python-pptx-compatible calls as `python-pptx` |
| pptx | `apache-poi` | Apache POI 5.5.1 (Docker) | `new XMLSlideShow(InputStream)`, the `XSLFTextShape` with the `cNvPr` id, `getTextParagraphs()`, the n-th direct `XSLFTextRun`, `setText(String)`, `write(OutputStream)` |
| pptx | `open-xml-sdk` | Open XML SDK 3.5.1 (Docker) | `PresentationDocument.Open(stream, isEditable: true)`, the `Shape` with the `NonVisualDrawingProperties.Id`, paragraph and direct `Run` by index, run children except `RunProperties` replaced by one `Text`, `Save()` |
| pptx | `libreoffice` | LibreOffice (installed) | Python-UNO script as above. UNO exposes no `cNvPr` id, so the shape is the one at the target's position among the slide's top-level `spTree` shapes (its name is cross-checked); the paragraph comes from the shape's text enumeration and the run is the run's character range, replaced through a text cursor's `setString`; stored with the `Impress MS PowerPoint 2007 XML` filter |
| pptx | `pptx-automizer` | pptx-automizer 0.9.4 | unsupported: it cannot edit a slide of the root presentation in place (it adds slides; modifying requires truncating the root and re-adding every slide), and its text modifiers address a whole shape text body or tag patterns, not one run, so addressing one `a:r` would need raw XML patching |
| docx | `python-docx` | python-docx 1.2.0 | `Document(path)`, `paragraphs[p].runs[r].text = value` (direct `w:p` and `w:r` children), `save(out)` |
| docx | `docxtpl` | docxtpl 0.20.2 | `DocxTemplate(path)`, then `get_docx()`, docxtpl's documented handle on the base document; it is a python-docx `Document`, so the edit is python-docx's `runs[r].text = value`; then `render({})` and `save(out)` (`save()` without `render()` reloads the template from disk and drops edits made through `get_docx()`) |
| docx | `wolfdocx-wheel` | WolfDocx | wheel built from a `git archive` of the local source HEAD into an isolated environment; the documented CLI: `wolfdocx author report-targets` for the target paragraph's structure fingerprint, then `wolfdocx author report` with one text edit (paragraph = the target's 1-based index among the part's `w:p`, `expected_old_text` = the run's text, which WolfDocx requires to occur once in the paragraph) |
| docx | `apache-poi` | Apache POI 5.5.1 (Docker) | `new XWPFDocument(InputStream)`, the n-th body `XWPFParagraph`, the run of the n-th direct `w:r`, run content except `w:rPr` cleared, `setText(String, 0)`, `write(OutputStream)` |
| docx | `docx4j` | docx4j 17.2.1 (Docker) | `WordprocessingMLPackage.load(File)`, the n-th `P` of the main document content, its n-th direct `R`, content cleared and one `Text` added (`rPr` kept), `save(File)` |
| docx | `open-xml-sdk` | Open XML SDK 3.5.1 (Docker) | `WordprocessingDocument.Open(stream, isEditable: true)`, `Body.Elements<Paragraph>()` and direct `Elements<Run>()` by index, run children except `RunProperties` replaced by one `Text`, `Save()` |
| docx | `libreoffice` | LibreOffice (installed) | Python-UNO script as above: the target's paragraph in the body text enumeration (tables skipped) and the run's character range, replaced through a text cursor with change tracking off; stored with the `MS Word 2007 XML` filter. Writer splits a paragraph at a page or column break and drops empty section-break paragraphs, so the paragraph ordinal is mapped through those rules and checked against the paragraph's text (if it differs, the nearest paragraph with exactly that text is used); the run's text is verified at its offset before it is replaced |
| docx | `pandoc` | pandoc (installed) | unsupported: pandoc's AST has no run or paragraph identity. The docx reader dissolves runs into word-level `Str`/`Space` inlines grouped by formatting and drops empty paragraphs, so a filter cannot address the target run |

Python adapters each get a fresh uv environment under `.envs/<adapter>`; `run.json` records the installed package set, interpreter, and for WolfPPT and WolfDocx the source revision and wheel SHA-256. The LibreOffice adapter's receipt carries the SHA-256 of its UNO script and launcher.

Docker adapters run a small helper (`adapters/jvm/apache-poi`, `adapters/jvm/docx4j`, `adapters/dotnet/open-xml-sdk`, `adapters/go/excelize`) in an image tagged `officecorpus/<library>:<version>` on the Docker context named by `OFFICECORPUS_DOCKER_CONTEXT` (default: the current context). The image carries a label with the SHA-256 of its helper source and is rebuilt when that label does not match the checkout. One helper process per container serves many files over stdin/stdout, each with a private input copy; a timeout or crash kills the container and the next file starts a fresh one. `run.json` records the image id, platform, labels, helper source SHA-256, and Docker server OS/arch.

Not included, by policy: pandas, pyexcel, tablib, and pylightxl. They read cell values into their own data structures and write a new workbook from those values; they do not open, edit, and save the file, so the lane would measure a rebuild rather than preservation.

### Scoring

Scored checks:

| Check | Passes when |
|---|---|
| `opened` | The library loaded the file without an exception |
| `saved` | The library applied the edit and wrote a non-empty output without an exception (an edit failure is recorded with its error and fails this check) |
| `package-valid` | The output is a readable zip with parseable `[Content_Types].xml`, a content type for every part, a resolvable officeDocument relationship, and no dangling internal relationship that the input did not already have |
| `edit-applied` | The output's content model shows the marker at the target: xlsx, the target cell is the marker string; pptx and docx, the paragraph that carries the marker in the expected model has the same text in the output, at the same model path or, if earlier content moved, elsewhere on the same slide or in the body (`found_at` records where) |
| `content-preserved` | The content model of the output equals the expected model: the input's model with the edit applied (below) |

A file is **fully preserved** when every scored check passes.

Unscored diagnostics, recorded in `run.json` with full detail but never ranked on:

| Check | Category | Records |
|---|---|---|
| `render-similarity` | `render` | LibreOffice PDF renders of input and output have the same page count and every compared page is within tolerance (below). Unscored because the render oracle is LibreOffice, which is also a library under test here and renders its own output. The edit changes a few words, so a correct output differs slightly from the input render |
| `parts-preserved` | `byte-only` | Every input part has an output counterpart: same name, same relationship path from the package root with the same content type, or identical bytes under a new name |
| `semantic-preserved` | `byte-only` | Every mapped part is equal after XML canonicalization (namespace prefixes, attribute order, whitespace-only text, and relationship ids normalized). The edited part always differs |

The byte-only checks compare how parts are spelled, not what they contain: a library that rewrites a document into an equivalent form (explicit defaults, reordered elements, inline instead of shared strings, renamed parts) fails them. That is why nothing ranks on them.

**Content model.** `officecorpus.model.<format>` turns a package into a model of what a user of the document can observe, and `compare(before, after)` lists every difference, each attributed to one of the format's features (`FEATURES` in the module). Part names, relationship ids, namespace prefixes, attribute order, XML layout, and whether a default value is written or omitted never enter the model, so an equivalent rewrite scores the same as a minimal patch. The exact equivalence rules (defaults, inheritance, colors, numbers, formulas, ranges, binary parts, what is and is not modeled) are the docstring of each module:

- `src/officecorpus/model/xlsx.py`: workbook, defined names, cells, formulas, cell styles, merged cells, columns, rows, views, auto filters, data validations, conditional formats, hyperlinks, comments, protection, page setup, images, charts, tables, pivot tables, external links, VBA, theme;
- `src/officecorpus/model/pptx.py`: slide size, slides, hidden slides, layouts, backgrounds, shapes, geometry, fill and line, placeholders, text and text formatting, tables, pictures, groups, charts, SmartArt, embedded objects, hyperlinks, notes, comments, transitions, animations, media, theme, masters and layouts, custom XML, sections;
- `src/officecorpus/model/docx.py`: body, sections, headers and footers, footnotes, endnotes, comments, tracked changes, fields, bookmarks, hyperlinks, content controls, images, shapes, text boxes, embedded objects, styles, custom and core properties, custom XML, glossary.

**Expected model.** xlsx: the input model with the new string cell added under `cells`. The edit sets a value and specifies no style; the new cell is expected to carry the style Excel gives it, the row or column style an absent cell already shows there (the workbook default where neither row nor column is styled). Any other style on the new cell, including the workbook default under a styled row or column (what writing the cell without a style index gives), is a `cell_styles` difference and fails `content-preserved`. pptx and docx: the model extracted from a copy of the input whose target run holds only its properties and one text element with the marker. The models merge adjacent runs with equal formatting and derive other features (hyperlink text, comment anchors) from run text, so only the model's own extraction knows every observable effect of the edit.

**New-cell default style (xlsx, unscored diagnostic).** Each xlsx result records in `content-preserved` `new_cell_default_style`: true when the new cell carries the workbook default style (`cellXfs[0]`) where the inherited row or column style differs from it. Each adapter's summary counts those files (`new_cell_default_style`, `new_cell_default_style_files`). It names the cause of that `cell_styles` difference and changes no outcome. An earlier version of the lane accepted either style for the new cell. In the first 2026-10-02 xlsx one-edit run that rule turned 22 failures into passes, all of them for WolfXL (which SynthGL maintains); openpyxl (24 files) and Aspose.Cells FOSS (8 files) also wrote the default style, but those files failed on other differences, and the other libraries wrote the inherited style, which is what Excel does. A rule whose only effect on outcomes favored one library is not neutral, so it was removed from scoring and kept only as this diagnostic.

**Formula cached values (xlsx).** A spreadsheet stores each formula's last computed result next to the formula. Excel recomputes formulas on open, so a missing or stale cached value is invisible there; readers that do not recalculate (pandas, file previewers, search indexers) show the cached value, or nothing when it is gone. A difference in a cell whose input cell holds a formula is therefore attributed to its own scored feature, `formula_cached_values`, instead of `cells`: it still fails `content-preserved`, and the per-feature summary shows whether a library loses data or only cached results. The model itself is unchanged; only the attribution moves.

The input and expected models are extracted once per file, in a child process with a timeout (`--content-timeout`, default 300 s). An input whose model cannot be extracted is excluded from `edit-applied` and `content-preserved` with the reason recorded in `run.json` `content_inputs.excluded[]`. An output whose model cannot be extracted fails both.

`content-preserved` detail per file: `differing_features` (feature to number of differences), `input_features` (features present in the input), and `sample` (the first 10 differences, values cut to 200 characters). A feature is present in an input when `compare` attributes at least one difference to it after every scalar of the input model that is not null, false, or the empty string is replaced by a probe value and every non-empty list gains a probe item; presence therefore uses exactly the attribution that scoring uses. Every result records the edit spec it was given (`edit`).

**Summary.** Per adapter, `run.json` `summary[]` reports files total, `no_edit_target`, `locator_mismatch`, files attempted, opened, saved, package-valid, `edit_applied`, content preserved, and fully preserved (the last three over the files in the content lane), plus per feature `feature_preservation[feature] = {preserved, present, rate}`: `present` counts files whose input has the feature, `preserved` those of them whose output shows no difference for it. A file the adapter could not open, edit, or save, or whose output model could not be extracted, preserves none of its features. `files_differing_by_feature` counts files with at least one difference per feature; xlsx summaries add the unscored `new_cell_default_style` count and files (above); `diagnostics` holds the unscored render and byte-only counts and how many outputs were byte-identical to the input (always a failure of `edit-applied`).

**Render tolerance** (diagnostic): pages are rasterized with pymupdf at 50 dpi in RGB. A pixel differs when any channel differs by more than 32 of 255. A page passes when at most 1% of its pixels differ. Up to the first 20 pages are compared; the page-count check covers the whole document.

## Reproduce

```
uv sync
# Rebuild the corpus from pinned sources (optional; the corpus is committed)
uv run officecorpus fetch-sources
uv run officecorpus build-corpus --engines-root <directory containing wolfxl, wolfppt, wolfdocx>
# Run the lane
uv run officecorpus roundtrip --format xlsx --adapter all --split holdout --output evidence/<date>/xlsx-edit
uv run officecorpus roundtrip --format pptx --adapter all --split holdout --output evidence/<date>/pptx-edit
uv run officecorpus roundtrip --format docx --adapter all --split holdout --output evidence/<date>/docx-edit
```

The committed one-edit evidence is `evidence/2026-10-02/{xlsx,pptx,docx}-edit/run.json`. `evidence/2026-10-02/{xlsx,pptx,docx}/run.json` are the earlier no-edit round trip (open and save only), kept as history; the one-edit lane supersedes them, because a byte copy could pass the no-edit lane and those runs predate the current runner.

Requirements: uv, Node.js with npm (for `npm ci` in `adapters/node`), LibreOffice (`soffice`), a Rust toolchain to build the WolfPPT wheel, and Docker (any context; images build from `adapters/<runtime>/<library>/`). The exact command, tool versions, LibreOffice version, the content model's and edit module's source SHA-256, and a SHA-256 of the runner source (including the format's content model and the LibreOffice scripts) are written into each `run.json`.

## Limits

- The content model is the measure. It models what a viewer can observe, not everything a file can carry; each model docstring lists what is not modeled (for example text body insets and shape effects in pptx). A library can lose an unmodeled property and still score as preserved, and a difference the model reports may be invisible in a particular Office application.
- `render-similarity` is a diagnostic only. Its oracle is LibreOffice, not Microsoft Office, and LibreOffice is also a library under test that renders its own output, so a render score would favor the `libreoffice` adapter.
- The edit is deliberately small: one new cell or one run's text. It proves a library parses and re-serializes the document and that an edit through its public API leaves everything else intact; it does not exercise structural edits (inserting rows, slides, or tables). Those are measured by the template-mutation lanes of ExcelBench, PPTBench, and DocxBench.
- Files without a target (pptx decks with no text in a top-level shape, docx files whose top-level paragraphs carry no direct text run) are unscored for every library.
- Library test suites are over-represented relative to real-world use because Govdocs1 has few OOXML files (it was collected in 2009).
- Timings in `run.json` are wall-clock times for one cold child process with up to four files in flight (Docker adapters: in-process open plus save time inside a warm helper on the Docker host); they are not a performance benchmark.
- Govdocs1 redistribution relies on Digital Corpora's statement that the corpus may be freely redistributed to the best of their knowledge; see `LICENSES/LicenseRef-Govdocs1-Public.txt`.
