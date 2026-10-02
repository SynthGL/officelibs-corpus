// Child process for Node.js adapters: open one document, apply one edit, and save it.
// Usage: roundtrip.cjs <adapter> <source> <target> --edit <edit.json>   (or --version <adapter>)
// Protocol matches child_python.py: JSON lines {"event": "opened"|"saved"|"error"|"version"},
// where an error carries phase "open", "edit" or "save".
"use strict";

function emit(payload) {
  process.stdout.write(JSON.stringify(payload) + "\n");
}

function error(phase, err) {
  emit({
    event: "error",
    phase,
    type: (err && err.constructor && err.constructor.name) || typeof err,
    message: String((err && err.message) || err).slice(0, 2000),
    traceback: String((err && err.stack) || "").slice(-4000),
  });
}

class Unsupported extends Error {}

// SheetJS restricts subpath exports, so its version comes from the module itself.
const VERSIONS = {
  exceljs: () => require("exceljs/package.json").version,
  sheetjs: () => require("xlsx").version,
  "pptx-automizer": () => require("pptx-automizer/package.json").version,
};

// pptx-automizer is template-driven: loadRoot() registers the input as the root template (its
// zip is read lazily, so parse failures surface at write()), and write() with no added slides
// keeps the root's existing slides (removeExistingSlides defaults to false).
function automizer(source, target) {
  const path = require("node:path");
  const Automizer = require("pptx-automizer").default;
  const pres = new Automizer({
    templateDir: path.dirname(source),
    outputDir: path.dirname(target),
    verbosity: 0,
  });
  return pres.loadRoot(path.basename(source));
}

// pptx-automizer documents no way to change one run of a slide that is already in the root
// presentation: it "is currently limited to adding things" (modifying a root slide means truncating
// the root and re-adding every slide), and its text modifiers (ModifyTextHelper.setText,
// replaceText) address a whole shape's text body or tag patterns, not a run by paragraph and run
// index. Reaching one a:r would take a custom XmlElement callback, i.e. raw XML patching.
// https://singerla.github.io/pptx-automizer/concepts
const AUTOMIZER_UNSUPPORTED =
  "pptx-automizer cannot edit a slide of the root presentation in place (documented as limited to " +
  "adding slides; modifying requires truncating the root and re-adding every slide), and its text " +
  "modifiers (ModifyTextHelper.setText/replaceText) address a whole shape text body or tag " +
  "patterns, not one run by paragraph/run index; addressing one a:r needs a custom XmlElement " +
  "callback, i.e. raw XML patching";

function requireEdit(edit, format, kind) {
  if (edit.format !== format || edit.kind !== kind) {
    throw new Error(`unsupported edit ${edit.format}/${edit.kind} for ${format}`);
  }
}

function applyEdit(adapter, document, edit) {
  if (adapter === "exceljs") {
    // Worksheet by name, then Cell.value = string (a plain string cell).
    requireEdit(edit, "xlsx", "set-cell");
    const sheet = document.getWorksheet(edit.sheet);
    if (!sheet) {
      throw new Error(`no sheet named ${edit.sheet}`);
    }
    sheet.getCell(edit.cell).value = edit.value;
  } else if (adapter === "sheetjs") {
    // utils.sheet_add_aoa writes the cell and extends the sheet's !ref range.
    requireEdit(edit, "xlsx", "set-cell");
    const sheet = document.Sheets[edit.sheet];
    if (!sheet) {
      throw new Error(`no sheet named ${edit.sheet}`);
    }
    require("xlsx").utils.sheet_add_aoa(sheet, [[edit.value]], { origin: edit.cell });
  } else {
    throw new Unsupported(AUTOMIZER_UNSUPPORTED);
  }
}

function parseArgs(argv) {
  const args = [];
  let editPath = null;
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--edit") {
      editPath = argv[++i];
    } else {
      args.push(argv[i]);
    }
  }
  return { args, editPath };
}

async function main() {
  const { args, editPath } = parseArgs(process.argv.slice(2));
  const [adapter, source, target] = args;
  if (adapter === "--version") {
    emit({ event: "version", version: VERSIONS[source]() });
    return 0;
  }
  let document;
  try {
    if (adapter === "exceljs") {
      const ExcelJS = require("exceljs");
      document = new ExcelJS.Workbook();
      await document.xlsx.readFile(source);
    } else if (adapter === "sheetjs") {
      document = require("xlsx").readFile(source);
    } else if (adapter === "pptx-automizer") {
      document = automizer(source, target);
    } else {
      throw new Error(`unknown adapter ${adapter}`);
    }
  } catch (err) {
    error("open", err);
    return 3;
  }
  emit({ event: "opened" });
  try {
    if (!editPath) {
      throw new Error("missing --edit <edit.json>");
    }
    const edit = JSON.parse(require("node:fs").readFileSync(editPath, "utf8"));
    applyEdit(adapter, document, edit);
  } catch (err) {
    error("edit", err);
    return 6;
  }
  try {
    if (adapter === "exceljs") {
      await document.xlsx.writeFile(target);
    } else if (adapter === "pptx-automizer") {
      await document.write(require("node:path").basename(target));
    } else {
      require("xlsx").writeFile(document, target);
    }
  } catch (err) {
    error("save", err);
    return 4;
  }
  emit({ event: "saved" });
  return 0;
}

main().then(
  (code) => process.exit(code),
  (err) => {
    error("internal", err);
    process.exit(5);
  },
);
