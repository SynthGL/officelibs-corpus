package org.officecorpus.poi;

import java.io.BufferedReader;
import java.io.Closeable;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.PrintStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.presentationml.x2006.main.CTShape;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;

/**
 * officecorpus one-edit round trip for Apache POI.
 *
 * <p>Serve protocol (one long-lived process per container, many files): each stdin line is
 * {@code <ext>\t<base64 input>\t<base64 edit JSON>}. For each request the helper writes the input to
 * a private file, opens it with the documented usermodel constructor for the format, applies the
 * edit through the usermodel API, writes it with {@code write(OutputStream)}, and answers with JSON
 * lines: {@code opened}, {@code saved} or {@code error} (phase {@code open}, {@code edit} or
 * {@code save}), then {@code done} carrying {@code elapsed_ms} (open, edit and save, in process) and
 * the base64 output.
 *
 * <p>Edits:
 * <ul>
 *   <li>xlsx set-cell: {@code XSSFWorkbook.getSheet(name)}, {@code getRow}/{@code createRow},
 *       {@code getCell}/{@code createCell}, {@code XSSFCell.setCellValue(String)}.</li>
 *   <li>pptx set-run-text: {@code getSlides()} (presentation order), the first top-level p:sp of the
 *       slide's spTree with the cNvPr id, mapped to its {@code XSLFTextShape} from
 *       {@code XSLFSlide.getShapes()}; {@code getTextParagraphs()} by index; the run is the n-th
 *       direct a:r of that paragraph, mapped to its {@code XSLFTextRun} from {@code getTextRuns()}
 *       (which also lists a:br and a:fld, so it is not indexed directly); {@code XSLFTextRun.setText}.
 *       An a:r holds only a:rPr and a:t, so setText leaves exactly one a:t.</li>
 *   <li>docx set-run-text: the n-th {@code XWPFParagraph} of {@code getBodyElements()}, checked to be
 *       the n-th top-level w:body/w:p; the run is the n-th direct w:r of that paragraph, mapped with
 *       {@code XWPFParagraph.getRun(CTR)}, because {@code getRuns()} also lists runs nested in
 *       w:hyperlink, w:fldSimple, w:ins and w:smartTag. XWPFRun has no API that clears a run's
 *       content, so the run's children other than w:rPr (w:t, w:tab, w:br, w:lastRenderedPageBreak,
 *       ...) are removed from the {@code CTR} POI exposes via {@code getCTR()}, then
 *       {@code XWPFRun.setText(text, 0)} writes the single w:t.</li>
 * </ul>
 */
public final class OfficeCorpusPoi {
    private static final PrintStream OUT = new PrintStream(System.out, false, StandardCharsets.UTF_8);

    private OfficeCorpusPoi() {}

    public static void main(String[] args) throws Exception {
        if (args.length == 1 && args[0].equals("--version")) {
            emit("{\"event\":\"version\",\"version\":" + quote(org.apache.poi.Version.getVersion()) + "}");
            return;
        }
        Path work = Files.createTempDirectory("officecorpus");
        BufferedReader in = new BufferedReader(new InputStreamReader(System.in, StandardCharsets.UTF_8));
        String line;
        while ((line = in.readLine()) != null) {
            if (line.isEmpty()) {
                continue;
            }
            String[] fields = line.split("\t", -1);
            String ext = fields[0];
            Path source = work.resolve("input." + ext);
            Path target = work.resolve("output." + ext);
            Files.deleteIfExists(target);
            Files.write(source, Base64.getDecoder().decode(fields[1]));
            String edit = fields.length > 2
                    ? new String(Base64.getDecoder().decode(fields[2]), StandardCharsets.UTF_8)
                    : null;
            long start = System.nanoTime();
            roundTrip(ext, source, target, edit);
            double elapsed = (System.nanoTime() - start) / 1e6;
            String output = Files.exists(target)
                    ? Base64.getEncoder().encodeToString(Files.readAllBytes(target))
                    : "";
            emit("{\"event\":\"done\",\"elapsed_ms\":" + elapsed + ",\"output\":\"" + output + "\"}");
        }
    }

    private static void roundTrip(String ext, Path source, Path target, String editJson) {
        Closeable document;
        try (InputStream input = Files.newInputStream(source)) {
            document = switch (ext) {
                case "docx" -> new XWPFDocument(input);
                case "pptx" -> new XMLSlideShow(input);
                case "xlsx" -> new XSSFWorkbook(input);
                default -> throw new IllegalArgumentException("unsupported format " + ext);
            };
        } catch (Throwable error) {
            error("open", error);
            return;
        }
        emit("{\"event\":\"opened\"}");
        try (Closeable owned = document) {
            try {
                if (editJson == null) {
                    throw new IllegalArgumentException("request carries no edit spec");
                }
                Map<String, Object> edit = Json.parse(editJson);
                if (!ext.equals(edit.get("format"))) {
                    throw new IllegalArgumentException("edit format " + edit.get("format") + " for ." + ext);
                }
                if (owned instanceof XSSFWorkbook xlsx) {
                    setCell(xlsx, edit);
                } else if (owned instanceof XMLSlideShow pptx) {
                    setRunText(pptx, edit);
                } else {
                    setRunText((XWPFDocument) owned, edit);
                }
            } catch (Throwable error) {
                error("edit", error);
                return;
            }
            try (OutputStream output = Files.newOutputStream(target)) {
                if (owned instanceof XWPFDocument docx) {
                    docx.write(output);
                } else if (owned instanceof XMLSlideShow pptx) {
                    pptx.write(output);
                } else {
                    ((XSSFWorkbook) owned).write(output);
                }
            } catch (Throwable error) {
                error("save", error);
                return;
            }
        } catch (Throwable error) {
            error("save", error);
            return;
        }
        emit("{\"event\":\"saved\"}");
    }

    private static void requireKind(Map<String, Object> edit, String kind) {
        if (!kind.equals(edit.get("kind"))) {
            throw new IllegalArgumentException("unsupported edit kind " + edit.get("kind"));
        }
    }

    private static void setCell(XSSFWorkbook workbook, Map<String, Object> edit) {
        requireKind(edit, "set-cell");
        String name = Json.string(edit, "sheet");
        XSSFSheet sheet = workbook.getSheet(name);
        if (sheet == null) {
            throw new IllegalArgumentException("no sheet named " + name);
        }
        CellReference ref = new CellReference(Json.string(edit, "cell"));
        XSSFRow row = sheet.getRow(ref.getRow());
        if (row == null) {
            row = sheet.createRow(ref.getRow());
        }
        XSSFCell cell = row.getCell(ref.getCol());
        if (cell == null) {
            cell = row.createCell(ref.getCol());
        }
        cell.setCellValue(Json.string(edit, "value"));
    }

    private static void setRunText(XMLSlideShow deck, Map<String, Object> edit) {
        requireKind(edit, "set-run-text");
        XSLFSlide slide = deck.getSlides().get(Json.integer(edit, "slide") - 1);
        long shapeId = Json.integer(edit, "shape_id");
        CTShape sp = null;
        for (CTShape candidate : slide.getXmlObject().getCSld().getSpTree().getSpArray()) {
            if (candidate.getNvSpPr().getCNvPr().getId() == shapeId) {
                sp = candidate;
                break;
            }
        }
        if (sp == null) {
            throw new IllegalArgumentException("no top-level p:sp with id " + shapeId);
        }
        XSLFTextShape shape = null;
        for (XSLFShape candidate : slide.getShapes()) {
            if (candidate.getXmlObject() == sp && candidate instanceof XSLFTextShape text) {
                shape = text;
                break;
            }
        }
        if (shape == null) {
            throw new IllegalStateException("POI exposes no XSLFTextShape for p:sp id " + shapeId);
        }
        XSLFTextParagraph paragraph = shape.getTextParagraphs().get(Json.integer(edit, "paragraph"));
        CTRegularTextRun r = paragraph.getXmlObject().getRArray(Json.integer(edit, "run"));
        for (XSLFTextRun run : paragraph.getTextRuns()) {
            if (run.getXmlObject() == r) {
                run.setText(Json.string(edit, "text"));
                return;
            }
        }
        throw new IllegalStateException("POI exposes no XSLFTextRun for the target a:r");
    }

    private static void setRunText(XWPFDocument document, Map<String, Object> edit) {
        requireKind(edit, "set-run-text");
        int index = Json.integer(edit, "paragraph");
        List<XWPFParagraph> paragraphs = document.getBodyElements().stream()
                .filter(XWPFParagraph.class::isInstance)
                .map(XWPFParagraph.class::cast)
                .toList();
        CTP p = document.getDocument().getBody().getPArray(index);
        XWPFParagraph paragraph = paragraphs.get(index);
        if (paragraph.getCTP() != p) {
            throw new IllegalStateException("getBodyElements() paragraph " + index + " is not w:body/w:p " + index);
        }
        CTR r = p.getRArray(Json.integer(edit, "run"));
        XWPFRun run = paragraph.getRun(r);
        if (run == null) {
            throw new IllegalStateException("POI exposes no XWPFRun for the target w:r");
        }
        for (XmlObject child : run.getCTR().selectPath("./*")) {
            if (!(child instanceof CTRPr)) {
                try (XmlCursor cursor = child.newCursor()) {
                    cursor.removeXml();
                }
            }
        }
        run.setText(Json.string(edit, "text"), 0);
    }

    private static void error(String phase, Throwable error) {
        StringWriter trace = new StringWriter();
        error.printStackTrace(new PrintWriter(trace));
        String text = trace.toString();
        String message = String.valueOf(error.getMessage());
        emit("{\"event\":\"error\",\"phase\":" + quote(phase)
                + ",\"type\":" + quote(error.getClass().getName())
                + ",\"message\":" + quote(message.substring(0, Math.min(2000, message.length())))
                + ",\"traceback\":" + quote(text.substring(Math.max(0, text.length() - 4000))) + "}");
    }

    private static void emit(String json) {
        OUT.print(json);
        OUT.print('\n');
        OUT.flush();
    }

    private static String quote(String value) {
        StringBuilder out = new StringBuilder(value.length() + 2).append('"');
        for (int i = 0; i < value.length(); i++) {
            char c = value.charAt(i);
            switch (c) {
                case '"' -> out.append("\\\"");
                case '\\' -> out.append("\\\\");
                case '\n' -> out.append("\\n");
                case '\r' -> out.append("\\r");
                case '\t' -> out.append("\\t");
                default -> {
                    if (c < 0x20) {
                        out.append(String.format("\\u%04x", (int) c));
                    } else {
                        out.append(c);
                    }
                }
            }
        }
        return out.append('"').toString();
    }

    /** Parser for the flat edit spec: one JSON object whose values are strings or integers. */
    private static final class Json {
        private final String text;
        private int pos;

        private Json(String text) {
            this.text = text;
        }

        static Map<String, Object> parse(String text) {
            Json json = new Json(text);
            Map<String, Object> out = new LinkedHashMap<>();
            json.expect('{');
            if (json.peek() == '}') {
                json.pos++;
                return out;
            }
            while (true) {
                String key = json.string();
                json.expect(':');
                out.put(key, json.peek() == '"' ? json.string() : json.number());
                if (json.peek() == ',') {
                    json.pos++;
                    continue;
                }
                json.expect('}');
                return out;
            }
        }

        static String string(Map<String, Object> edit, String key) {
            if (edit.get(key) instanceof String value) {
                return value;
            }
            throw new IllegalArgumentException("edit field " + key + " must be a string");
        }

        static int integer(Map<String, Object> edit, String key) {
            if (edit.get(key) instanceof Long value) {
                return Math.toIntExact(value);
            }
            throw new IllegalArgumentException("edit field " + key + " must be an integer");
        }

        private char peek() {
            while (pos < text.length() && Character.isWhitespace(text.charAt(pos))) {
                pos++;
            }
            if (pos >= text.length()) {
                throw new IllegalArgumentException("truncated edit JSON");
            }
            return text.charAt(pos);
        }

        private void expect(char c) {
            if (peek() != c) {
                throw new IllegalArgumentException("expected '" + c + "' at " + pos + " in edit JSON");
            }
            pos++;
        }

        private Long number() {
            int start = pos;
            if (peek() == '-') {
                pos++;
            }
            while (pos < text.length() && Character.isDigit(text.charAt(pos))) {
                pos++;
            }
            return Long.parseLong(text.substring(start, pos));
        }

        private String string() {
            expect('"');
            StringBuilder out = new StringBuilder();
            while (true) {
                char c = text.charAt(pos++);
                if (c == '"') {
                    return out.toString();
                }
                if (c != '\\') {
                    out.append(c);
                    continue;
                }
                char e = text.charAt(pos++);
                switch (e) {
                    case 'b' -> out.append('\b');
                    case 'f' -> out.append('\f');
                    case 'n' -> out.append('\n');
                    case 'r' -> out.append('\r');
                    case 't' -> out.append('\t');
                    case 'u' -> {
                        out.append((char) Integer.parseInt(text.substring(pos, pos + 4), 16));
                        pos += 4;
                    }
                    default -> out.append(e);
                }
            }
        }
    }
}
