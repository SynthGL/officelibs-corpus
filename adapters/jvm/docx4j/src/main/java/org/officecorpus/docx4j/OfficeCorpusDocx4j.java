package org.officecorpus.docx4j;

import java.io.BufferedReader;
import java.io.InputStream;
import java.io.InputStreamReader;
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
import java.util.Properties;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;

/**
 * officecorpus one-edit round trip for docx4j.
 *
 * <p>Serve protocol (one long-lived process per container, many files): each stdin line is
 * {@code <ext>\t<base64 input>\t<base64 edit JSON>}. For each request the helper writes the input to
 * a private file, opens it with {@code WordprocessingMLPackage.load(File)}, applies the edit through
 * the JAXB content model, saves it with {@code save(File)}, and answers with JSON lines:
 * {@code opened}, {@code saved} or {@code error} (phase {@code open}, {@code edit} or {@code save}),
 * then {@code done} carrying {@code elapsed_ms} (open, edit and save, in process) and the base64
 * output.
 *
 * <p>docx set-run-text: the n-th {@code P} among {@code MainDocumentPart.getContent()} (the w:body
 * children, unwrapped with {@code XmlUtils.unwrap}), then the n-th {@code R} among that paragraph's
 * {@code getContent()} (its direct children, so runs inside hyperlinks, fields, insertions and smart
 * tags are not counted). The run's {@code getContent()} is cleared (w:rPr is held separately by
 * {@code getRPr()} and kept) and one {@code Text} with the marker is added.
 */
public final class OfficeCorpusDocx4j {
    private static final PrintStream OUT = new PrintStream(System.out, false, StandardCharsets.UTF_8);

    private OfficeCorpusDocx4j() {}

    public static void main(String[] args) throws Exception {
        if (args.length == 1 && args[0].equals("--version")) {
            emit("{\"event\":\"version\",\"version\":" + quote(version()) + "}");
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

    private static String version() throws Exception {
        Properties properties = new Properties();
        try (InputStream stream = WordprocessingMLPackage.class.getResourceAsStream(
                "/META-INF/maven/org.docx4j/docx4j-core/pom.properties")) {
            properties.load(stream);
        }
        return properties.getProperty("version");
    }

    private static void roundTrip(String ext, Path source, Path target, String editJson) {
        WordprocessingMLPackage document;
        try {
            if (!ext.equals("docx")) {
                throw new IllegalArgumentException("unsupported format " + ext);
            }
            document = WordprocessingMLPackage.load(source.toFile());
        } catch (Throwable error) {
            error("open", error);
            return;
        }
        emit("{\"event\":\"opened\"}");
        try {
            if (editJson == null) {
                throw new IllegalArgumentException("request carries no edit spec");
            }
            setRunText(document, Json.parse(editJson));
        } catch (Throwable error) {
            error("edit", error);
            return;
        }
        try {
            document.save(target.toFile());
        } catch (Throwable error) {
            error("save", error);
            return;
        }
        emit("{\"event\":\"saved\"}");
    }

    private static void setRunText(WordprocessingMLPackage document, Map<String, Object> edit) {
        if (!"docx".equals(edit.get("format")) || !"set-run-text".equals(edit.get("kind"))) {
            throw new IllegalArgumentException("unsupported edit " + edit.get("format") + "/" + edit.get("kind"));
        }
        List<P> paragraphs = document.getMainDocumentPart().getContent().stream()
                .map(XmlUtils::unwrap)
                .filter(P.class::isInstance)
                .map(P.class::cast)
                .toList();
        P paragraph = paragraphs.get(Json.integer(edit, "paragraph"));
        List<R> runs = paragraph.getContent().stream()
                .map(XmlUtils::unwrap)
                .filter(R.class::isInstance)
                .map(R.class::cast)
                .toList();
        R run = runs.get(Json.integer(edit, "run"));
        ObjectFactory factory = Context.getWmlObjectFactory();
        Text text = factory.createText();
        text.setValue(Json.string(edit, "text"));
        run.getContent().clear();
        run.getContent().add(factory.createRT(text));
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
