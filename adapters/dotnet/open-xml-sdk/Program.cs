// officecorpus one-edit round trip for the Open XML SDK (DocumentFormat.OpenXml).
//
// Serve protocol (one long-lived process per container, many files): each stdin line is
// "<ext>\t<base64 input>\t<base64 edit JSON>". The helper copies the input into an expandable
// MemoryStream, opens it editable with WordprocessingDocument.Open(stream, true) or
// PresentationDocument.Open(stream, true), applies the edit to the typed DOM, calls Save(), disposes
// the package, and writes the stream to the output. It answers with JSON lines: opened, saved or
// error (phase open, edit or save), then done carrying elapsed_ms (open, edit and save, in process)
// and the base64 output.
//
// Edits (set-run-text): docx takes Body.Elements<Paragraph>() (direct w:p children) by index, then
// that paragraph's direct Elements<Run>() by index. pptx takes PresentationPart.Presentation
// .SlideIdList in order, resolves the SlidePart with GetPartById, and takes the first direct Shape of
// the ShapeTree whose NonVisualDrawingProperties.Id matches, then TextBody.Elements<A.Paragraph>()
// and that paragraph's direct Elements<A.Run>(). In both, every run child except the run properties
// is removed and one Text element holding the marker is appended.
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using W = DocumentFormat.OpenXml.Wordprocessing;

internal static class Program
{
    private static readonly Stream Stdout = Console.OpenStandardOutput();
    private static readonly JsonSerializerOptions Json = new() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

    private static int Main(string[] args)
    {
        if (args.Length == 1 && args[0] == "--version")
        {
            var version = typeof(OpenXmlPackage).Assembly.GetName().Version!;
            Emit(new { @event = "version", version = $"{version.Major}.{version.Minor}.{version.Build}" });
            return 0;
        }
        var work = Directory.CreateTempSubdirectory("officecorpus").FullName;
        using var reader = new StreamReader(Console.OpenStandardInput(), Encoding.UTF8);
        string? line;
        while ((line = reader.ReadLine()) != null)
        {
            if (line.Length == 0)
            {
                continue;
            }
            var fields = line.Split('\t');
            var ext = fields[0];
            var source = Path.Combine(work, "input." + ext);
            var target = Path.Combine(work, "output." + ext);
            File.Delete(target);
            File.WriteAllBytes(source, Convert.FromBase64String(fields[1]));
            var edit = fields.Length > 2 ? Encoding.UTF8.GetString(Convert.FromBase64String(fields[2])) : null;
            var clock = Stopwatch.StartNew();
            RoundTrip(ext, source, target, edit);
            var elapsed = clock.Elapsed.TotalMilliseconds;
            var output = File.Exists(target) ? Convert.ToBase64String(File.ReadAllBytes(target)) : "";
            Emit(new { @event = "done", elapsed_ms = elapsed, output });
        }
        return 0;
    }

    private static void RoundTrip(string ext, string source, string target, string? editJson)
    {
        var buffer = new MemoryStream();
        OpenXmlPackage document;
        try
        {
            using (var input = File.OpenRead(source))
            {
                input.CopyTo(buffer);
            }
            buffer.Position = 0;
            switch (ext)
            {
                case "docx":
                    var word = WordprocessingDocument.Open(buffer, true);
                    document = word;
                    _ = word.MainDocumentPart!.Document;
                    break;
                case "pptx":
                    var deck = PresentationDocument.Open(buffer, true);
                    document = deck;
                    _ = deck.PresentationPart!.Presentation;
                    break;
                default:
                    throw new ArgumentException($"unsupported format {ext}");
            }
        }
        catch (Exception error)
        {
            Error("open", error);
            return;
        }
        Emit(new { @event = "opened" });
        try
        {
            if (editJson == null)
            {
                throw new ArgumentException("request carries no edit spec");
            }
            using var edit = JsonDocument.Parse(editJson);
            var spec = edit.RootElement;
            if (spec.GetProperty("format").GetString() != ext || spec.GetProperty("kind").GetString() != "set-run-text")
            {
                throw new ArgumentException($"unsupported edit {spec.GetProperty("format")}/{spec.GetProperty("kind")} for .{ext}");
            }
            if (document is WordprocessingDocument word)
            {
                SetRunText(word, spec);
            }
            else
            {
                SetRunText((PresentationDocument)document, spec);
            }
        }
        catch (Exception error)
        {
            document.Dispose();
            Error("edit", error);
            return;
        }
        try
        {
            using (document)
            {
                document.Save();
            }
            File.WriteAllBytes(target, buffer.ToArray());
        }
        catch (Exception error)
        {
            Error("save", error);
            return;
        }
        Emit(new { @event = "saved" });
    }

    private static void SetRunText(WordprocessingDocument word, JsonElement spec)
    {
        var paragraph = word.MainDocumentPart!.Document.Body!.Elements<W.Paragraph>()
            .ElementAt(spec.GetProperty("paragraph").GetInt32());
        var run = paragraph.Elements<W.Run>().ElementAt(spec.GetProperty("run").GetInt32());
        ReplaceContent<W.RunProperties>(run, new W.Text(spec.GetProperty("text").GetString()!));
    }

    private static void SetRunText(PresentationDocument deck, JsonElement spec)
    {
        var presentation = deck.PresentationPart!;
        var slideId = presentation.Presentation.SlideIdList!.Elements<P.SlideId>()
            .ElementAt(spec.GetProperty("slide").GetInt32() - 1);
        var slide = (SlidePart)presentation.GetPartById(slideId.RelationshipId!.Value!);
        var shapeId = spec.GetProperty("shape_id").GetUInt32();
        var shape = slide.Slide.CommonSlideData!.ShapeTree!.Elements<P.Shape>()
            .FirstOrDefault(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value == shapeId)
            ?? throw new ArgumentException($"no top-level p:sp with id {shapeId}");
        var paragraph = shape.TextBody!.Elements<A.Paragraph>().ElementAt(spec.GetProperty("paragraph").GetInt32());
        var run = paragraph.Elements<A.Run>().ElementAt(spec.GetProperty("run").GetInt32());
        ReplaceContent<A.RunProperties>(run, new A.Text(spec.GetProperty("text").GetString()!));
    }

    // Keeps the run properties and makes `text` the run's only other child.
    private static void ReplaceContent<TProperties>(OpenXmlElement run, OpenXmlElement text)
        where TProperties : OpenXmlElement
    {
        foreach (var child in run.ChildElements.Where(c => c is not TProperties).ToList())
        {
            child.Remove();
        }
        run.AppendChild(text);
    }

    private static void Error(string phase, Exception error)
    {
        var message = error.Message;
        var trace = error.ToString();
        Emit(new
        {
            @event = "error",
            phase,
            type = error.GetType().FullName,
            message = message[..Math.Min(2000, message.Length)],
            traceback = trace[Math.Max(0, trace.Length - 4000)..],
        });
    }

    private static void Emit(object payload)
    {
        var bytes = JsonSerializer.SerializeToUtf8Bytes(payload, Json);
        Stdout.Write(bytes);
        Stdout.WriteByte((byte)'\n');
        Stdout.Flush();
    }
}
