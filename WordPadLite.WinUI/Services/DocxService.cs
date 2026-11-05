using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace WordPadLite.WinUI;

internal static class DocxService
{
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string PackageRelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";

    public static string ExtractRtf(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        var entry = archive.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("word/document.xml が見つかりません。");

        using var stream = entry.Open();
        var document = XDocument.Load(stream);
        XNamespace w = WordNamespace;
        var body = document.Root?.Element(w + "body")
            ?? throw new InvalidDataException("文書本文が見つかりません。");

        var rtf = new StringBuilder();
        rtf.Append("{\\rtf1\\ansi\\ansicpg1252\\deff0");
        rtf.Append("{\\fonttbl{\\f0\\fnil\\fcharset128 Yu Mincho;}{\\f1\\fnil\\fcharset128 Yu Gothic;}{\\f2\\fnil\\fcharset0 Arial;}}");

        foreach (var element in body.Elements())
        {
            if (element.Name == w + "p")
            {
                AppendParagraph(rtf, element, w);
            }
            else if (element.Name == w + "tbl")
            {
                AppendTable(rtf, element, w);
            }
        }

        rtf.Append('}');
        return rtf.ToString();
    }

    private static void AppendParagraph(StringBuilder rtf, XElement paragraph, XNamespace w)
    {
        var properties = paragraph.Element(w + "pPr");
        var styleId = properties?.Element(w + "pStyle")?.Attribute(w + "val")?.Value ?? string.Empty;
        var isHeading1 = styleId.Contains("Heading1", StringComparison.OrdinalIgnoreCase);
        var isHeading2 = styleId.Contains("Heading2", StringComparison.OrdinalIgnoreCase);
        var isTitle = styleId.Contains("Title", StringComparison.OrdinalIgnoreCase);

        var alignment = properties?.Element(w + "jc")?.Attribute(w + "val")?.Value;
        rtf.Append("\\pard\\plain");
        rtf.Append(alignment switch
        {
            "center" => "\\qc",
            "right" or "end" => "\\qr",
            "both" => "\\qj",
            _ => "\\ql",
        });

        var spacing = properties?.Element(w + "spacing");
        if (spacing is not null)
        {
            if (int.TryParse(spacing.Attribute(w + "before")?.Value, out var before))
            {
                rtf.Append("\\sb").Append(before);
            }
            if (int.TryParse(spacing.Attribute(w + "after")?.Value, out var after))
            {
                rtf.Append("\\sa").Append(after);
            }
        }

        var indentation = properties?.Element(w + "ind");
        if (indentation is not null)
        {
            if (int.TryParse(indentation.Attribute(w + "left")?.Value, out var left))
            {
                rtf.Append("\\li").Append(left);
            }
            if (int.TryParse(indentation.Attribute(w + "right")?.Value, out var right))
            {
                rtf.Append("\\ri").Append(right);
            }
            if (int.TryParse(indentation.Attribute(w + "firstLine")?.Value, out var firstLine))
            {
                rtf.Append("\\fi").Append(firstLine);
            }
        }

        var hasNumbering = properties?.Element(w + "numPr") is not null;
        if (hasNumbering)
        {
            rtf.Append("\\li720\\fi-360 ");
            rtf.Append(RtfText("• "));
        }

        var defaultSize = isTitle ? 56 : isHeading1 ? 36 : isHeading2 ? 30 : 21;
        var defaultBold = isTitle || isHeading1 || isHeading2;

        var runs = paragraph.Descendants(w + "r");
        var wroteRun = false;
        foreach (var run in runs)
        {
            var runProperties = run.Element(w + "rPr");
            var bold = GetBoolean(runProperties?.Element(w + "b"), w) ?? defaultBold;
            var italic = GetBoolean(runProperties?.Element(w + "i"), w) ?? false;
            var underline = HasUnderline(runProperties?.Element(w + "u"), w);
            var strike = GetBoolean(runProperties?.Element(w + "strike"), w) ?? false;
            var size = int.TryParse(runProperties?.Element(w + "sz")?.Attribute(w + "val")?.Value, out var parsedSize)
                ? parsedSize
                : defaultSize;
            var font = FontIndex(runProperties?.Element(w + "rFonts"));
            var vertical = runProperties?.Element(w + "vertAlign")?.Attribute(w + "val")?.Value;

            rtf.Append("\\plain\\f").Append(font).Append("\\fs").Append(size);
            rtf.Append(bold ? "\\b" : "\\b0");
            rtf.Append(italic ? "\\i" : "\\i0");
            rtf.Append(underline ? "\\ul" : "\\ul0");
            rtf.Append(strike ? "\\strike" : "\\strike0");
            if (vertical == "superscript")
            {
                rtf.Append("\\super");
            }
            else if (vertical == "subscript")
            {
                rtf.Append("\\sub");
            }

            foreach (var node in run.Nodes())
            {
                if (node is XElement element)
                {
                    if (element.Name == w + "t")
                    {
                        rtf.Append(RtfText(element.Value));
                        wroteRun = true;
                    }
                    else if (element.Name == w + "tab")
                    {
                        rtf.Append("\\tab ");
                        wroteRun = true;
                    }
                    else if (element.Name == w + "br")
                    {
                        var type = element.Attribute(w + "type")?.Value;
                        rtf.Append(type == "page" ? "\\page " : "\\line ");
                        wroteRun = true;
                    }
                    else if (element.Name == w + "cr")
                    {
                        rtf.Append("\\line ");
                        wroteRun = true;
                    }
                }
            }
        }

        if (!wroteRun && !hasNumbering)
        {
            rtf.Append("\\f0\\fs21 ");
        }

        rtf.Append("\\par\r\n");
    }

    private static void AppendTable(StringBuilder rtf, XElement table, XNamespace w)
    {
        foreach (var row in table.Elements(w + "tr"))
        {
            var cells = row.Elements(w + "tc").ToList();
            for (var index = 0; index < cells.Count; index++)
            {
                var cellText = string.Join(
                    string.Empty,
                    cells[index].Descendants(w + "t").Select(text => text.Value));
                rtf.Append(RtfText(cellText));
                if (index < cells.Count - 1)
                {
                    rtf.Append("\\tab ");
                }
            }
            rtf.Append("\\par\r\n");
        }
    }

    private static bool? GetBoolean(XElement? element, XNamespace w)
    {
        if (element is null)
        {
            return null;
        }

        var value = element.Attribute(w + "val")?.Value;
        return value is null or "1" or "true" or "on";
    }

    private static bool HasUnderline(XElement? element, XNamespace w)
    {
        if (element is null)
        {
            return false;
        }

        var value = element.Attribute(w + "val")?.Value;
        return value is not "none" and not "0" and not "false";
    }

    private static int FontIndex(XElement? fonts)
    {
        var name = fonts?.Attribute(W + "ascii")?.Value
            ?? fonts?.Attribute(W + "hAnsi")?.Value
            ?? fonts?.Attribute(W + "eastAsia")?.Value
            ?? string.Empty;

        if (name.Contains("Gothic", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("ゴシック", StringComparison.Ordinal) ||
            name.Contains("Meiryo", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("メイリオ", StringComparison.Ordinal))
        {
            return 1;
        }

        if (name.Contains("Arial", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("Calibri", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("Segoe", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }

        return 0;
    }

    private static XNamespace W => WordNamespace;

    private static string RtfText(string value)
    {
        var builder = new StringBuilder(value.Length + 12);
        foreach (var character in value)
        {
            switch (character)
            {
                case '\\':
                    builder.Append("\\\\");
                    break;
                case '{':
                    builder.Append("\\{");
                    break;
                case '}':
                    builder.Append("\\}");
                    break;
                case '\r':
                case '\n':
                    builder.Append("\\line ");
                    break;
                case '\t':
                    builder.Append("\\tab ");
                    break;
                default:
                    if (character < 128)
                    {
                        builder.Append(character);
                    }
                    else
                    {
                        var code = (int)character;
                        if (code > 32767)
                        {
                            code -= 65536;
                        }
                        builder.Append("\\u").Append(code).Append('?');
                    }
                    break;
            }
        }
        return builder.ToString();
    }

    public static byte[] CreateDocx(string text, double pageWidth, double pageHeight, double pageMargin)
    {
        XNamespace w = WordNamespace;
        var body = new XElement(w + "body");

        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        if (lines.Length > 1 && lines[^1].Length == 0)
        {
            lines = lines[..^1];
        }

        if (lines.Length == 0)
        {
            lines = new[] { string.Empty };
        }

        foreach (var line in lines)
        {
            body.Add(new XElement(w + "p",
                new XElement(w + "r",
                    new XElement(w + "t",
                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                        line))));
        }

        var widthTwips = Math.Max(1000, (int)Math.Round(pageWidth * 15));
        var heightTwips = Math.Max(1000, (int)Math.Round(pageHeight * 15));
        var marginTwips = Math.Max(0, (int)Math.Round(pageMargin * 15));
        body.Add(new XElement(w + "sectPr",
            new XElement(w + "pgSz",
                new XAttribute(w + "w", widthTwips),
                new XAttribute(w + "h", heightTwips)),
            new XElement(w + "pgMar",
                new XAttribute(w + "top", marginTwips),
                new XAttribute(w + "right", marginTwips),
                new XAttribute(w + "bottom", marginTwips),
                new XAttribute(w + "left", marginTwips),
                new XAttribute(w + "header", Math.Max(0, marginTwips / 2)),
                new XAttribute(w + "footer", Math.Max(0, marginTwips / 2)),
                new XAttribute(w + "gutter", 0))));

        var document = new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(w + "document",
                new XAttribute(XNamespace.Xmlns + "w", WordNamespace),
                body));

        const string contentTypes = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>
            """;

        const string rootRelationships = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """;

        const string documentRelationships = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
            """;

        using var memory = new MemoryStream();
        using (var archive = new ZipArchive(memory, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(archive, "[Content_Types].xml", contentTypes);
            WriteEntry(archive, "_rels/.rels", rootRelationships);
            WriteEntry(archive, "word/_rels/document.xml.rels", documentRelationships);
            WriteEntry(archive, "word/document.xml", document.ToString(SaveOptions.DisableFormatting));
        }

        return memory.ToArray();
    }

    private static void WriteEntry(ZipArchive archive, string name, string content)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(content);
    }
}
