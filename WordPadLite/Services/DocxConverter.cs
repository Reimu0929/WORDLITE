using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace WordPadLite.Services;

public sealed record DocxMargins(double Top, double Right, double Bottom, double Left);

public sealed record DocxPageSettings(
    string PageSize = "A4",
    string Orientation = "portrait",
    DocxMargins? Margins = null,
    string DefaultFont = "Meiryo UI",
    double DefaultFontSize = 11,
    string? HeaderText = null,
    string? FooterText = null,
    bool ShowPageNumbers = false);

public sealed record DocxExportRequest(string Html, string Title = "document", DocxPageSettings? Layout = null);

public sealed record DocxImportResult(string Html);

public static class DocxConverter
{
    private static readonly HtmlParser HtmlParser = new();

    private const double TwipsPerMillimeter = 56.6929133858;
    private const double PixelsToPoints = 0.75;

    /* =========================================================
     * Export: HTML -> DOCX
     * ========================================================= */

    public static byte[] Export(string html, string title, DocxPageSettings? layout)
    {
        layout ??= new DocxPageSettings();
        var document = HtmlParser.ParseDocument(html);
        var body = document.Body ?? throw new ArgumentException("HTML body is missing.", nameof(html));

        using var stream = new MemoryStream();
        using (var wordDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = wordDocument.AddMainDocumentPart();
            mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
            var wordBody = new Body();

            AppendBlocks(body.ChildNodes, wordBody, mainPart, layout);
            wordBody.Append(CreateSectionProperties(layout, mainPart));

            mainPart.Document.Append(wordBody);
            mainPart.Document.Save();
        }

        return stream.ToArray();
    }

    private static void AppendBlocks(INodeList nodes, Body body, MainDocumentPart mainPart, DocxPageSettings layout)
    {
        foreach (var node in nodes)
        {
            if (node is not IElement element) continue;

            switch (element.LocalName)
            {
                case "p":
                case "div":
                    body.Append(CreateParagraph(element, mainPart, layout, GetAlignment(element)));
                    break;
                case "h1":
                case "h2":
                case "h3":
                case "h4":
                case "h5":
                case "h6":
                    body.Append(CreateParagraph(
                        element,
                        mainPart,
                        layout,
                        GetAlignment(element),
                        HeadingSize(element.LocalName),
                        styleId: $"Heading{element.LocalName[1]}"));
                    break;
                case "ul":
                case "ol":
                    AppendList(element, body, mainPart, layout);
                    break;
                case "table":
                    body.Append(CreateTable(element, mainPart, layout));
                    break;
                case "hr":
                    body.Append(CreatePageBreakParagraph());
                    break;
                case "blockquote":
                    body.Append(CreateParagraph(element, mainPart, layout, JustificationValues.Left, 11, quote: true));
                    break;
                default:
                    body.Append(CreateParagraph(element, mainPart, layout, GetAlignment(element)));
                    break;
            }
        }
    }

    private static Paragraph CreateParagraph(
        IElement element,
        MainDocumentPart mainPart,
        DocxPageSettings layout,
        JustificationValues alignment,
        double? fontSize = null,
        bool quote = false,
        string? styleId = null)
    {
        var paragraph = new Paragraph();
        var properties = new ParagraphProperties();
        if (!string.IsNullOrWhiteSpace(styleId))
        {
            properties.Append(new ParagraphStyleId { Val = styleId });
        }

        var (before, after) = GetParagraphSpacing(element);
        properties.Append(new SpacingBetweenLines
        {
            Before = ((int)Math.Round(before * 20)).ToString(CultureInfo.InvariantCulture),
            After = ((int)Math.Round(after * 20)).ToString(CultureInfo.InvariantCulture),
        });
        if (quote)
        {
            properties.Append(new Indentation { Left = "720", Right = "720" });
        }
        properties.Append(new Justification { Val = alignment });
        paragraph.Append(properties);

        AppendInlineNodes(element.ChildNodes, paragraph, mainPart, layout, fontSize);

        if (!paragraph.Elements<Run>().Any() && !paragraph.Elements<Hyperlink>().Any())
        {
            paragraph.Append(CreateRun(string.Empty, layout, fontSize));
        }

        return paragraph;
    }

    private static void AppendInlineNodes(
        INodeList nodes,
        OpenXmlElement container,
        MainDocumentPart mainPart,
        DocxPageSettings layout,
        double? inheritedFontSize = null,
        InlineFormat? inherited = null)
    {
        var format = inherited ?? new InlineFormat();

        foreach (var node in nodes)
        {
            if (node is IText text)
            {
                if (!string.IsNullOrEmpty(text.Data))
                {
                    container.Append(CreateRun(text.Data, layout, inheritedFontSize, format));
                }
                continue;
            }

            if (node is not IElement element) continue;

            switch (element.LocalName)
            {
                case "br":
                    container.Append(new Run(new Break()));
                    break;
                case "img":
                    if (!TryAppendImage(element, container, mainPart))
                    {
                        container.Append(CreateRun($"[画像: {element.GetAttribute("alt") ?? ""}]", layout, inheritedFontSize, format));
                    }
                    break;
                case "a":
                {
                    var href = element.GetAttribute("href");
                    if (!string.IsNullOrWhiteSpace(href) && Uri.TryCreate(href, UriKind.RelativeOrAbsolute, out var uri))
                    {
                        var relationship = mainPart.AddHyperlinkRelationship(uri, true);
                        var hyperlink = new Hyperlink { Id = relationship.Id };
                        var innerFormat = (InlineFormat)format.Clone();
                        innerFormat.Underline = true;
                        innerFormat.Color = innerFormat.Color ?? "0F6CBD";
                        AppendInlineNodes(element.ChildNodes, hyperlink, mainPart, layout, inheritedFontSize, innerFormat);
                        container.Append(hyperlink);
                    }
                    else
                    {
                        AppendInlineNodes(element.ChildNodes, container, mainPart, layout, inheritedFontSize, format);
                    }
                    break;
                }
                case "strong":
                case "b":
                {
                    var next = (InlineFormat)format.Clone();
                    next.Bold = true;
                    AppendInlineNodes(element.ChildNodes, container, mainPart, layout, inheritedFontSize, next);
                    break;
                }
                case "em":
                case "i":
                {
                    var next = (InlineFormat)format.Clone();
                    next.Italic = true;
                    AppendInlineNodes(element.ChildNodes, container, mainPart, layout, inheritedFontSize, next);
                    break;
                }
                case "u":
                {
                    var next = (InlineFormat)format.Clone();
                    next.Underline = true;
                    AppendInlineNodes(element.ChildNodes, container, mainPart, layout, inheritedFontSize, next);
                    break;
                }
                case "s":
                case "strike":
                {
                    var next = (InlineFormat)format.Clone();
                    next.Strike = true;
                    AppendInlineNodes(element.ChildNodes, container, mainPart, layout, inheritedFontSize, next);
                    break;
                }
                case "span":
                case "font":
                {
                    var next = ApplyCss(format, element.GetAttribute("style"));
                    if (element.LocalName == "font")
                    {
                        next.FontFamily = element.GetAttribute("face") ?? next.FontFamily;
                        next.Color = NormalizeColor(element.GetAttribute("color")) ?? next.Color;
                        if (double.TryParse(element.GetAttribute("size"), NumberStyles.Float, CultureInfo.InvariantCulture, out var legacySize))
                        {
                            next.FontSize = 7 + legacySize;
                        }
                    }
                    AppendInlineNodes(element.ChildNodes, container, mainPart, layout, inheritedFontSize, next);
                    break;
                }
                case "code":
                {
                    var next = (InlineFormat)format.Clone();
                    next.FontFamily = "Consolas";
                    AppendInlineNodes(element.ChildNodes, container, mainPart, layout, inheritedFontSize, next);
                    break;
                }
                default:
                    AppendInlineNodes(element.ChildNodes, container, mainPart, layout, inheritedFontSize, format);
                    break;
            }
        }
    }

    private static Run CreateRun(
        string text,
        DocxPageSettings layout,
        double? fontSize = null,
        InlineFormat? format = null)
    {
        format ??= new InlineFormat();
        var run = new Run();
        var properties = new RunProperties();

        var family = format.FontFamily ?? layout.DefaultFont;
        if (!string.IsNullOrWhiteSpace(family))
        {
            properties.Append(new RunFonts { Ascii = family, HighAnsi = family, EastAsia = family });
        }

        if (format.Bold) properties.Append(new Bold());
        if (format.Italic) properties.Append(new Italic());
        if (format.Strike) properties.Append(new Strike());

        var color = NormalizeColor(format.Color);
        if (color is not null) properties.Append(new Color { Val = color });

        var size = format.FontSize ?? fontSize ?? layout.DefaultFontSize;
        properties.Append(new FontSize { Val = ToHalfPoints(size) });
        properties.Append(new FontSizeComplexScript { Val = ToHalfPoints(size) });

        if (format.Underline) properties.Append(new Underline { Val = UnderlineValues.Single });
        if (!string.IsNullOrWhiteSpace(format.Background))
        {
            properties.Append(new Shading
            {
                Val = ShadingPatternValues.Clear,
                Fill = NormalizeColor(format.Background),
            });
        }

        if (properties.HasChildren) run.Append(properties);
        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    private static void AppendList(IElement list, Body body, MainDocumentPart mainPart, DocxPageSettings layout)
    {
        var ordered = string.Equals(list.LocalName, "ol", StringComparison.OrdinalIgnoreCase);
        var index = 1;
        foreach (var item in list.Children.Where(c => c.LocalName == "li"))
        {
            var marker = ordered ? $"{index}. " : "• ";
            index++;

            var paragraph = new Paragraph();
            var properties = new ParagraphProperties();
            properties.Append(new SpacingBetweenLines { After = "60" });
            properties.Append(new Indentation { Left = "720", Hanging = "360" });
            paragraph.Append(properties);
            paragraph.Append(CreateRun(marker, layout));
            AppendInlineNodes(item.ChildNodes, paragraph, mainPart, layout);
            body.Append(paragraph);
        }
    }

    private static Table CreateTable(IElement tableElement, MainDocumentPart mainPart, DocxPageSettings layout)
    {
        var table = new Table();
        table.Append(new TableProperties(
            new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4, Color = "BFBFBF" },
                new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "BFBFBF" },
                new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "BFBFBF" },
                new RightBorder { Val = BorderValues.Single, Size = 4, Color = "BFBFBF" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "BFBFBF" },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "BFBFBF" })));

        var columnCount = Math.Max(1, tableElement.QuerySelectorAll("tr")
            .Select(row => row.Children.Count(c => c.LocalName is "td" or "th"))
            .DefaultIfEmpty(1)
            .Max());
        table.Append(new TableGrid(Enumerable.Range(0, columnCount).Select(_ => new GridColumn()).ToArray()));

        foreach (var rowElement in tableElement.QuerySelectorAll("tr"))
        {
            var row = new TableRow();
            foreach (var cellElement in rowElement.Children.Where(c => c.LocalName is "td" or "th"))
            {
                var cell = new TableCell();
                var paragraphs = cellElement.Children.Where(c => c.LocalName is "p" or "div").ToArray();
                if (paragraphs.Length == 0)
                {
                    var paragraph = new Paragraph();
                    AppendInlineNodes(cellElement.ChildNodes, paragraph, mainPart, layout);
                    cell.Append(paragraph);
                }
                else
                {
                    foreach (var paragraphElement in paragraphs)
                    {
                        cell.Append(CreateParagraph(paragraphElement, mainPart, layout, GetAlignment(paragraphElement)));
                    }
                }
                row.Append(cell);
            }
            for (var i = row.ChildElements.Count; i < columnCount; i++)
            {
                row.Append(new TableCell(new Paragraph()));
            }
            table.Append(row);
        }

        return table;
    }

    private static int _imageId;

    private static bool TryAppendImage(IElement element, OpenXmlElement container, MainDocumentPart mainPart)
    {
        var src = element.GetAttribute("src");
        if (string.IsNullOrWhiteSpace(src) || !src.StartsWith("data:image/", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var comma = src.IndexOf(',');
        if (comma < 0) return false;

        var metadata = src[5..comma];
        if (!metadata.Contains(";base64", StringComparison.OrdinalIgnoreCase)) return false;

        var mime = metadata.Split(';')[0].ToLowerInvariant();
        var partType = mime switch
        {
            "image/png" => ImagePartType.Png,
            "image/jpeg" or "image/jpg" => ImagePartType.Jpeg,
            "image/gif" => ImagePartType.Gif,
            "image/bmp" => ImagePartType.Bmp,
            "image/tiff" => ImagePartType.Tiff,
            _ => (PartTypeInfo?)null,
        };
        if (partType is null) return false;

        byte[] bytes;
        try
        {
            bytes = Convert.FromBase64String(src[(comma + 1)..]);
        }
        catch (FormatException)
        {
            return false;
        }

        using var imageStream = new MemoryStream(bytes);
        var imagePart = mainPart.AddImagePart(partType.Value);
        imagePart.FeedData(imageStream);

        TryGetImageDimensions(bytes, out var naturalWidth, out var naturalHeight);
        var (width, height) = GetImageSize(element, naturalWidth, naturalHeight);
        var relationshipId = mainPart.GetIdOfPart(imagePart);
        container.Append(new Run(CreateImageDrawing(relationshipId, width, height)));
        return true;
    }

    private static (long Width, long Height) GetImageSize(IElement element, int naturalWidth, int naturalHeight)
    {
        var widthPx = ParsePixels(element.GetAttribute("width")) ?? ParseStylePixels(element, "width");
        var heightPx = ParsePixels(element.GetAttribute("height")) ?? ParseStylePixels(element, "height");

        if (widthPx is null)
        {
            widthPx = naturalWidth > 0 ? Math.Min(naturalWidth, 480) : 480;
        }

        if (heightPx is null || heightPx <= 0)
        {
            heightPx = naturalWidth > 0 && naturalHeight > 0
                ? widthPx.Value * naturalHeight / naturalWidth
                : widthPx.Value * 0.75;
        }

        const double emuPerPixel = 9525.0;
        return (
            Math.Max(1, (long)Math.Round(widthPx.Value * emuPerPixel)),
            Math.Max(1, (long)Math.Round(heightPx.Value * emuPerPixel)));
    }

    private static (double Before, double After) GetParagraphSpacing(IElement element)
    {
        var before = ParseStyleLength(element, "margin-top") ?? 0;
        var after = ParseStyleLength(element, "margin-bottom") ?? 6;
        return (before, after);
    }

    private static double? ParseStyleLength(IElement element, string property)
    {
        var style = element.GetAttribute("style");
        if (string.IsNullOrWhiteSpace(style)) return null;

        var match = Regex.Match(
            style,
            $@"(?:^|;)\s*{Regex.Escape(property)}\s*:\s*([\d.]+)\s*(pt|px)?",
            RegexOptions.IgnoreCase);
        if (!match.Success ||
            !double.TryParse(match.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
        {
            return null;
        }

        return string.Equals(match.Groups[2].Value, "px", StringComparison.OrdinalIgnoreCase)
            ? number * 0.75
            : number;
    }

    private static double? ParsePixels(string? value) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && number > 0
            ? number
            : null;

    private static double? ParseStylePixels(IElement element, string property)
    {
        var style = element.GetAttribute("style");
        if (string.IsNullOrWhiteSpace(style)) return null;

        var match = Regex.Match(
            style,
            $@"(?:^|;)\s*{Regex.Escape(property)}\s*:\s*([\d.]+)\s*px",
            RegexOptions.IgnoreCase);
        return match.Success &&
               double.TryParse(match.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) &&
               number > 0
            ? number
            : null;
    }

    private static bool TryGetImageDimensions(byte[] bytes, out int width, out int height)
    {
        width = 0;
        height = 0;

        if (bytes.Length >= 24 &&
            bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
        {
            width = ReadBigEndianInt32(bytes, 16);
            height = ReadBigEndianInt32(bytes, 20);
            return width > 0 && height > 0;
        }

        if (bytes.Length >= 10 && bytes[0] == 'G' && bytes[1] == 'I' && bytes[2] == 'F')
        {
            width = bytes[6] | (bytes[7] << 8);
            height = bytes[8] | (bytes[9] << 8);
            return width > 0 && height > 0;
        }

        if (bytes.Length >= 26 && bytes[0] == 'B' && bytes[1] == 'M')
        {
            width = ReadLittleEndianInt32(bytes, 18);
            height = Math.Abs(ReadLittleEndianInt32(bytes, 22));
            return width > 0 && height > 0;
        }

        if (bytes.Length >= 4 && bytes[0] == 0xFF && bytes[1] == 0xD8)
        {
            var index = 2;
            while (index + 9 < bytes.Length)
            {
                if (bytes[index] != 0xFF)
                {
                    index++;
                    continue;
                }

                var marker = bytes[index + 1];
                index += 2;
                if (marker is 0xD8 or 0xD9) continue;
                if (index + 2 > bytes.Length) break;

                var length = (bytes[index] << 8) | bytes[index + 1];
                if (length < 2) break;

                if (marker is >= 0xC0 and <= 0xCF && marker is not 0xC4 and not 0xC8 and not 0xCC)
                {
                    height = (bytes[index + 3] << 8) | bytes[index + 4];
                    width = (bytes[index + 5] << 8) | bytes[index + 6];
                    return width > 0 && height > 0;
                }

                index += length;
            }
        }

        return false;
    }

    private static int ReadBigEndianInt32(byte[] bytes, int offset) =>
        (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];

    private static int ReadLittleEndianInt32(byte[] bytes, int offset) =>
        bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24);

    private static Drawing CreateImageDrawing(string relationshipId, long width, long height)
    {
        var id = (uint)Interlocked.Increment(ref _imageId);
        return new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = width, Cy = height },
                new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.DocProperties { Id = id, Name = $"Picture {id}" },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0U, Name = $"Picture {id}" },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relationshipId },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0L, Y = 0L },
                                    new A.Extents { Cx = width, Cy = height }),
                                new A.PresetGeometry(new A.AdjustValueList())
                                {
                                    Preset = A.ShapeTypeValues.Rectangle,
                                })))
                    {
                        Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture",
                    }))
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
            });
    }

    private static Paragraph CreatePageBreakParagraph()
    {
        var paragraph = new Paragraph();
        paragraph.Append(new Run(new Break { Type = BreakValues.Page }));
        return paragraph;
    }

    private static SectionProperties CreateSectionProperties(DocxPageSettings layout, MainDocumentPart mainPart)
    {
        var (widthMm, heightMm) = layout.PageSize switch
        {
            "B5" => (182.0, 257.0),
            "Letter" => (215.9, 279.4),
            "Legal" => (215.9, 355.6),
            _ => (210.0, 297.0),
        };

        if (string.Equals(layout.Orientation, "landscape", StringComparison.OrdinalIgnoreCase))
        {
            (widthMm, heightMm) = (heightMm, widthMm);
        }

        var margins = layout.Margins ?? new DocxMargins(25, 25, 25, 25);
        var section = new SectionProperties();

        var headerId = CreateHeader(mainPart, layout);
        if (headerId is not null)
        {
            section.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = headerId });
        }

        var footerId = CreateFooter(mainPart, layout);
        if (footerId is not null)
        {
            section.Append(new FooterReference { Type = HeaderFooterValues.Default, Id = footerId });
        }

        section.Append(
            new PageSize
            {
                Width = MmToTwips(widthMm),
                Height = MmToTwips(heightMm),
                Orient = string.Equals(layout.Orientation, "landscape", StringComparison.OrdinalIgnoreCase)
                    ? PageOrientationValues.Landscape
                    : PageOrientationValues.Portrait,
            },
            new PageMargin
            {
                Top = (int)MmToTwips(margins.Top),
                Right = MmToTwips(margins.Right),
                Bottom = (int)MmToTwips(margins.Bottom),
                Left = MmToTwips(margins.Left),
                Header = 720U,
                Footer = 720U,
                Gutter = 0U,
            });

        return section;
    }

    private static string? CreateHeader(MainDocumentPart mainPart, DocxPageSettings layout)
    {
        if (string.IsNullOrWhiteSpace(layout.HeaderText)) return null;

        var headerPart = mainPart.AddNewPart<HeaderPart>();
        var paragraph = new Paragraph();
        paragraph.Append(CreateSmallTextRun(layout.HeaderText, layout));
        headerPart.Header = new Header(paragraph);
        headerPart.Header.Save();
        return mainPart.GetIdOfPart(headerPart);
    }

    private static string? CreateFooter(MainDocumentPart mainPart, DocxPageSettings layout)
    {
        var hasText = !string.IsNullOrWhiteSpace(layout.FooterText);
        if (!hasText && !layout.ShowPageNumbers) return null;

        var footerPart = mainPart.AddNewPart<FooterPart>();
        var paragraph = new Paragraph();
        if (hasText)
        {
            paragraph.Append(CreateSmallTextRun(layout.FooterText!, layout));
            if (layout.ShowPageNumbers) paragraph.Append(CreateSmallTextRun("    ", layout));
        }
        if (layout.ShowPageNumbers)
        {
            paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
            paragraph.Append(new Run(new FieldCode(" PAGE ") { Space = SpaceProcessingModeValues.Preserve }));
            paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
            paragraph.Append(new Run(new Text("1")));
            paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }

        footerPart.Footer = new Footer(paragraph);
        footerPart.Footer.Save();
        return mainPart.GetIdOfPart(footerPart);
    }

    private static Run CreateSmallTextRun(string text, DocxPageSettings layout)
    {
        var run = new Run();
        var properties = new RunProperties();
        var family = layout.DefaultFont ?? "Meiryo UI";
        properties.Append(new RunFonts { Ascii = family, HighAnsi = family, EastAsia = family });
        properties.Append(new Color { Val = "666666" });
        properties.Append(new FontSize { Val = "18" });
        run.Append(properties);
        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    private static uint MmToTwips(double mm) => (uint)Math.Round(mm * TwipsPerMillimeter);

    private static double HeadingSize(string tag) => tag switch
    {
        "h1" => 24,
        "h2" => 18,
        "h3" => 14,
        "h4" => 12,
        "h5" => 11,
        _ => 10.5,
    };

    private static JustificationValues GetAlignment(IElement element)
    {
        var value = element.GetAttribute("style") ?? string.Empty;
        var match = Regex.Match(value, @"text-align\s*:\s*(left|center|right|justify)", RegexOptions.IgnoreCase);
        return match.Groups[1].Value.ToLowerInvariant() switch
        {
            "center" => JustificationValues.Center,
            "right" => JustificationValues.Right,
            "justify" => JustificationValues.Both,
            _ => JustificationValues.Left,
        };
    }

    private static InlineFormat ApplyCss(InlineFormat current, string? style)
    {
        var next = (InlineFormat)current.Clone();
        if (string.IsNullOrWhiteSpace(style)) return next;

        foreach (var declaration in style.Split(';', StringSplitOptions.RemoveEmptyEntries))
        {
            var separator = declaration.IndexOf(':');
            if (separator <= 0) continue;
            var name = declaration[..separator].Trim().ToLowerInvariant();
            var value = declaration[(separator + 1)..].Trim();

            switch (name)
            {
                case "font-weight":
                    next.Bold = value is "bold" or "600" or "700" or "800" or "900";
                    break;
                case "font-style":
                    next.Italic = value == "italic";
                    break;
                case "text-decoration":
                case "text-decoration-line":
                    next.Underline = value.Contains("underline", StringComparison.OrdinalIgnoreCase);
                    next.Strike = value.Contains("line-through", StringComparison.OrdinalIgnoreCase);
                    break;
                case "color":
                    next.Color = NormalizeColor(value) ?? next.Color;
                    break;
                case "background-color":
                case "background":
                    next.Background = NormalizeColor(value) ?? next.Background;
                    break;
                case "font-family":
                    next.FontFamily = value.Split(',')[0].Trim().Trim('\'', '"');
                    break;
                case "font-size":
                    if (TryParseCssFontSize(value, out var size)) next.FontSize = size;
                    break;
            }
        }

        return next;
    }

    private static bool TryParseCssFontSize(string value, out double points)
    {
        points = 0;
        var match = Regex.Match(value, @"([\d.]+)\s*(pt|px|em|rem)?", RegexOptions.IgnoreCase);
        if (!match.Success) return false;
        if (!double.TryParse(match.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)) return false;

        points = match.Groups[2].Value.ToLowerInvariant() switch
        {
            "px" => number * PixelsToPoints,
            "em" or "rem" => 11 * number,
            _ => number,
        };
        return true;
    }

    private static string ToHalfPoints(double points) =>
        Math.Max(2, (int)Math.Round(points * 2)).ToString(CultureInfo.InvariantCulture);

    private static string? NormalizeColor(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        value = value.Trim();

        if (value.StartsWith('#'))
        {
            var hex = value[1..];
            return hex.Length switch
            {
                3 => $"{hex[0]}{hex[0]}{hex[1]}{hex[1]}{hex[2]}{hex[2]}".ToUpperInvariant(),
                4 => $"{hex[0]}{hex[0]}{hex[1]}{hex[1]}{hex[2]}{hex[2]}".ToUpperInvariant(),
                6 => hex.ToUpperInvariant(),
                8 => hex[..6].ToUpperInvariant(),
                _ => null,
            };
        }

        var rgb = Regex.Match(value, @"rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)", RegexOptions.IgnoreCase);
        if (rgb.Success)
        {
            var r = Math.Clamp(int.Parse(rgb.Groups[1].Value, CultureInfo.InvariantCulture), 0, 255);
            var g = Math.Clamp(int.Parse(rgb.Groups[2].Value, CultureInfo.InvariantCulture), 0, 255);
            var b = Math.Clamp(int.Parse(rgb.Groups[3].Value, CultureInfo.InvariantCulture), 0, 255);
            return $"{r:X2}{g:X2}{b:X2}";
        }

        return value.ToUpperInvariant() switch
        {
            "BLACK" => "000000",
            "WHITE" => "FFFFFF",
            "RED" => "FF0000",
            "GREEN" => "008000",
            "BLUE" => "0000FF",
            _ => null,
        };
    }

    private sealed class InlineFormat
    {
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strike { get; set; }
        public string? Color { get; set; }
        public string? Background { get; set; }
        public string? FontFamily { get; set; }
        public double? FontSize { get; set; }

        public object Clone() => MemberwiseClone();
    }

    /* =========================================================
     * Import: DOCX -> HTML
     * ========================================================= */

    public static DocxImportResult Import(Stream stream)
    {
        using var wordDocument = WordprocessingDocument.Open(stream, false);
        var mainPart = wordDocument.MainDocumentPart;
        var body = mainPart?.Document?.Body;
        if (body is null) return new DocxImportResult("<p><br></p>");

        var html = new StringBuilder();
        List<string>? listItems = null;

        void FlushList()
        {
            if (listItems is { Count: > 0 })
            {
                html.Append("<ul>");
                foreach (var item in listItems) html.Append(item);
                html.Append("</ul>");
            }
            listItems = null;
        }

        foreach (var element in body.ChildElements)
        {
            if (element is Paragraph paragraph)
            {
                var numbering = paragraph.ParagraphProperties?.NumberingProperties;
                if (numbering is not null)
                {
                    listItems ??= new List<string>();
                    listItems.Add($"<li>{ParagraphInlineHtml(paragraph, mainPart)}</li>");
                    continue;
                }

                FlushList();
                html.Append(ParagraphToHtml(paragraph, mainPart));
            }
            else if (element is Table table)
            {
                FlushList();
                html.Append(TableToHtml(table, mainPart));
            }
        }

        FlushList();

        return new DocxImportResult(html.Length == 0 ? "<p><br></p>" : html.ToString());
    }

    private static string ParagraphToHtml(Paragraph paragraph, MainDocumentPart? mainPart)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? string.Empty;
        var alignment = paragraph.ParagraphProperties?.Justification?.Val?.Value;
        var style = string.Empty;
        if (alignment == JustificationValues.Center) style = " style=\"text-align: center\"";
        else if (alignment == JustificationValues.Right) style = " style=\"text-align: right\"";
        else if (alignment == JustificationValues.Both || alignment == JustificationValues.Distribute)
        {
            style = " style=\"text-align: justify\"";
        }

        var tag = "p";
        if (styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
        {
            var digit = styleId.FirstOrDefault(char.IsDigit);
            tag = digit is >= '1' and <= '6' ? $"h{digit}" : "h2";
        }
        else if (styleId.Equals("Title", StringComparison.OrdinalIgnoreCase))
        {
            tag = "h1";
        }

        return $"<{tag}{style}>{ParagraphInlineHtml(paragraph, mainPart)}</{tag}>";
    }

    private static string ParagraphInlineHtml(Paragraph paragraph, MainDocumentPart? mainPart)
    {
        var html = new StringBuilder();
        foreach (var element in paragraph.ChildElements)
        {
            switch (element)
            {
                case Run run:
                    html.Append(RunToHtml(run));
                    break;
                case Hyperlink hyperlink:
                {
                    var href = hyperlink.Id is not null
                        ? mainPart?.HyperlinkRelationships.FirstOrDefault(r => r.Id == hyperlink.Id.Value)?.Uri.ToString()
                        : hyperlink.Anchor?.Value;
                    var inner = new StringBuilder();
                    foreach (var run in hyperlink.Elements<Run>()) inner.Append(RunToHtml(run));
                    if (string.IsNullOrWhiteSpace(href)) html.Append(inner);
                    else html.Append($"<a href=\"{EscapeAttribute(href)}\">{inner}</a>");
                    break;
                }
                case SimpleField simpleField:
                    foreach (var run in simpleField.Elements<Run>()) html.Append(RunToHtml(run));
                    break;
                case InsertedRun insertedRun:
                    foreach (var run in insertedRun.Elements<Run>()) html.Append(RunToHtml(run));
                    break;
            }
        }
        return html.ToString();
    }

    private static string RunToHtml(Run run)
    {
        var html = new StringBuilder();
        var properties = run.RunProperties;

        foreach (var child in run.ChildElements)
        {
            switch (child)
            {
                case Text text:
                    html.Append(EscapeText(text.Text));
                    break;
                case Break:
                    html.Append("<br>");
                    break;
                case TabChar:
                    html.Append("&emsp;");
                    break;
                case CarriageReturn:
                    html.Append("<br>");
                    break;
                case NoBreakHyphen:
                    html.Append('-');
                    break;
                case SoftHyphen:
                    html.Append("&shy;");
                    break;
            }
        }

        if (html.Length == 0) return string.Empty;

        if (properties?.Bold is not null) html.Insert(0, "<strong>").Append("</strong>");
        if (properties?.Italic is not null) html.Insert(0, "<em>").Append("</em>");
        if (properties?.Strike is not null) html.Insert(0, "<s>").Append("</s>");
        if (properties?.Underline is not null && properties.Underline.Val?.Value != UnderlineValues.None)
        {
            html.Insert(0, "<u>").Append("</u>");
        }

        var styles = new List<string>();
        var color = properties?.Color?.Val?.Value;
        if (!string.IsNullOrWhiteSpace(color)) styles.Add($"color: #{color}");

        var size = properties?.FontSize?.Val?.Value;
        if (double.TryParse(size, NumberStyles.Float, CultureInfo.InvariantCulture, out var halfPoints))
        {
            styles.Add($"font-size: {halfPoints / 2:0.##}pt");
        }

        var font = properties?.RunFonts?.EastAsia?.Value ?? properties?.RunFonts?.Ascii?.Value;
        if (!string.IsNullOrWhiteSpace(font)) styles.Add($"font-family: {font}");

        var shading = properties?.Shading?.Fill?.Value;
        if (!string.IsNullOrWhiteSpace(shading)) styles.Add($"background-color: #{shading}");

        if (styles.Count > 0)
        {
            html.Insert(0, $"<span style=\"{string.Join("; ", styles)}\">").Append("</span>");
        }

        return html.ToString();
    }

    private static string TableToHtml(Table table, MainDocumentPart? mainPart)
    {
        var html = new StringBuilder("<table><tbody>");
        foreach (var row in table.Elements<TableRow>())
        {
            html.Append("<tr>");
            foreach (var cell in row.Elements<TableCell>())
            {
                html.Append("<td>");
                foreach (var paragraph in cell.Elements<Paragraph>())
                {
                    html.Append(ParagraphToHtml(paragraph, mainPart));
                }
                html.Append("</td>");
            }
            html.Append("</tr>");
        }
        html.Append("</tbody></table>");
        return html.ToString();
    }

    private static string EscapeText(string? value) =>
        (value ?? string.Empty)
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;");

    private static string EscapeAttribute(string? value) =>
        EscapeText(value).Replace("'", "&#39;");
}
