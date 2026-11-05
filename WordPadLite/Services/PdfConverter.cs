using System.Globalization;
using System.Text.RegularExpressions;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using IElement = AngleSharp.Dom.IElement;

namespace WordPadLite.Services;

public sealed record PdfExportRequest(string Html, string Title = "document", DocxPageSettings? Layout = null);

public static class PdfConverter
{
    private static readonly HtmlParser HtmlParser = new();
    private static bool _fontRegistered;

    static PdfConverter()
    {
        QuestPDF.Settings.License = LicenseType.Community;
    }

    public static byte[] Export(string html, string title, DocxPageSettings? layout, string? fontDirectory = null)
    {
        EnsureFonts(fontDirectory);
        layout ??= new DocxPageSettings();
        var document = HtmlParser.ParseDocument(html);
        var body = document.Body ?? throw new ArgumentException("HTML body is missing.", nameof(html));

        return QuestPDF.Fluent.Document.Create(container =>
        {
            container.Page(page =>
            {
                ConfigurePage(page, layout);
                page.Content().Column(column =>
                {
                    column.Spacing(4);
                    AddBlocks(column, body.ChildNodes, layout);
                });
            });
        }).GeneratePdf();
    }

    private static void ConfigurePage(PageDescriptor page, DocxPageSettings layout)
    {
        var size = layout.PageSize switch
        {
            "B5" => PageSizes.B5,
            "Letter" => PageSizes.Letter,
            "Legal" => PageSizes.Legal,
            _ => PageSizes.A4,
        };

        if (string.Equals(layout.Orientation, "landscape", StringComparison.OrdinalIgnoreCase))
        {
            size = size.Landscape();
        }

        page.Size(size);
        var margins = layout.Margins ?? new DocxMargins(25, 25, 25, 25);
        page.MarginTop((float)margins.Top, Unit.Millimetre);
        page.MarginRight((float)margins.Right, Unit.Millimetre);
        page.MarginBottom((float)margins.Bottom, Unit.Millimetre);
        page.MarginLeft((float)margins.Left, Unit.Millimetre);

        var font = string.IsNullOrWhiteSpace(layout.DefaultFont) ? "Noto Sans JP" : layout.DefaultFont;
        page.DefaultTextStyle(style => style
            .FontFamily(new[] { font, "Noto Sans JP", "Lato" })
            .FontSize((float)layout.DefaultFontSize)
            .LineHeight(1.4f));

        if (!string.IsNullOrWhiteSpace(layout.HeaderText))
        {
            page.Header().Text(layout.HeaderText).FontSize(9).FontColor(Colors.Grey.Medium);
        }

        if (!string.IsNullOrWhiteSpace(layout.FooterText) || layout.ShowPageNumbers)
        {
            page.Footer().Row(row =>
            {
                if (!string.IsNullOrWhiteSpace(layout.FooterText))
                {
                    row.RelativeItem().Text(layout.FooterText).FontSize(9).FontColor(Colors.Grey.Medium);
                }
                else
                {
                    row.RelativeItem();
                }

                if (layout.ShowPageNumbers)
                {
                    row.ConstantItem(100).AlignRight().Text(text =>
                    {
                        text.DefaultTextStyle(style => style.FontSize(9).FontColor(Colors.Grey.Medium));
                        text.CurrentPageNumber();
                        text.Span(" / ");
                        text.TotalPages();
                    });
                }
            });
        }
    }

    private static void ApplyTextAlignment(TextDescriptor text, IElement element)
    {
        var style = element.GetAttribute("style") ?? string.Empty;
        var match = Regex.Match(style, @"text-align\s*:\s*(left|center|right|justify)", RegexOptions.IgnoreCase);
        switch (match.Groups[1].Value.ToLowerInvariant())
        {
            case "center":
                text.AlignCenter();
                break;
            case "right":
                text.AlignRight();
                break;
            case "justify":
                text.Justify();
                break;
        }
    }

    private static void AddImageBlock(ColumnDescriptor column, byte[] bytes, IElement image)
    {
        var widthPx = ParseStylePixels(image, "width") ?? 460;
        var widthPt = Math.Min(widthPx * 0.75, 460);
        column.Item()
            .PaddingVertical(4)
            .AlignCenter()
            .MaxWidth((float)widthPt, Unit.Point)
            .Image(bytes);
    }

    private static bool TryAppendInlineImage(TextDescriptor text, IElement element)
    {
        if (!TryGetDataImageBytes(element.GetAttribute("src"), out var bytes)) return false;

        var widthPx = ParseStylePixels(element, "width");
        var widthPt = widthPx.HasValue ? Math.Min(widthPx.Value * 0.75, 320) : 240;
        text.Element(
            container => container.Width((float)widthPt, Unit.Point).Image(bytes),
            TextInjectedElementAlignment.AboveBaseline);
        return true;
    }

    private static bool TryGetDataImageBytes(string? src, out byte[] bytes)
    {
        bytes = Array.Empty<byte>();
        if (string.IsNullOrWhiteSpace(src) || !src.StartsWith("data:image/", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var comma = src.IndexOf(',');
        if (comma < 0) return false;

        var metadata = src[5..comma];
        if (!metadata.Contains(";base64", StringComparison.OrdinalIgnoreCase)) return false;

        try
        {
            bytes = Convert.FromBase64String(src[(comma + 1)..]);
            return bytes.Length > 0;
        }
        catch (FormatException)
        {
            return false;
        }
    }

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

    private static double? ParseStylePoints(IElement element, string property)
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

    private static void AddBlocks(ColumnDescriptor column, INodeList nodes, DocxPageSettings layout)
    {
        var inherited = new PdfInlineFormat { FontSize = (float)layout.DefaultFontSize, FontFamily = layout.DefaultFont };

        foreach (var node in nodes)
        {
            if (node is not IElement element) continue;

            switch (element.LocalName)
            {
                case "p":
                case "div":
                {
                    var image = element.QuerySelector("img");
                    var textContent = (element.TextContent ?? string.Empty).Replace("\u00a0", string.Empty).Trim();
                    if (image is not null && textContent.Length == 0 &&
                        TryGetDataImageBytes(image.GetAttribute("src"), out var imageBytes))
                    {
                        AddImageBlock(column, imageBytes, image);
                        break;
                    }

                    var before = ParseStylePoints(element, "margin-top") ?? 0;
                    var after = ParseStylePoints(element, "margin-bottom") ?? 4;
                    column.Item()
                        .PaddingTop((float)before)
                        .PaddingBottom((float)after)
                        .Text(text =>
                        {
                            ApplyTextAlignment(text, element);
                            AppendInline(text, element, inherited);
                        });
                    break;
                }
                case "h1":
                case "h2":
                case "h3":
                case "h4":
                case "h5":
                case "h6":
                {
                    var size = HeadingSize(element.LocalName);
                    var format = (PdfInlineFormat)inherited.Clone();
                    format.Bold = true;
                    format.FontSize = size;
                    var before = 8 + (ParseStylePoints(element, "margin-top") ?? 0);
                    var after = 2 + (ParseStylePoints(element, "margin-bottom") ?? 0);
                    column.Item()
                        .PaddingTop((float)before)
                        .PaddingBottom((float)after)
                        .Text(text =>
                        {
                            ApplyTextAlignment(text, element);
                            AppendInline(text, element, format);
                        });
                    break;
                }
                case "ul":
                case "ol":
                    AddList(column, element, inherited);
                    break;
                case "table":
                    AddTable(column, element, inherited);
                    break;
                case "hr":
                    column.Item().PageBreak();
                    break;
                case "blockquote":
                    column.Item()
                        .PaddingLeft(14)
                        .BorderLeft(3)
                        .BorderColor(Colors.BlueGrey.Lighten1)
                        .PaddingLeft(10)
                        .PaddingVertical(4)
                        .Text(text => AppendInline(text, element, inherited));
                    break;
                default:
                    column.Item().PaddingBottom(4).Text(text => AppendInline(text, element, inherited));
                    break;
            }
        }
    }

    private static void AddList(ColumnDescriptor column, IElement list, PdfInlineFormat inherited)
    {
        var ordered = string.Equals(list.LocalName, "ol", StringComparison.OrdinalIgnoreCase);
        var index = 1;

        column.Item().PaddingLeft(16).Column(listColumn =>
        {
            listColumn.Spacing(2);
            foreach (var item in list.Children.Where(c => c.LocalName == "li"))
            {
                var marker = ordered ? $"{index++}." : "•";
                listColumn.Item().Row(row =>
                {
                    row.ConstantItem(18).Text(text =>
                    {
                        text.DefaultTextStyle(style => style
                            .FontFamily(new[] { inherited.FontFamily ?? "Noto Sans JP", "Noto Sans JP" })
                            .FontSize(inherited.FontSize ?? 11));
                        text.Span(marker);
                    });
                    row.RelativeItem().Text(text =>
                    {
                        ApplyTextAlignment(text, item);
                        AppendInline(text, item, inherited);
                    });
                });
            }
        });
    }

    private static void AddTable(ColumnDescriptor column, IElement tableElement, PdfInlineFormat inherited)
    {
        var rows = tableElement.QuerySelectorAll("tr").ToArray();
        if (rows.Length == 0) return;

        var columnCount = rows.Max(row => row.Children.Count(c => c.LocalName is "td" or "th"));
        columnCount = Math.Clamp(columnCount, 1, 20);

        column.Item().PaddingVertical(4).Table(table =>
        {
            table.ColumnsDefinition(columns =>
            {
                for (var i = 0; i < columnCount; i++) columns.RelativeColumn();
            });

            foreach (var rowElement in rows)
            {
                var cells = rowElement.Children.Where(c => c.LocalName is "td" or "th").ToArray();
                for (var i = 0; i < columnCount; i++)
                {
                    var cell = i < cells.Length ? cells[i] : null;
                    table.Cell()
                        .Border(0.5f)
                        .BorderColor(Colors.Grey.Lighten1)
                        .Padding(4)
                        .Text(text =>
                        {
                            if (cell is not null)
                            {
                                ApplyTextAlignment(text, cell);
                                AppendInline(text, cell, inherited);
                            }
                        });
                }
            }
        });
    }

    private static void AppendInline(TextDescriptor text, IElement element, PdfInlineFormat inherited) =>
        AppendInline(text, element.ChildNodes, inherited);

    private static void AppendInline(TextDescriptor text, INodeList nodes, PdfInlineFormat inherited)
    {
        var format = inherited;

        foreach (var node in nodes)
        {
            if (node is IText textNode)
            {
                if (!string.IsNullOrEmpty(textNode.Data))
                {
                    AppendSpan(text, textNode.Data, format);
                }
                continue;
            }

            if (node is not IElement element) continue;

            switch (element.LocalName)
            {
                case "br":
                    text.Span("\n");
                    break;
                case "img":
                    if (!TryAppendInlineImage(text, element))
                    {
                        AppendSpan(text, $"[画像: {element.GetAttribute("alt") ?? ""}]", format);
                    }
                    break;
                case "a":
                {
                    var next = (PdfInlineFormat)format.Clone();
                    next.Underline = true;
                    next.Color = next.Color ?? "#0F6CBD";
                    AppendInline(text, element.ChildNodes, next);
                    break;
                }
                case "strong":
                case "b":
                {
                    var next = (PdfInlineFormat)format.Clone();
                    next.Bold = true;
                    AppendInline(text, element.ChildNodes, next);
                    break;
                }
                case "em":
                case "i":
                {
                    var next = (PdfInlineFormat)format.Clone();
                    next.Italic = true;
                    AppendInline(text, element.ChildNodes, next);
                    break;
                }
                case "u":
                {
                    var next = (PdfInlineFormat)format.Clone();
                    next.Underline = true;
                    AppendInline(text, element.ChildNodes, next);
                    break;
                }
                case "s":
                case "strike":
                {
                    var next = (PdfInlineFormat)format.Clone();
                    next.Strike = true;
                    AppendInline(text, element.ChildNodes, next);
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
                    }
                    AppendInline(text, element.ChildNodes, next);
                    break;
                }
                case "code":
                {
                    var next = (PdfInlineFormat)format.Clone();
                    next.FontFamily = "Consolas";
                    AppendInline(text, element.ChildNodes, next);
                    break;
                }
                default:
                    AppendInline(text, element.ChildNodes, format);
                    break;
            }
        }
    }

    private static void AppendSpan(TextDescriptor text, string value, PdfInlineFormat format)
    {
        var span = text.Span(value);
        if (format.Bold) span.Bold();
        if (format.Italic) span.Italic();
        if (format.Underline) span.Underline();
        if (format.Strike) span.Strikethrough();
        if (!string.IsNullOrWhiteSpace(format.Color)) span.FontColor(format.Color);
        if (format.FontSize.HasValue) span.FontSize(format.FontSize.Value);
        var families = new[] { format.FontFamily, "Noto Sans JP" }
            .Where(family => !string.IsNullOrWhiteSpace(family))
            .Select(family => family!)
            .Distinct()
            .ToArray();
        if (families.Length > 0) span.FontFamily(families);
    }

    private static PdfInlineFormat ApplyCss(PdfInlineFormat current, string? style)
    {
        var next = (PdfInlineFormat)current.Clone();
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
                case "font-family":
                    next.FontFamily = value.Split(',')[0].Trim().Trim('\'', '"');
                    break;
                case "font-size":
                    if (TryParseCssFontSize(value, out var size)) next.FontSize = (float)size;
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
            "px" => number * 0.75,
            "em" or "rem" => 11 * number,
            _ => number,
        };
        return true;
    }

    private static string? NormalizeColor(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        value = value.Trim();

        if (value.StartsWith('#'))
        {
            var hex = value[1..];
            return hex.Length switch
            {
                3 => $"#{hex[0]}{hex[0]}{hex[1]}{hex[1]}{hex[2]}{hex[2]}",
                4 => $"#{hex[0]}{hex[0]}{hex[1]}{hex[1]}{hex[2]}{hex[2]}",
                6 => $"#{hex}",
                8 => $"#{hex[..6]}",
                _ => null,
            };
        }

        var rgb = Regex.Match(value, @"rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)", RegexOptions.IgnoreCase);
        if (rgb.Success)
        {
            return $"#{int.Parse(rgb.Groups[1].Value, CultureInfo.InvariantCulture):X2}{int.Parse(rgb.Groups[2].Value, CultureInfo.InvariantCulture):X2}{int.Parse(rgb.Groups[3].Value, CultureInfo.InvariantCulture):X2}";
        }

        return value.ToUpperInvariant() switch
        {
            "BLACK" => "#000000",
            "WHITE" => "#FFFFFF",
            "RED" => "#FF0000",
            "GREEN" => "#008000",
            "BLUE" => "#0000FF",
            _ => null,
        };
    }

    private static float HeadingSize(string tag) => tag switch
    {
        "h1" => 24,
        "h2" => 18,
        "h3" => 14,
        "h4" => 12,
        "h5" => 11,
        _ => 10.5f,
    };

    private static void EnsureFonts(string? fontDirectory)
    {
        if (_fontRegistered || string.IsNullOrWhiteSpace(fontDirectory) || !Directory.Exists(fontDirectory)) return;

        try
        {
            foreach (var path in Directory.EnumerateFiles(fontDirectory, "NotoSansJP-*.*")
                         .Where(path => path.EndsWith(".otf", StringComparison.OrdinalIgnoreCase) ||
                                        path.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase)))
            {
                using var stream = File.OpenRead(path);
                QuestPDF.Drawing.FontManager.RegisterFont(stream);
            }
            _fontRegistered = true;
        }
        catch
        {
            // A missing font should not prevent PDF export; Lato is used as a fallback.
        }
    }

    private sealed class PdfInlineFormat
    {
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strike { get; set; }
        public string? Color { get; set; }
        public string? FontFamily { get; set; }
        public float? FontSize { get; set; }

        public object Clone() => MemberwiseClone();
    }
}
