using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace WordPadLite.Services;

public sealed record PdfImportResult(string Html);

public static class PdfImporter
{
    private sealed class TextLine
    {
        public double Y { get; init; }
        public List<Word> Words { get; } = new();
    }

    public static PdfImportResult Import(Stream stream)
    {
        using var document = PdfDocument.Open(stream);
        var html = new StringBuilder();

        for (var pageNumber = 1; pageNumber <= document.NumberOfPages; pageNumber++)
        {
            var page = document.GetPage(pageNumber);
            var lines = BuildLines(page);
            if (lines.Count == 0)
            {
                var fallback = page.Text?.Trim();
                if (!string.IsNullOrEmpty(fallback))
                {
                    html.Append($"<p>{EscapeText(fallback).Replace("\n", "<br>")}</p>");
                }
            }
            else
            {
                html.Append("<p>");
                for (var i = 0; i < lines.Count; i++)
                {
                    if (i > 0) html.Append("<br>");
                    html.Append(EscapeText(lines[i]));
                }
                html.Append("</p>");
            }

            if (pageNumber < document.NumberOfPages)
            {
                html.Append("<hr class=\"page-break\"><p><br></p>");
            }
        }

        return new PdfImportResult(html.Length == 0 ? "<p><br></p>" : html.ToString());
    }

    private static List<string> BuildLines(Page page)
    {
        var textLines = new List<TextLine>();

        foreach (var word in page.GetWords())
        {
            var y = word.BoundingBox.Bottom;
            var line = textLines.FirstOrDefault(candidate => Math.Abs(candidate.Y - y) <= 2.5);
            if (line is null)
            {
                line = new TextLine { Y = y };
                textLines.Add(line);
            }
            line.Words.Add(word);
        }

        var result = new List<string>();
        foreach (var line in textLines.OrderByDescending(line => line.Y))
        {
            line.Words.Sort((left, right) => left.BoundingBox.Left.CompareTo(right.BoundingBox.Left));
            var text = JoinWords(line.Words);
            if (!string.IsNullOrWhiteSpace(text)) result.Add(text);
        }

        return result;
    }

    private static string JoinWords(IEnumerable<Word> words)
    {
        var builder = new StringBuilder();
        double? previousRight = null;

        foreach (var word in words)
        {
            if (string.IsNullOrEmpty(word.Text)) continue;
            if (previousRight is not null && word.BoundingBox.Left - previousRight.Value > 1.0)
            {
                builder.Append(' ');
            }
            builder.Append(word.Text);
            previousRight = word.BoundingBox.Right;
        }

        return builder.ToString();
    }

    private static string EscapeText(string value) =>
        value
            .Normalize(NormalizationForm.FormKC)
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;");
}
