using WordPadLite.Services;

var builder = WebApplication.CreateBuilder(args);
builder.WebHost.ConfigureKestrel(options =>
{
    options.Limits.MaxRequestBodySize = 50 * 1024 * 1024;
});

var app = builder.Build();

app.UseDefaultFiles();
app.UseStaticFiles();

app.MapPost("/api/docx/export", (DocxExportRequest request, ILogger<Program> logger) =>
{
    try
    {
        var bytes = DocxConverter.Export(request.Html, request.Title, request.Layout);
        var fileName = $"{SanitizeFileName(request.Title)}.docx";
        return Results.File(
            bytes,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            fileName);
    }
    catch (Exception exception)
    {
        logger.LogError(exception, "DOCX export failed.");
        return Results.BadRequest(new { error = "DOCX への書き出しに失敗しました。" });
    }
});

app.MapPost("/api/pdf/export", (PdfExportRequest request, ILogger<Program> logger) =>
{
    try
    {
        var webRoot = app.Environment.WebRootPath ?? Path.Combine(app.Environment.ContentRootPath, "wwwroot");
        var fontDirectory = Path.Combine(webRoot, "fonts");
        var bytes = PdfConverter.Export(request.Html, request.Title, request.Layout, fontDirectory);
        var fileName = $"{SanitizeFileName(request.Title)}.pdf";
        return Results.File(bytes, "application/pdf", fileName);
    }
    catch (Exception exception)
    {
        logger.LogError(exception, "PDF export failed.");
        return Results.BadRequest(new { error = "PDF への書き出しに失敗しました。" });
    }
});

app.MapPost("/api/pdf/import", async (HttpRequest request, ILogger<Program> logger) =>
{
    try
    {
        using var buffer = new MemoryStream();
        await request.Body.CopyToAsync(buffer);
        if (buffer.Length == 0)
        {
            return Results.BadRequest(new { error = "ファイルが空です。" });
        }

        buffer.Position = 0;
        var result = PdfImporter.Import(buffer);
        return Results.Json(result);
    }
    catch (Exception exception)
    {
        logger.LogError(exception, "PDF import failed.");
        return Results.BadRequest(new { error = "PDF を読み込めませんでした。" });
    }
});

app.MapPost("/api/docx/import", async (HttpRequest request, ILogger<Program> logger) =>
{
    try
    {
        using var buffer = new MemoryStream();
        await request.Body.CopyToAsync(buffer);
        if (buffer.Length == 0)
        {
            return Results.BadRequest(new { error = "ファイルが空です。" });
        }

        buffer.Position = 0;
        var result = DocxConverter.Import(buffer);
        return Results.Json(result);
    }
    catch (Exception exception)
    {
        logger.LogError(exception, "DOCX import failed.");
        return Results.BadRequest(new { error = "DOCX を読み込めませんでした。" });
    }
});

app.MapFallbackToFile("index.html");

app.Run();

static string SanitizeFileName(string? name)
{
    var value = string.IsNullOrWhiteSpace(name) ? "document" : name.Trim();
    var invalid = Path.GetInvalidFileNameChars();
    var sanitized = new string(value.Select(c => invalid.Contains(c) ? '_' : c).ToArray()).Trim();
    return string.IsNullOrWhiteSpace(sanitized) ? "document" : sanitized;
}
