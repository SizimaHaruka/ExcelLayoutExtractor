using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelLayoutExtractor.Cli;

internal sealed class LayoutOutputWriter
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    public OutputPaths Write(string outputDirectory, string baseName, LayoutDocument document)
    {
        Directory.CreateDirectory(outputDirectory);
        Directory.CreateDirectory(Path.Combine(outputDirectory, "images"));

        var jsonPath = Path.Combine(outputDirectory, $"{baseName}.layout.json");
        var htmlPath = Path.Combine(outputDirectory, $"{baseName}.preview.html");

        File.WriteAllText(jsonPath, JsonSerializer.Serialize(document, JsonOptions), Encoding.UTF8);
        File.WriteAllText(htmlPath, HtmlPreviewRenderer.Render(document, baseName), Encoding.UTF8);

        foreach (var asset in document.Assets)
        {
            var assetPath = Path.Combine(outputDirectory, asset.FileName.Replace('/', Path.DirectorySeparatorChar));
            var assetDirectory = Path.GetDirectoryName(assetPath);
            if (!string.IsNullOrWhiteSpace(assetDirectory))
            {
                Directory.CreateDirectory(assetDirectory);
            }

            File.WriteAllBytes(assetPath, asset.Content);
        }

        var warningPath = string.Empty;
        if (document.Warnings.Count > 0)
        {
            warningPath = Path.Combine(outputDirectory, $"{baseName}.warnings.txt");
            File.WriteAllLines(warningPath, document.Warnings, Encoding.UTF8);
        }

        return new OutputPaths(jsonPath, htmlPath, warningPath);
    }
}

internal sealed record OutputPaths(string JsonPath, string HtmlPath, string WarningPath);
