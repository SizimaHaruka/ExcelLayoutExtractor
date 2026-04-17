using System.Text.Json.Serialization;

namespace ExcelLayoutExtractor.Cli;

internal sealed class LayoutDocument
{
    public required PageDefinition Page { get; init; }
    public required List<LayoutElement> Elements { get; init; }

    [JsonIgnore]
    public List<ImageAsset> Assets { get; init; } = [];

    [JsonIgnore]
    public List<string> Warnings { get; init; } = [];
}

internal sealed class PageDefinition
{
    public required double Width { get; init; }
    public required double Height { get; init; }
    public string Unit { get; init; } = "mm";
    public double MarginTop { get; init; }
    public double MarginRight { get; init; }
    public double MarginBottom { get; init; }
    public double MarginLeft { get; init; }
}

internal sealed class LayoutElement
{
    public required string Type { get; init; }
    public double? X { get; init; }
    public double? Y { get; init; }
    public double? Width { get; init; }
    public double? Height { get; init; }
    public double? X1 { get; init; }
    public double? Y1 { get; init; }
    public double? X2 { get; init; }
    public double? Y2 { get; init; }
    public string? Text { get; init; }
    public string? Key { get; init; }
    public string? Format { get; init; }
    public string? Source { get; set; }
    public StyleDefinition? Style { get; init; }
    public string? Origin { get; init; }
}

internal sealed class StyleDefinition
{
    public string? FontFamily { get; init; }
    public double? FontSize { get; init; }
    public string? FontWeight { get; init; }
    public string? TextAlign { get; init; }
    public string? VerticalAlign { get; init; }
    public string? BorderColor { get; init; }
    public string? BorderStyle { get; init; }
    public string? FillColor { get; init; }
    public string? ForegroundColor { get; init; }
}

internal sealed class ImageAsset
{
    public required string FileName { get; init; }
    public required byte[] Content { get; init; }
}

internal sealed record VariableToken(string Key, string? Format);
