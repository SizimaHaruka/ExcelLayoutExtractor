using ClosedXML.Excel;

namespace ExcelLayoutExtractor.Cli;

internal sealed class ExcelLayoutExtractionService
{
    public LayoutDocument Extract(string workbookPath, string? worksheetName = null)
    {
        using var workbook = new XLWorkbook(workbookPath);
        var worksheet = ResolveWorksheet(workbook, worksheetName);
        var printRange = ResolvePrintRange(worksheet);
        var coordinateSpace = CoordinateSpace.Create(worksheet, printRange);

        var pageSize = PageSizeCatalog.Resolve(worksheet.PageSetup.PaperSize);
        var margins = worksheet.PageSetup.Margins;

        var document = new LayoutDocument
        {
            Page = new PageDefinition
            {
                Width = pageSize.WidthMm,
                Height = pageSize.HeightMm,
                MarginTop = UnitConversion.InchesToMillimeters(margins.Top),
                MarginRight = UnitConversion.InchesToMillimeters(margins.Right),
                MarginBottom = UnitConversion.InchesToMillimeters(margins.Bottom),
                MarginLeft = UnitConversion.InchesToMillimeters(margins.Left)
            },
            Elements = [],
            Assets = [],
            Warnings = []
        };

        document.Elements.AddRange(ExtractLines(printRange, coordinateSpace));
        document.Elements.AddRange(ExtractRectangles(printRange, coordinateSpace));
        document.Elements.AddRange(ExtractTextAndVariables(printRange, coordinateSpace, document.Warnings));
        document.Elements.AddRange(ExtractImages(worksheet, document.Assets));

        document.Elements.Sort(static (left, right) =>
        {
            var y = Nullable.Compare(left.Y ?? left.Y1, right.Y ?? right.Y1);
            return y != 0 ? y : Nullable.Compare(left.X ?? left.X1, right.X ?? right.X1);
        });

        return document;
    }

    private static IXLWorksheet ResolveWorksheet(XLWorkbook workbook, string? worksheetName)
    {
        if (!string.IsNullOrWhiteSpace(worksheetName))
        {
            return workbook.Worksheet(worksheetName);
        }

        return workbook.Worksheets.First();
    }

    private static IXLRange ResolvePrintRange(IXLWorksheet worksheet)
    {
        var printArea = worksheet.PageSetup.PrintAreas.FirstOrDefault();
        if (printArea is not null)
        {
            return printArea;
        }

        return worksheet.RangeUsed() ?? worksheet.Range("A1:A1");
    }

    private static IEnumerable<LayoutElement> ExtractLines(IXLRange range, CoordinateSpace coordinateSpace)
    {
        var horizontal = new Dictionary<LineKey, LineAccumulator>();
        var vertical = new Dictionary<LineKey, LineAccumulator>();

        foreach (var cell in range.CellsUsed(XLCellsUsedOptions.All))
        {
            var geometry = coordinateSpace.GetCell(cell.Address.RowNumber, cell.Address.ColumnNumber);
            var border = cell.Style.Border;

            AddLine(horizontal, border.TopBorder, border.TopBorderColor, geometry.Left, geometry.Top, geometry.Right);
            AddLine(horizontal, border.BottomBorder, border.BottomBorderColor, geometry.Left, geometry.Bottom, geometry.Right);
            AddLine(vertical, border.LeftBorder, border.LeftBorderColor, geometry.Top, geometry.Left, geometry.Bottom);
            AddLine(vertical, border.RightBorder, border.RightBorderColor, geometry.Top, geometry.Right, geometry.Bottom);
        }

        return MergeLines(horizontal.Values, vertical.Values);
    }

    private static IEnumerable<LayoutElement> MergeLines(IEnumerable<LineAccumulator> horizontal, IEnumerable<LineAccumulator> vertical)
    {
        foreach (var group in horizontal.GroupBy(line => new { line.Fixed, line.Color, line.Style, line.Weight }))
        {
            foreach (var merged in MergeCollinear(group, true))
            {
                yield return merged;
            }
        }

        foreach (var group in vertical.GroupBy(line => new { line.Fixed, line.Color, line.Style, line.Weight }))
        {
            foreach (var merged in MergeCollinear(group, false))
            {
                yield return merged;
            }
        }
    }

    private static IEnumerable<LayoutElement> MergeCollinear(IEnumerable<LineAccumulator> lines, bool isHorizontal)
    {
        var ordered = lines.OrderBy(line => line.Start).ToList();
        if (ordered.Count == 0)
        {
            yield break;
        }

        var current = ordered[0];
        for (var index = 1; index < ordered.Count; index++)
        {
            var next = ordered[index];
            if (Math.Abs(current.End - next.Start) <= 0.1)
            {
                current = current with { End = Math.Max(current.End, next.End) };
                continue;
            }

            yield return current.ToElement(isHorizontal);
            current = next;
        }

        yield return current.ToElement(isHorizontal);
    }

    private static void AddLine(
        Dictionary<LineKey, LineAccumulator> lines,
        XLBorderStyleValues borderStyle,
        XLColor borderColor,
        double start,
        double fixedCoordinate,
        double end)
    {
        if (borderStyle == XLBorderStyleValues.None)
        {
            return;
        }

        var style = BorderStyleMapper.ToCss(borderStyle);
        var weight = BorderStyleMapper.ToWeight(borderStyle);
        var color = ColorMapper.ToHtml(borderColor);
        var key = new LineKey(fixedCoordinate, start, end, style, color, weight);
        lines[key] = new LineAccumulator(key.Fixed, key.Start, key.End, key.Style, key.Color, key.Weight);
    }

    private static IEnumerable<LayoutElement> ExtractRectangles(IXLRange range, CoordinateSpace coordinateSpace)
    {
        foreach (var mergedRange in range.Worksheet.MergedRanges.Where(merged => merged.Intersects(range)))
        {
            var first = mergedRange.FirstCell();
            var border = first.Style.Border;
            if (border.TopBorder == XLBorderStyleValues.None &&
                border.BottomBorder == XLBorderStyleValues.None &&
                border.LeftBorder == XLBorderStyleValues.None &&
                border.RightBorder == XLBorderStyleValues.None)
            {
                continue;
            }

            var geometry = coordinateSpace.GetRange(mergedRange.RangeAddress.FirstAddress.RowNumber,
                mergedRange.RangeAddress.FirstAddress.ColumnNumber,
                mergedRange.RangeAddress.LastAddress.RowNumber,
                mergedRange.RangeAddress.LastAddress.ColumnNumber);

            yield return new LayoutElement
            {
                Type = "rect",
                X = geometry.Left,
                Y = geometry.Top,
                Width = geometry.Width,
                Height = geometry.Height,
                Origin = mergedRange.RangeAddress.ToStringRelative(includeSheet: false),
                Style = new StyleDefinition
                {
                    BorderColor = ColorMapper.ToHtml(border.TopBorderColor),
                    BorderStyle = BorderStyleMapper.ToCss(border.TopBorder)
                }
            };
        }
    }

    private static IEnumerable<LayoutElement> ExtractTextAndVariables(IXLRange range, CoordinateSpace coordinateSpace, List<string> warnings)
    {
        var mergedTopLefts = range.Worksheet.MergedRanges
            .Where(merged => merged.Intersects(range))
            .Select(merged => merged.FirstCell().Address.ToStringRelative())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var cell in range.CellsUsed(XLCellsUsedOptions.All))
        {
            var value = cell.GetFormattedString();
            if (string.IsNullOrWhiteSpace(value))
            {
                continue;
            }

            if (cell.IsMerged() && !mergedTopLefts.Contains(cell.Address.ToStringRelative()))
            {
                continue;
            }

            var bounds = cell.IsMerged()
                ? coordinateSpace.GetRange(cell.MergedRange().RangeAddress.FirstAddress.RowNumber,
                    cell.MergedRange().RangeAddress.FirstAddress.ColumnNumber,
                    cell.MergedRange().RangeAddress.LastAddress.RowNumber,
                    cell.MergedRange().RangeAddress.LastAddress.ColumnNumber)
                : coordinateSpace.GetCell(cell.Address.RowNumber, cell.Address.ColumnNumber);

            if (VariableTemplateParser.TryParse(value, out var token))
            {
                yield return BuildVariableElement(cell, bounds, token);
                continue;
            }

            if (value.Contains("{{", StringComparison.Ordinal) && value.Contains("}}", StringComparison.Ordinal))
            {
                warnings.Add($"Inline variable syntax is not fully supported yet: {cell.Address}");
            }

            yield return BuildTextElement(cell, bounds, value);
        }
    }

    private static LayoutElement BuildTextElement(IXLCell cell, Box bounds, string value)
    {
        var font = cell.Style.Font;
        var alignment = cell.Style.Alignment;

        return new LayoutElement
        {
            Type = "text",
            X = bounds.Left,
            Y = bounds.Top,
            Width = bounds.Width,
            Height = bounds.Height,
            Text = value,
            Origin = cell.Address.ToStringRelative(),
            Style = new StyleDefinition
            {
                FontFamily = font.FontName,
                FontSize = Math.Round(UnitConversion.PointsToMillimeters(font.FontSize), 2),
                FontWeight = font.Bold ? "bold" : "normal",
                TextAlign = AlignmentMapper.ToHorizontal(alignment.Horizontal),
                VerticalAlign = AlignmentMapper.ToVertical(alignment.Vertical),
                ForegroundColor = ColorMapper.ToHtml(font.FontColor),
                FillColor = ColorMapper.ToHtml(cell.Style.Fill.BackgroundColor)
            }
        };
    }

    private static LayoutElement BuildVariableElement(IXLCell cell, Box bounds, VariableToken token)
    {
        var font = cell.Style.Font;
        var alignment = cell.Style.Alignment;

        return new LayoutElement
        {
            Type = "variable",
            X = bounds.Left,
            Y = bounds.Top,
            Width = bounds.Width,
            Height = bounds.Height,
            Key = token.Key,
            Format = token.Format,
            Origin = cell.Address.ToStringRelative(),
            Style = new StyleDefinition
            {
                FontFamily = font.FontName,
                FontSize = Math.Round(UnitConversion.PointsToMillimeters(font.FontSize), 2),
                FontWeight = font.Bold ? "bold" : "normal",
                TextAlign = AlignmentMapper.ToHorizontal(alignment.Horizontal),
                VerticalAlign = AlignmentMapper.ToVertical(alignment.Vertical),
                ForegroundColor = ColorMapper.ToHtml(font.FontColor),
                FillColor = ColorMapper.ToHtml(cell.Style.Fill.BackgroundColor)
            }
        };
    }

    private static IEnumerable<LayoutElement> ExtractImages(IXLWorksheet worksheet, List<ImageAsset> assets)
    {
        var imageIndex = 0;
        foreach (var picture in worksheet.Pictures)
        {
            imageIndex++;
            var extension = ImageFormatMapper.ToExtension(picture.Format);
            var fileName = $"images/{SanitizeFileName(picture.Name, imageIndex)}.{extension}";
            using var memory = new MemoryStream();
            picture.ImageStream.Position = 0;
            picture.ImageStream.CopyTo(memory);

            assets.Add(new ImageAsset
            {
                FileName = fileName,
                Content = memory.ToArray()
            });

            yield return new LayoutElement
            {
                Type = "image",
                X = UnitConversion.PixelsToMillimeters(picture.Left),
                Y = UnitConversion.PixelsToMillimeters(picture.Top),
                Width = UnitConversion.PixelsToMillimeters(picture.Width),
                Height = UnitConversion.PixelsToMillimeters(picture.Height),
                Source = fileName,
                Origin = picture.TopLeftCell?.Address.ToStringRelative() ?? "picture"
            };
        }
    }

    private static string SanitizeFileName(string? input, int fallbackIndex)
    {
        var raw = string.IsNullOrWhiteSpace(input) ? $"image_{fallbackIndex:000}" : input;
        var invalid = Path.GetInvalidFileNameChars();
        var sanitized = new string(raw.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray()).Trim();
        return sanitized.Length == 0 ? $"image_{fallbackIndex:000}" : sanitized;
    }

    private readonly record struct LineKey(double Fixed, double Start, double End, string Style, string Color, double Weight);

    private readonly record struct LineAccumulator(double Fixed, double Start, double End, string Style, string Color, double Weight)
    {
        public LayoutElement ToElement(bool isHorizontal)
        {
            return new LayoutElement
            {
                Type = "line",
                X1 = isHorizontal ? Start : Fixed,
                Y1 = isHorizontal ? Fixed : Start,
                X2 = isHorizontal ? End : Fixed,
                Y2 = isHorizontal ? Fixed : End,
                Style = new StyleDefinition
                {
                    BorderStyle = Style,
                    BorderColor = Color
                }
            };
        }
    }
}
