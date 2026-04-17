using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

namespace ExcelLayoutExtractor.Cli;

internal sealed class CoordinateSpace
{
    private readonly Dictionary<int, double> _columnLefts;
    private readonly Dictionary<int, double> _rowTops;
    private readonly Dictionary<int, double> _columnWidths;
    private readonly Dictionary<int, double> _rowHeights;

    private CoordinateSpace(
        Dictionary<int, double> columnLefts,
        Dictionary<int, double> rowTops,
        Dictionary<int, double> columnWidths,
        Dictionary<int, double> rowHeights)
    {
        _columnLefts = columnLefts;
        _rowTops = rowTops;
        _columnWidths = columnWidths;
        _rowHeights = rowHeights;
    }

    public static CoordinateSpace Create(IXLWorksheet worksheet, IXLRange range)
    {
        var firstRow = range.RangeAddress.FirstAddress.RowNumber;
        var lastRow = range.RangeAddress.LastAddress.RowNumber;
        var firstColumn = range.RangeAddress.FirstAddress.ColumnNumber;
        var lastColumn = range.RangeAddress.LastAddress.ColumnNumber;

        var columnLefts = new Dictionary<int, double>();
        var rowTops = new Dictionary<int, double>();
        var columnWidths = new Dictionary<int, double>();
        var rowHeights = new Dictionary<int, double>();

        double left = 0;
        for (var columnNumber = firstColumn; columnNumber <= lastColumn; columnNumber++)
        {
            columnLefts[columnNumber] = left;
            var width = UnitConversion.ExcelColumnWidthToMillimeters(worksheet.Column(columnNumber).Width);
            columnWidths[columnNumber] = width;
            left += width;
        }

        double top = 0;
        for (var rowNumber = firstRow; rowNumber <= lastRow; rowNumber++)
        {
            rowTops[rowNumber] = top;
            var height = UnitConversion.PointsToMillimeters(worksheet.Row(rowNumber).Height);
            rowHeights[rowNumber] = height;
            top += height;
        }

        return new CoordinateSpace(columnLefts, rowTops, columnWidths, rowHeights);
    }

    public Box GetCell(int rowNumber, int columnNumber)
    {
        var left = _columnLefts.GetValueOrDefault(columnNumber, 0);
        var top = _rowTops.GetValueOrDefault(rowNumber, 0);
        var width = _columnWidths.GetValueOrDefault(columnNumber, 0);
        var height = _rowHeights.GetValueOrDefault(rowNumber, 0);
        return new Box(left, top, width, height);
    }

    public Box GetRange(int firstRow, int firstColumn, int lastRow, int lastColumn)
    {
        var first = GetCell(firstRow, firstColumn);
        var last = GetCell(lastRow, lastColumn);
        return new Box(first.Left, first.Top, (last.Left + last.Width) - first.Left, (last.Top + last.Height) - first.Top);
    }
}

internal readonly record struct Box(double Left, double Top, double Width, double Height)
{
    public double Right => Left + Width;
    public double Bottom => Top + Height;
}

internal static class UnitConversion
{
    public static double PointsToMillimeters(double points) => Math.Round(points * 25.4 / 72.0, 3);
    public static double InchesToMillimeters(double inches) => Math.Round(inches * 25.4, 3);
    public static double PixelsToMillimeters(double pixels) => Math.Round(pixels * 25.4 / 96.0, 3);

    public static double ExcelColumnWidthToMillimeters(double excelWidth)
    {
        var pixels = Math.Truncate(((256 * excelWidth + Math.Truncate(128.0 / 7.0)) / 256.0) * 7.0);
        return PixelsToMillimeters(pixels);
    }
}

internal static class BorderStyleMapper
{
    public static string ToCss(XLBorderStyleValues style) => style switch
    {
        XLBorderStyleValues.Dashed => "dashed",
        XLBorderStyleValues.Dotted => "dotted",
        XLBorderStyleValues.Double => "double",
        XLBorderStyleValues.Thick => "solid",
        XLBorderStyleValues.Medium => "solid",
        XLBorderStyleValues.MediumDashDot => "dashed",
        XLBorderStyleValues.MediumDashDotDot => "dashed",
        XLBorderStyleValues.MediumDashed => "dashed",
        XLBorderStyleValues.SlantDashDot => "dashed",
        XLBorderStyleValues.Hair => "solid",
        XLBorderStyleValues.Thin => "solid",
        _ => "solid"
    };

    public static double ToWeight(XLBorderStyleValues style) => style switch
    {
        XLBorderStyleValues.Hair => 0.2,
        XLBorderStyleValues.Thin => 0.4,
        XLBorderStyleValues.Medium => 0.8,
        XLBorderStyleValues.Thick => 1.2,
        XLBorderStyleValues.Double => 1.0,
        _ => 0.5
    };
}

internal static class AlignmentMapper
{
    public static string ToHorizontal(XLAlignmentHorizontalValues alignment) => alignment switch
    {
        XLAlignmentHorizontalValues.Center => "center",
        XLAlignmentHorizontalValues.Right => "right",
        XLAlignmentHorizontalValues.Justify => "justify",
        _ => "left"
    };

    public static string ToVertical(XLAlignmentVerticalValues alignment) => alignment switch
    {
        XLAlignmentVerticalValues.Bottom => "bottom",
        XLAlignmentVerticalValues.Center => "middle",
        XLAlignmentVerticalValues.Justify => "middle",
        _ => "top"
    };
}

internal static class ColorMapper
{
    public static string ToHtml(XLColor color)
    {
        if (color is null)
        {
            return "#000000";
        }

        var drawingColor = color.Color;
        return $"#{drawingColor.R:X2}{drawingColor.G:X2}{drawingColor.B:X2}";
    }
}

internal static class PageSizeCatalog
{
    public static (double WidthMm, double HeightMm) Resolve(XLPaperSize paperSize) => paperSize switch
    {
        XLPaperSize.A4Paper => (210, 297),
        XLPaperSize.A3Paper => (297, 420),
        XLPaperSize.A5Paper => (148, 210),
        XLPaperSize.B4Paper => (257, 364),
        XLPaperSize.B5Paper => (182, 257),
        XLPaperSize.LetterPaper => (215.9, 279.4),
        XLPaperSize.LegalPaper => (215.9, 355.6),
        _ => (210, 297)
    };
}

internal static class ImageFormatMapper
{
    public static string ToExtension(XLPictureFormat format) => format switch
    {
        XLPictureFormat.Bmp => "bmp",
        XLPictureFormat.Gif => "gif",
        XLPictureFormat.Icon => "ico",
        XLPictureFormat.Jpeg => "jpg",
        XLPictureFormat.Png => "png",
        XLPictureFormat.Tiff => "tiff",
        XLPictureFormat.Webp => "webp",
        _ => "bin"
    };
}
