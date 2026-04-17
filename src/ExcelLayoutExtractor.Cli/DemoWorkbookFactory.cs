using ClosedXML.Excel;

namespace ExcelLayoutExtractor.Cli;

internal static class DemoWorkbookFactory
{
    public static string Create(string outputDirectory)
    {
        Directory.CreateDirectory(outputDirectory);
        var path = Path.Combine(outputDirectory, "demo-order-form.xlsx");

        using var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("OrderForm");

        ws.PageSetup.PaperSize = XLPaperSize.A4Paper;
        ws.PageSetup.Margins.Top = 0.4;
        ws.PageSetup.Margins.Right = 0.4;
        ws.PageSetup.Margins.Bottom = 0.5;
        ws.PageSetup.Margins.Left = 0.4;

        ws.Column(1).Width = 4;
        ws.Column(2).Width = 12;
        ws.Column(3).Width = 18;
        ws.Column(4).Width = 18;
        ws.Column(5).Width = 18;
        ws.Column(6).Width = 18;
        ws.Column(7).Width = 18;
        ws.Column(8).Width = 18;

        ws.Row(1).Height = 18;
        ws.Row(2).Height = 20;

        ws.Range("B2:G2").Merge().Value = "注文書";
        ws.Range("B2:G2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Range("B2:G2").Style.Font.Bold = true;
        ws.Range("B2:G2").Style.Font.FontSize = 16;

        ws.Cell("B4").Value = "取引先";
        ws.Cell("C4").Value = "{{customer_name}}";
        ws.Range("C4:G4").Merge();
        ws.Cell("B5").Value = "発行日";
        ws.Cell("C5").Value = "{{issue_date|yyyy-MM-dd}}";

        ws.Range("B7:G7").Style.Fill.BackgroundColor = XLColor.FromHtml("#E2E8F0");
        ws.Cell("B7").Value = "No";
        ws.Cell("C7").Value = "品目";
        ws.Cell("E7").Value = "数量";
        ws.Cell("F7").Value = "単価";
        ws.Cell("G7").Value = "金額";

        for (var row = 8; row <= 12; row++)
        {
            ws.Cell(row, 2).Value = $"{{{{detail[{row - 8}].no}}}}";
            ws.Cell(row, 3).Value = $"{{{{detail[{row - 8}].item}}}}";
            ws.Range(row, 3, row, 4).Merge();
            ws.Cell(row, 5).Value = $"{{{{detail[{row - 8}].quantity|#,##0}}}}";
            ws.Cell(row, 6).Value = $"{{{{detail[{row - 8}].unit_price|#,##0}}}}";
            ws.Cell(row, 7).Value = $"{{{{detail[{row - 8}].amount|#,##0}}}}";
        }

        ws.Cell("F14").Value = "合計";
        ws.Cell("G14").Value = "{{total_amount|#,##0}}";

        ws.Range("B4:G14").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Range("B4:G14").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        ws.Range("B2:G2").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
        ws.Range("B4:C5").Style.Fill.BackgroundColor = XLColor.FromHtml("#F8FAFC");
        ws.Range("F14:G14").Style.Fill.BackgroundColor = XLColor.FromHtml("#FEF3C7");

        ws.PageSetup.PrintAreas.Add("B2:G14");
        workbook.SaveAs(path);
        return path;
    }
}
