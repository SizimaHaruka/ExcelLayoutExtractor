using System.Globalization;
using System.Net;
using System.Text;

namespace ExcelLayoutExtractor.Cli;

internal static class HtmlPreviewRenderer
{
    public static string Render(LayoutDocument document, string title)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang=\"ja\">");
        sb.AppendLine("<head>");
        sb.AppendLine("<meta charset=\"utf-8\" />");
        sb.AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />");
        sb.AppendLine($"<title>{WebUtility.HtmlEncode(title)} Preview</title>");
        sb.AppendLine("<style>");
        sb.AppendLine(":root { color-scheme: light; font-family: 'Segoe UI', 'Yu Gothic UI', sans-serif; }");
        sb.AppendLine("body { margin: 0; background: #f3f4f6; color: #111827; }");
        sb.AppendLine(".shell { padding: 24px; }");
        sb.AppendLine(".card { width: max-content; padding: 24px; background: white; border-radius: 16px; box-shadow: 0 18px 48px rgba(15, 23, 42, 0.14); }");
        sb.AppendLine(".meta { margin-bottom: 16px; font-size: 14px; color: #475569; }");
        sb.AppendLine(".page { position: relative; background: #fff; border: 1px solid #cbd5e1; overflow: hidden; }");
        sb.AppendLine(".label { position: absolute; box-sizing: border-box; overflow: hidden; white-space: pre-wrap; }");
        sb.AppendLine(".variable { border: 0.35mm dashed #2563eb; background: rgba(37, 99, 235, 0.08); color: #1d4ed8; }");
        sb.AppendLine(".image { position: absolute; box-sizing: border-box; border: 0.35mm dashed #64748b; background: rgba(148, 163, 184, 0.15); display: flex; align-items: center; justify-content: center; color: #475569; font-size: 10px; }");
        sb.AppendLine("</style>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.AppendLine("<div class=\"shell\">");
        sb.AppendLine("<div class=\"card\">");
        sb.AppendLine($"<div class=\"meta\">Page: {Format(document.Page.Width)}mm x {Format(document.Page.Height)}mm / Elements: {document.Elements.Count}</div>");
        sb.AppendLine($"<div class=\"page\" style=\"width:{Format(document.Page.Width)}mm;height:{Format(document.Page.Height)}mm;\">");

        foreach (var element in document.Elements)
        {
            switch (element.Type)
            {
                case "line":
                    RenderLine(sb, element);
                    break;
                case "rect":
                    RenderRect(sb, element);
                    break;
                case "text":
                    RenderText(sb, element, false);
                    break;
                case "variable":
                    RenderText(sb, element, true);
                    break;
                case "image":
                    RenderImage(sb, element);
                    break;
            }
        }

        sb.AppendLine("</div>");
        sb.AppendLine("</div>");
        sb.AppendLine("</div>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");
        return sb.ToString();
    }

    private static void RenderLine(StringBuilder sb, LayoutElement element)
    {
        var x1 = element.X1 ?? 0;
        var y1 = element.Y1 ?? 0;
        var x2 = element.X2 ?? 0;
        var y2 = element.Y2 ?? 0;
        var width = Math.Abs(x2 - x1);
        var height = Math.Abs(y2 - y1);
        var color = element.Style?.BorderColor ?? "#111827";
        var css = element.Style?.BorderStyle ?? "solid";

        if (Math.Abs(y1 - y2) < 0.01)
        {
            sb.AppendLine($"<div style=\"position:absolute;left:{Format(Math.Min(x1, x2))}mm;top:{Format(y1)}mm;width:{Format(Math.Max(width, 0.2))}mm;border-top:0.4mm {css} {color};\"></div>");
        }
        else
        {
            sb.AppendLine($"<div style=\"position:absolute;left:{Format(x1)}mm;top:{Format(Math.Min(y1, y2))}mm;height:{Format(Math.Max(height, 0.2))}mm;border-left:0.4mm {css} {color};\"></div>");
        }
    }

    private static void RenderRect(StringBuilder sb, LayoutElement element)
    {
        sb.AppendLine($"<div style=\"position:absolute;left:{Format(element.X ?? 0)}mm;top:{Format(element.Y ?? 0)}mm;width:{Format(element.Width ?? 0)}mm;height:{Format(element.Height ?? 0)}mm;border:0.4mm {element.Style?.BorderStyle ?? "solid"} {element.Style?.BorderColor ?? "#0f172a"};box-sizing:border-box;\"></div>");
    }

    private static void RenderText(StringBuilder sb, LayoutElement element, bool variable)
    {
        var className = variable ? "label variable" : "label";
        var text = variable ? $"{{{{{element.Key}}}}}" : element.Text ?? string.Empty;
        var style = element.Style;

        sb.AppendLine(
            $"<div class=\"{className}\" style=\"left:{Format(element.X ?? 0)}mm;top:{Format(element.Y ?? 0)}mm;width:{Format(element.Width ?? 0)}mm;height:{Format(element.Height ?? 0)}mm;" +
            $"font-family:{WebUtility.HtmlEncode(style?.FontFamily ?? "sans-serif")};font-size:{Format(style?.FontSize ?? 3.5)}mm;" +
            $"font-weight:{style?.FontWeight ?? "normal"};text-align:{style?.TextAlign ?? "left"};" +
            $"color:{style?.ForegroundColor ?? "#111827"};\">{WebUtility.HtmlEncode(text)}</div>");
    }

    private static void RenderImage(StringBuilder sb, LayoutElement element)
    {
        var caption = WebUtility.HtmlEncode(Path.GetFileName(element.Source ?? "image"));
        sb.AppendLine($"<div class=\"image\" style=\"left:{Format(element.X ?? 0)}mm;top:{Format(element.Y ?? 0)}mm;width:{Format(element.Width ?? 0)}mm;height:{Format(element.Height ?? 0)}mm;\">{caption}</div>");
    }

    private static string Format(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);
}
