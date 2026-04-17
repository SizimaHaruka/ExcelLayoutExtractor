using System.Text;

namespace ExcelLayoutExtractor.Cli;

internal static class Program
{
    public static int Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;

        try
        {
            if (args.Length == 0 || HasFlag(args, "--help") || HasFlag(args, "-h"))
            {
                PrintUsage();
                return 0;
            }

            if (string.Equals(args[0], "demo", StringComparison.OrdinalIgnoreCase))
            {
                return RunDemo(args);
            }

            return RunExtract(args);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[error] {ex.Message}");
            return 1;
        }
    }

    private static int RunDemo(IReadOnlyList<string> args)
    {
        var outputDirectory = GetOptionValue(args, "--output-dir")
            ?? Path.Combine(Environment.CurrentDirectory, "samples", "demo-output");
        var workbookPath = DemoWorkbookFactory.Create(outputDirectory);

        Console.WriteLine($"[info] Demo workbook created: {workbookPath}");
        return RunExtract(["extract", workbookPath, "--output-dir", outputDirectory]);
    }

    private static int RunExtract(IReadOnlyList<string> args)
    {
        var commandOffset = string.Equals(args[0], "extract", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
        if (args.Count <= commandOffset)
        {
            throw new InvalidOperationException("Excel ファイルのパスを指定してください。");
        }

        var workbookPath = Path.GetFullPath(args[commandOffset]);
        if (!File.Exists(workbookPath))
        {
            throw new FileNotFoundException("Excel ファイルが見つかりません。", workbookPath);
        }

        var worksheetName = GetOptionValue(args, "--sheet");
        var outputDirectory = GetOptionValue(args, "--output-dir")
            ?? Path.Combine(Path.GetDirectoryName(workbookPath)!, "output");

        var extractor = new ExcelLayoutExtractionService();
        var document = extractor.Extract(workbookPath, worksheetName);

        var writer = new LayoutOutputWriter();
        var output = writer.Write(outputDirectory, Path.GetFileNameWithoutExtension(workbookPath), document);

        Console.WriteLine($"[ok] JSON: {output.JsonPath}");
        Console.WriteLine($"[ok] HTML: {output.HtmlPath}");

        if (!string.IsNullOrWhiteSpace(output.WarningPath))
        {
            Console.WriteLine($"[warn] Warnings: {output.WarningPath}");
        }

        Console.WriteLine($"[info] Elements: {document.Elements.Count}");
        return 0;
    }

    private static string? GetOptionValue(IReadOnlyList<string> args, string name)
    {
        for (var i = 0; i < args.Count - 1; i++)
        {
            if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase))
            {
                return args[i + 1];
            }
        }

        return null;
    }

    private static bool HasFlag(IReadOnlyList<string> args, string flag)
        => args.Any(arg => string.Equals(arg, flag, StringComparison.OrdinalIgnoreCase));

    private static void PrintUsage()
    {
        Console.WriteLine("ExcelLayoutExtractor");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  extract <template.xlsx> [--sheet <name>] [--output-dir <dir>]");
        Console.WriteLine("  demo [--output-dir <dir>]");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  extract .\\samples\\invoice.xlsx --output-dir .\\artifacts");
        Console.WriteLine("  demo --output-dir .\\samples\\demo-output");
    }
}
