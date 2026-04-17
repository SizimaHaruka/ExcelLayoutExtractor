using System.Text.RegularExpressions;

namespace ExcelLayoutExtractor.Cli;

internal static partial class VariableTemplateParser
{
    [GeneratedRegex(@"^\{\{\s*(?<body>.+?)\s*\}\}$", RegexOptions.Compiled | RegexOptions.CultureInvariant)]
    private static partial Regex FullVariableRegex();

    public static bool TryParse(string input, out VariableToken token)
    {
        token = default!;
        var match = FullVariableRegex().Match(input);
        if (!match.Success)
        {
            return false;
        }

        var body = match.Groups["body"].Value.Trim();
        if (body.Length == 0)
        {
            return false;
        }

        var separatorIndex = body.IndexOf('|');
        if (separatorIndex < 0)
        {
            token = new VariableToken(body, null);
            return true;
        }

        var key = body[..separatorIndex].Trim();
        var format = body[(separatorIndex + 1)..].Trim();
        if (key.Length == 0)
        {
            return false;
        }

        token = new VariableToken(key, format.Length == 0 ? null : format);
        return true;
    }
}
