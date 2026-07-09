using System;
using System.Linq;

namespace SalesforceRestAddin.Tables;

internal static class DescribeSheetName
{
    private const int MaxLength = 31;

    public static string Sanitize(string name)
    {
        var invalid = new[] { ':', '\\', '/', '?', '*', '[', ']' };
        var cleaned = new string(name.Select(c => invalid.Contains(c) ? '_' : c).ToArray()).Trim();
        if (string.IsNullOrWhiteSpace(cleaned))
        {
            cleaned = "Describe";
        }

        return cleaned.Length <= MaxLength ? cleaned : cleaned.Substring(0, MaxLength);
    }
}
