using System;

namespace SalesforceRestAddin.Installer;

internal static class ExcelStartupRegistration
{
    private const string SalesforceRestAddinFilePrefix = "SalesforceRestAddin";

    public static bool IsOpenValueName(string name) => name.Equals("OPEN", StringComparison.OrdinalIgnoreCase)
        || (name.StartsWith("OPEN", StringComparison.OrdinalIgnoreCase) && name.Substring(4).All(char.IsDigit));

    public static bool IsSalesforceRestAddinXll(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        var searchStart = 0;
        while (searchStart < value.Length)
        {
            var extensionIndex = value.IndexOf(".xll", searchStart, StringComparison.OrdinalIgnoreCase);
            if (extensionIndex < 0)
            {
                return false;
            }

            var separatorIndex = Math.Max(
                value.LastIndexOf('\\', extensionIndex),
                value.LastIndexOf('/', extensionIndex));
            var filenameStart = separatorIndex + 1;
            var filenameLength = extensionIndex - filenameStart;
            if (filenameLength > 0
                && value.Substring(filenameStart, filenameLength).StartsWith(SalesforceRestAddinFilePrefix, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            searchStart = extensionIndex + 4;
        }

        return false;
    }

    public static bool IsLegacyForceConnectorAddinKey(string keyName) =>
        string.Equals(keyName, "ForceConnector", StringComparison.OrdinalIgnoreCase);
}
