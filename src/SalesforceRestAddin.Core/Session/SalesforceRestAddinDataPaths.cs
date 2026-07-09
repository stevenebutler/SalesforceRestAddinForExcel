using System;
using System.IO;

namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Stable per-user data locations under %LOCALAPPDATA%\SalesforceRestAddin.
/// Independent of where the add-in or spike exe is deployed from.
/// </summary>
public static class SalesforceRestAddinDataPaths
{
    public static string DataDirectory =>
        Path.Combine(GetLocalAppDataDirectory(), "SalesforceRestAddin");

    public static string PreferencesFile =>
        Path.Combine(DataDirectory, "login-preferences.json");

    public static string ConnectorOptionsFile =>
        Path.Combine(DataDirectory, "connector-options.json");

    public static string WebView2UserDataFolder =>
        Path.Combine(DataDirectory, "WebView2");

    public static string SessionTraceLogFile =>
        Path.Combine(DataDirectory, "session-trace.log");

    /// <summary>On-disk Salesforce metadata cache (object list + describe JSON).</summary>
    public static string MetadataCacheDirectory =>
        Path.Combine(DataDirectory, "metadata-cache");

    /// <summary>Extracted <c>WebView2Loader.dll</c> for the current process architecture.</summary>
    public static string WebView2NativeLoaderFolder(string processArchitectureFolder) =>
        Path.Combine(DataDirectory, "native", processArchitectureFolder, "native");

    public static void EnsureDataDirectory() =>
        Directory.CreateDirectory(DataDirectory);

    private static string GetLocalAppDataDirectory()
    {
        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (string.IsNullOrWhiteSpace(localAppData))
        {
            throw new InvalidOperationException("Local application data folder is not available.");
        }

        return localAppData;
    }
}
