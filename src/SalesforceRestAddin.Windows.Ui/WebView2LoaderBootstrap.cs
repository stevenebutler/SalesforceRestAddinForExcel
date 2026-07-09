using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using SalesforceRestAddin.Core.Session;
using Microsoft.Web.WebView2.Core;

namespace SalesforceRestAddin.Windows.Ui;

/// <summary>
/// Extracts architecture-specific <c>WebView2Loader.dll</c> from the embedded resource in
/// <c>SalesforceRestAddin.Windows.Ui</c> into <c>%LOCALAPPDATA%\SalesforceRestAddin\native\</c>, then
/// points WebView2 at that folder. Skips rewrite when size and stamped write-time already match.
/// </summary>
public static class WebView2LoaderBootstrap
{
    private const string LoaderFileName = "WebView2Loader.dll";
    private static string? _loaderFolder;
    private static bool _configured;

    public static void EnsureDeployed()
    {
        if (!string.IsNullOrEmpty(_loaderFolder))
        {
            return;
        }

        var archFolder = ProcessArchitectureFolder();
        var targetFolder = SalesforceRestAddinDataPaths.WebView2NativeLoaderFolder(archFolder);
        var targetFile = Path.Combine(targetFolder, LoaderFileName);
        var resourceName = $"SalesforceRestAddin.WebView2Loader.{archFolder}.dll";

        SalesforceRestAddinDataPaths.EnsureDataDirectory();
        Directory.CreateDirectory(targetFolder);

        using var resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName)
            ?? throw new InvalidOperationException($"Embedded WebView2 loader resource was not found: {resourceName}.");

        var expectedLength = resourceStream.Length;
        var stampUtc = GetLoaderStampUtc();

        if (!IsCurrentLoader(targetFile, expectedLength, stampUtc))
        {
            // Copy to a temp file then replace so a partial write cannot leave a "current" stamp.
            var tempFile = targetFile + ".tmp";
            try
            {
                using (var output = File.Create(tempFile))
                {
                    resourceStream.CopyTo(output);
                }

                File.SetLastWriteTimeUtc(tempFile, stampUtc);
                if (File.Exists(targetFile))
                {
                    File.Delete(targetFile);
                }

                File.Move(tempFile, targetFile);
                // Re-apply after move — some volumes rewrite timestamps on Move.
                File.SetLastWriteTimeUtc(targetFile, stampUtc);
            }
            finally
            {
                if (File.Exists(tempFile))
                {
                    try
                    {
                        File.Delete(tempFile);
                    }
                    catch
                    {
                        // Best-effort cleanup.
                    }
                }
            }
        }

        _loaderFolder = targetFolder;
    }

    public static void EnsureConfigured()
    {
        EnsureDeployed();

        if (_configured || string.IsNullOrEmpty(_loaderFolder))
        {
            return;
        }

        CoreWebView2Environment.SetLoaderDllFolderPath(_loaderFolder);
        _configured = true;
    }

    /// <summary>
    /// Deterministic UTC stamp from the managed WebView2 assembly version so a NuGet bump
    /// forces re-extract even when the on-disk file length happens to match.
    /// </summary>
    internal static DateTime GetLoaderStampUtc()
    {
        var version = typeof(CoreWebView2Environment).Assembly.GetName().Version
                      ?? new Version(0, 0, 0, 0);
        // Encode Build/Revision (and Major/Minor) into whole seconds — unique for WebView2 versioning.
        var seconds = (((long)version.Major * 1_000_000_000L)
                       + ((long)version.Minor * 1_000_000L)
                       + ((long)version.Build * 1_000L)
                       + version.Revision);
        return new DateTime(2001, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(seconds);
    }

    private static bool IsCurrentLoader(string targetFile, long expectedLength, DateTime stampUtc)
    {
        if (!File.Exists(targetFile))
        {
            return false;
        }

        var info = new FileInfo(targetFile);
        if (info.Length != expectedLength)
        {
            return false;
        }

        // Compare at second resolution — NTFS stores higher precision but we stamp whole seconds.
        var onDisk = info.LastWriteTimeUtc;
        return onDisk.Year == stampUtc.Year
               && onDisk.Month == stampUtc.Month
               && onDisk.Day == stampUtc.Day
               && onDisk.Hour == stampUtc.Hour
               && onDisk.Minute == stampUtc.Minute
               && onDisk.Second == stampUtc.Second;
    }

    private static string ProcessArchitectureFolder() =>
        RuntimeInformation.ProcessArchitecture switch
        {
            Architecture.X86 => "win-x86",
            Architecture.Arm64 => "win-arm64",
            _ => "win-x64",
        };
}
