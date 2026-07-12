using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace SalesforceRestAddin.Installer;

internal sealed class InstallerService
{
    private const string ExcelOptionsSuffix = @"\Excel\Options";
    private const string ExcelAddinsPath = @"SOFTWARE\Microsoft\Office\Excel\Addins";
    private const string UninstallPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
    private const string XllVersionMarker = "SalesforceRestAddin-Version: ";
    private readonly string _installDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "SalesforceRestAddin");

    public Task<InstallerState> LoadStateAsync(Action<string>? log = null) =>
        Task.Run(() => LoadState(log));

    public Task<InstallResult> InstallAsync(Action<string>? progress = null, Action<string>? log = null) =>
        Task.Run(() => Install(progress, log));

    public Task<RemoveResult> RemoveAsync(Action<string>? log = null) =>
        Task.Run(() => Remove(log));

    public Task<ForceConnectorUninstallResult> RunForceConnectorUninstallerAsync(ForceConnectorUninstallEntry entry, Action<string>? log = null) =>
        Task.Run(() => RunForceConnectorUninstaller(entry, log));

    private InstallerState LoadState(Action<string>? log)
    {
        var bitness = DetectExcelBitness();
        var assetName = AssetNameFor(bitness);
        var bundledPath = BundledAssetPath(assetName);
        var installedPath = Path.Combine(_installDirectory, assetName);
        var installed = File.Exists(installedPath);
        var installedVersion = installed ? ReadXllVersion(installedPath) : null;
        var bundledVersion = bundledPath is null ? null : ReadXllVersion(bundledPath);
        GitHubRelease? latest = null;
        string? githubError = null;

        if (bundledPath is not null)
        {
            Log(log, $"Using co-located package: {bundledPath}");
        }
        else
        {
            try
            {
                Log(log, "Checking GitHub latest release.");
                latest = GetLatestRelease();
                Log(log, $"GitHub latest build: {latest.TargetCommitish}.");
            }
            catch (Exception ex)
            {
                githubError = ex.Message;
                Log(log, $"GitHub latest-release check failed: {ex.Message}");
            }
        }

        var status = GetInstallStatus(installed, installedVersion, bundledPath, bundledVersion, latest, githubError);
        var forceConnectorEntries = FindForceConnectorUninstallEntries().ToList();
        var forceConnectorExcelRegistrations = FindEnabledForceConnectorExcelRegistrations().ToList();

        return new InstallerState
        {
            ExcelBitness = bitness,
            AssetName = assetName,
            InstallDirectory = _installDirectory,
            BundledAssetPath = bundledPath,
            InstalledVersion = installedVersion,
            AvailableVersion = bundledPath is not null ? bundledVersion : latest?.TargetCommitish,
            IsInstalled = installed,
            IsUpdateAvailable = status.IsUpdateAvailable,
            Status = status.Status,
            GitHubCheckError = githubError,
            ForceConnectorUninstallEntries = forceConnectorEntries,
            ForceConnectorExcelRegistrations = forceConnectorExcelRegistrations,
        };
    }

    private InstallResult Install(Action<string>? progress, Action<string>? log)
    {
        var bitness = DetectExcelBitness();
        var assetName = AssetNameFor(bitness);
        var bundledPath = BundledAssetPath(assetName);
        var targetPath = Path.Combine(_installDirectory, assetName);

        Directory.CreateDirectory(_installDirectory);
        DisableForceConnectorExcelRegistrations(log);
        RemoveSalesforceRestAddinRegistrations(log);

        if (bundledPath is not null)
        {
            progress?.Invoke("Installing bundled local XLL...");
            Log(log, $"Copying bundled XLL from {bundledPath} to {targetPath}.");
            CopyFile(bundledPath, targetPath);
            RegisterExcelAddin(targetPath, log);
            WriteManifest(new InstallManifest
            {
                Repository = InstallerBranding.Repository,
                ReleaseTag = "bundled-local",
                TargetCommitish = ReadXllVersion(bundledPath) ?? "unknown",
                AssetName = assetName,
                InstallDirectory = _installDirectory,
                XllPath = targetPath,
                InstalledAtUtc = DateTimeOffset.UtcNow,
            });

            return new InstallResult(assetName, targetPath, "Bundled local package", false);
        }

        progress?.Invoke("Checking GitHub latest release...");
        var release = GetLatestRelease();
        var asset = release.Assets.FirstOrDefault(candidate => string.Equals(candidate.Name, assetName, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"Asset '{assetName}' was not found in the latest GitHub release.");

        var alreadyCurrent = XllHasVersion(targetPath, release.TargetCommitish);

        if (!alreadyCurrent)
        {
            progress?.Invoke($"Downloading {assetName}...");
            var tempPath = targetPath + ".download";
            Log(log, $"Downloading {asset.DownloadUrl}.");
            DownloadFile(asset.DownloadUrl, tempPath);
            progress?.Invoke("Writing local install...");
            ReplaceFile(tempPath, targetPath);
            Log(log, $"Downloaded XLL to {targetPath}.");
        }
        else
        {
            Log(log, "Installed XLL matches GitHub latest release; re-registering it without downloading.");
        }

        RegisterExcelAddin(targetPath, log);
        WriteManifest(new InstallManifest
        {
            Repository = InstallerBranding.Repository,
            ReleaseTag = release.TagName,
            TargetCommitish = release.TargetCommitish,
            AssetName = assetName,
            InstallDirectory = _installDirectory,
            XllPath = targetPath,
            InstalledAtUtc = DateTimeOffset.UtcNow,
        });

        return new InstallResult(assetName, targetPath, release.TargetCommitish, !alreadyCurrent);
    }

    private RemoveResult Remove(Action<string>? log)
    {
        var removedRegistry = RemoveSalesforceRestAddinRegistrations(log);
        var removedXlls = new List<string>();
        if (Directory.Exists(_installDirectory))
        {
            foreach (var xllPath in Directory.EnumerateFiles(_installDirectory, "SalesforceRestAddin*.xll", SearchOption.TopDirectoryOnly))
            {
                File.Delete(xllPath);
                removedXlls.Add(xllPath);
                Log(log, $"Removed installed XLL {xllPath}.");
            }
        }

        var message = removedRegistry.Count > 0 || removedXlls.Count > 0
            ? $"Removed {InstallerBranding.ProductName} registrations and XLL files. Your settings were kept."
            : "No Salesforce REST Add-in startup registration or local install was found.";
        return new RemoveResult(message);
    }

    private static ForceConnectorUninstallResult RunForceConnectorUninstaller(ForceConnectorUninstallEntry entry, Action<string>? log)
    {
        if (string.IsNullOrWhiteSpace(entry.UninstallCommand))
        {
            throw new InvalidOperationException($"No uninstall command is registered for {entry.DisplayName}.");
        }

        Log(log, $"Launching registered uninstaller for {entry.DisplayName}: {entry.UninstallCommand}");
        var commandProcessor = Environment.GetEnvironmentVariable("ComSpec") ?? "cmd.exe";
        using var process = Process.Start(new ProcessStartInfo
        {
            FileName = commandProcessor,
            Arguments = "/c " + entry.UninstallCommand,
            UseShellExecute = true,
        }) ?? throw new InvalidOperationException("Windows did not start the ForceConnector uninstaller.");
        process.WaitForExit();
        Log(log, $"ForceConnector uninstaller exited with code {process.ExitCode}.");
        return new ForceConnectorUninstallResult(process.ExitCode);
    }

    private static GitHubRelease GetLatestRelease()
    {
        var request = WebRequest.CreateHttp($"https://api.github.com/repos/{InstallerBranding.Repository}/releases/latest");
        request.Method = "GET";
        request.UserAgent = "SalesforceRestAddin-Installer";
        request.Accept = "application/vnd.github+json";

        using var response = (HttpWebResponse)request.GetResponse();
        using var stream = response.GetResponseStream()
            ?? throw new InvalidOperationException("GitHub release response did not contain a body.");
        var serializer = new DataContractJsonSerializer(typeof(GitHubRelease));
        return (GitHubRelease?)serializer.ReadObject(stream)
            ?? throw new InvalidOperationException("GitHub release payload could not be parsed.");
    }

    private static void DownloadFile(string url, string targetPath)
    {
        using var client = new WebClient();
        client.Headers.Add(HttpRequestHeader.UserAgent, "SalesforceRestAddin-Installer");
        client.DownloadFile(url, targetPath);
    }

    private static void CopyFile(string sourcePath, string targetPath)
    {
        if (string.Equals(Path.GetFullPath(sourcePath), Path.GetFullPath(targetPath), StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        File.Copy(sourcePath, targetPath, overwrite: true);
    }

    private static void ReplaceFile(string sourcePath, string targetPath)
    {
        if (File.Exists(targetPath))
        {
            File.Delete(targetPath);
        }

        File.Move(sourcePath, targetPath);
    }

    private static string DetectExcelBitness()
    {
        var c2r = ReadOfficePlatform();
        if (!string.IsNullOrWhiteSpace(c2r))
        {
            return c2r!;
        }

        var exePath = ReadExcelExePath();
        if (string.IsNullOrWhiteSpace(exePath) || !File.Exists(exePath))
        {
            throw new InvalidOperationException("Could not locate Excel on this machine. Is Excel installed?");
        }

        var bytes = File.ReadAllBytes(exePath);
        var peOffset = BitConverter.ToInt32(bytes, 0x3C);
        var machine = BitConverter.ToUInt16(bytes, peOffset + 4);
        return machine switch
        {
            0x8664 => "x64",
            0x014c => "x86",
            _ => throw new InvalidOperationException($"Unrecognized Excel binary architecture 0x{machine:X}."),
        };
    }

    private static string? ReadOfficePlatform()
    {
        foreach (var view in RegistryViews())
        {
            using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view);
            using var key = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration");
            if (key?.GetValue("Platform") is string platform && !string.IsNullOrWhiteSpace(platform))
            {
                return platform;
            }
        }

        return null;
    }

    private static string? ReadExcelExePath()
    {
        foreach (var view in RegistryViews())
        {
            using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view);
            using var key = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE");
            if (key?.GetValue(string.Empty) is string path && !string.IsNullOrWhiteSpace(path))
            {
                return path;
            }
        }

        return null;
    }

    private static InstallStatus GetInstallStatus(
        bool installed,
        string? installedVersion,
        string? bundledPath,
        string? bundledVersion,
        GitHubRelease? latest,
        string? githubError)
    {
        if (bundledPath is not null)
        {
            var versionMatches = VersionMatches(installedVersion, bundledVersion);
            return new InstallStatus(!versionMatches && installed, installed && versionMatches
                ? $"Latest bundled version {bundledVersion} is already installed. Install / Update can repair the installation."
                : installed
                    ? $"Bundled version {DisplayVersion(bundledVersion)} differs from installed version {DisplayVersion(installedVersion)}."
                : "A bundled local XLL is ready to install without an internet connection.");
        }

        if (latest is not null)
        {
            var updateAvailable = installed && !VersionMatches(installedVersion, latest.TargetCommitish);
            return new InstallStatus(updateAvailable, installed
                ? updateAvailable
                    ? $"GitHub version {latest.TargetCommitish} differs from installed version {DisplayVersion(installedVersion)}."
                    : $"Latest GitHub version {latest.TargetCommitish} is already installed. Install / Update can repair the installation."
                : "No local per-user install detected.");
        }

        return new InstallStatus(false, installed
            ? "Installed build found. GitHub could not be checked."
            : $"No local per-user install detected. GitHub could not be checked: {githubError ?? "unknown error"}");
    }

    private void WriteManifest(InstallManifest manifest)
    {
        Directory.CreateDirectory(_installDirectory);
        using var stream = File.Create(ManifestPath());
        new DataContractJsonSerializer(typeof(InstallManifest)).WriteObject(stream, manifest);
    }

    private string ManifestPath() => Path.Combine(_installDirectory, "installer-manifest.json");

    private static string AssetNameFor(string bitness) => bitness == "x64"
        ? "SalesforceRestAddin64-packed.xll"
        : "SalesforceRestAddin-packed.xll";

    private static string? BundledAssetPath(string assetName)
    {
        var candidate = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, assetName);
        return File.Exists(candidate) ? candidate : null;
    }

    private static bool XllHasVersion(string xllPath, string expectedVersion) =>
        VersionMatches(ReadXllVersion(xllPath), expectedVersion);

    private static bool VersionMatches(string? left, string? right) =>
        !string.IsNullOrWhiteSpace(left)
        && !string.IsNullOrWhiteSpace(right)
        && string.Equals(left, right, StringComparison.OrdinalIgnoreCase);

    private static string DisplayVersion(string? version) => string.IsNullOrWhiteSpace(version) ? "unknown" : version!;

    private static string? ReadXllVersion(string xllPath)
    {
        if (!File.Exists(xllPath))
        {
            return null;
        }

        const int tailLength = 4096;
        using var stream = File.OpenRead(xllPath);
        stream.Seek(-Math.Min(stream.Length, tailLength), SeekOrigin.End);
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: false);
        var text = reader.ReadToEnd();
        var markerIndex = text.LastIndexOf(XllVersionMarker, StringComparison.Ordinal);
        if (markerIndex < 0)
        {
            return null;
        }

        var versionStart = markerIndex + XllVersionMarker.Length;
        var versionEnd = text.IndexOfAny(new[] { '\r', '\n', '\0' }, versionStart);
        var version = (versionEnd < 0 ? text.Substring(versionStart) : text.Substring(versionStart, versionEnd - versionStart)).Trim();
        return string.IsNullOrWhiteSpace(version) ? null : version;
    }

    private IEnumerable<ForceConnectorUninstallEntry> FindForceConnectorUninstallEntries()
    {
        var results = new List<ForceConnectorUninstallEntry>();
        foreach (var hive in new[] { RegistryHive.LocalMachine, RegistryHive.CurrentUser })
        {
            foreach (var view in RegistryViews())
            {
                using var baseKey = RegistryKey.OpenBaseKey(hive, view);
                using var uninstall = baseKey.OpenSubKey(UninstallPath);
                if (uninstall is null)
                {
                    continue;
                }

                foreach (var name in uninstall.GetSubKeyNames())
                {
                    using var key = uninstall.OpenSubKey(name);
                    var displayName = key?.GetValue("DisplayName") as string;
                    var uninstallCommand = key?.GetValue("UninstallString") as string;
                    if (!IsForceConnectorProduct(displayName) || string.IsNullOrWhiteSpace(uninstallCommand))
                    {
                        continue;
                    }

                    results.Add(new ForceConnectorUninstallEntry(
                        displayName!,
                        key?.GetValue("DisplayVersion") as string,
                        key?.GetValue("Publisher") as string,
                        hive == RegistryHive.LocalMachine ? "All users" : "Current user",
                        view == RegistryView.Registry64 ? "64-bit" : "32-bit",
                        uninstallCommand!));
                }
            }
        }

        return results
            .GroupBy(entry => entry.DisplayName + "\u001f" + entry.UninstallCommand, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First());
    }

    private static bool IsForceConnectorProduct(string? displayName) => !string.IsNullOrWhiteSpace(displayName)
        && (displayName.Contains("ForceConnector", StringComparison.OrdinalIgnoreCase)
            || displayName.Contains("Force.com Connector", StringComparison.OrdinalIgnoreCase));

    private IEnumerable<ForceConnectorExcelRegistration> FindEnabledForceConnectorExcelRegistrations()
    {
        foreach (var view in RegistryViews())
        {
            using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);
            using var addins = baseKey.OpenSubKey(ExcelAddinsPath);
            if (addins is null)
            {
                continue;
            }

            foreach (var keyName in addins.GetSubKeyNames().Where(ExcelStartupRegistration.IsLegacyForceConnectorAddinKey))
            {
                using var key = addins.OpenSubKey(keyName);
                var loadBehavior = key?.GetValue("LoadBehavior");
                if (loadBehavior is null || Convert.ToInt32(loadBehavior) == 0)
                {
                    continue;
                }

                yield return new ForceConnectorExcelRegistration(view, ExcelAddinsPath + @"\" + keyName, Convert.ToInt32(loadBehavior));
            }
        }
    }

    private void DisableForceConnectorExcelRegistrations(Action<string>? log)
    {
        foreach (var registration in FindEnabledForceConnectorExcelRegistrations())
        {
            using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, registration.RegistryView);
            using var key = baseKey.OpenSubKey(registration.RegistryPath, writable: true)
                ?? throw new InvalidOperationException($"ForceConnector registration '{registration.RegistryPath}' could not be opened for the current user.");
            key.SetValue("LoadBehavior", 0, RegistryValueKind.DWord);
            Log(log, $"Disabled legacy ForceConnector Excel add-in for the current user: {registration.Description}.");
        }
    }

    private IReadOnlyList<string> RemoveSalesforceRestAddinRegistrations(Action<string>? log)
    {
        var removed = new List<string>();
        foreach (var view in RegistryViews())
        {
            using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);
            foreach (var optionsPath in ExcelOptionsPaths(baseKey))
            {
                using var key = baseKey.OpenSubKey(optionsPath, writable: true);
                if (key is null)
                {
                    continue;
                }

                foreach (var valueName in key.GetValueNames().Where(ExcelStartupRegistration.IsOpenValueName).ToList())
                {
                    if (key.GetValue(valueName) is not string value || !ExcelStartupRegistration.IsSalesforceRestAddinXll(value))
                    {
                        continue;
                    }

                    key.DeleteValue(valueName, throwOnMissingValue: false);
                    var description = $"{view}:{optionsPath}:{valueName}={value}";
                    removed.Add(description);
                    Log(log, $"Removed Excel startup registration {description}");
                }
            }
        }

        return removed;
    }

    private void RegisterExcelAddin(string xllPath, Action<string>? log)
    {
        foreach (var view in RegistryViews())
        {
            using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);
            using var key = baseKey.CreateSubKey(@"SOFTWARE\Microsoft\Office\16.0\Excel\Options", writable: true)
                ?? throw new InvalidOperationException("Excel options registry key could not be created.");
            var slot = NextOpenSlot(key);
            key.SetValue(slot, QuotePath(xllPath), RegistryValueKind.String);
            Log(log, $"Registered Excel startup add-in under {view}:{slot}.");
        }
    }

    private static IEnumerable<string> ExcelOptionsPaths(RegistryKey baseKey)
    {
        using var officeKey = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Office");
        return officeKey?.GetSubKeyNames()
            .Where(name => name.All(character => char.IsDigit(character) || character == '.'))
            .Select(version => @"SOFTWARE\Microsoft\Office\" + version + ExcelOptionsSuffix)
            .ToList()
            ?? [];
    }

    private static string NextOpenSlot(RegistryKey key)
    {
        if (key.GetValue("OPEN") is null)
        {
            return "OPEN";
        }

        for (var index = 1; index < 100; index++)
        {
            var candidate = "OPEN" + index;
            if (key.GetValue(candidate) is null)
            {
                return candidate;
            }
        }

        throw new InvalidOperationException("Excel options registry is full.");
    }

    private static IEnumerable<RegistryView> RegistryViews()
    {
        yield return RegistryView.Registry64;
        yield return RegistryView.Registry32;
    }

    private static string QuotePath(string path) => "\"" + path + "\"";

    private static void Log(Action<string>? log, string message) => log?.Invoke(message);

    private sealed class InstallStatus
    {
        public InstallStatus(bool isUpdateAvailable, string status)
        {
            IsUpdateAvailable = isUpdateAvailable;
            Status = status;
        }

        public bool IsUpdateAvailable { get; }

        public string Status { get; }
    }
}

internal sealed class InstallerState
{
    public string ExcelBitness { get; init; } = "unknown";
    public string AssetName { get; init; } = string.Empty;
    public string InstallDirectory { get; init; } = string.Empty;
    public string? BundledAssetPath { get; init; }
    public bool IsInstalled { get; init; }
    public bool IsUpdateAvailable { get; init; }
    public string Status { get; init; } = string.Empty;
    public string? InstalledVersion { get; init; }
    public string? AvailableVersion { get; init; }
    public string? GitHubCheckError { get; init; }
    public IReadOnlyList<ForceConnectorUninstallEntry> ForceConnectorUninstallEntries { get; init; } = Array.Empty<ForceConnectorUninstallEntry>();
    public IReadOnlyList<ForceConnectorExcelRegistration> ForceConnectorExcelRegistrations { get; init; } = Array.Empty<ForceConnectorExcelRegistration>();
}

internal sealed class ForceConnectorExcelRegistration
{
    public ForceConnectorExcelRegistration(RegistryView registryView, string registryPath, int loadBehavior)
    {
        RegistryView = registryView;
        RegistryPath = registryPath;
        LoadBehavior = loadBehavior;
    }

    public RegistryView RegistryView { get; }
    public string RegistryPath { get; }
    public int LoadBehavior { get; }
    public string Description => $"{RegistryView} current-user Excel registration";
}

internal sealed class ForceConnectorUninstallEntry
{
    public ForceConnectorUninstallEntry(string displayName, string? version, string? publisher, string scope, string registryView, string uninstallCommand)
    {
        DisplayName = displayName;
        Version = version;
        Publisher = publisher;
        Scope = scope;
        RegistryView = registryView;
        UninstallCommand = uninstallCommand;
    }

    public string DisplayName { get; }
    public string? Version { get; }
    public string? Publisher { get; }
    public string Scope { get; }
    public string RegistryView { get; }
    public string UninstallCommand { get; }
    public string Description => $"{DisplayName}{(string.IsNullOrWhiteSpace(Version) ? string.Empty : " " + Version)} ({Scope}, {RegistryView})";
}

internal sealed class ForceConnectorUninstallResult
{
    public ForceConnectorUninstallResult(int exitCode) => ExitCode = exitCode;

    public int ExitCode { get; }

    public bool Succeeded => ExitCode == 0 || ExitCode == 3010;
}

internal sealed class InstallResult
{
    public InstallResult(string assetName, string xllPath, string source, bool downloaded)
    {
        AssetName = assetName;
        XllPath = xllPath;
        Source = source;
        Downloaded = downloaded;
    }

    public string AssetName { get; }
    public string XllPath { get; }
    public string Source { get; }
    public bool Downloaded { get; }
}

internal sealed class RemoveResult
{
    public RemoveResult(string message) => Message = message;
    public string Message { get; }
}

[DataContract]
internal sealed class InstallManifest
{
    [DataMember(Name = "repository")]
    public string Repository { get; set; } = string.Empty;
    [DataMember(Name = "releaseTag")]
    public string ReleaseTag { get; set; } = string.Empty;
    [DataMember(Name = "targetCommitish")]
    public string TargetCommitish { get; set; } = string.Empty;
    [DataMember(Name = "assetName")]
    public string AssetName { get; set; } = string.Empty;
    [DataMember(Name = "installDirectory")]
    public string InstallDirectory { get; set; } = string.Empty;
    [DataMember(Name = "xllPath")]
    public string XllPath { get; set; } = string.Empty;
    [DataMember(Name = "installedAtUtc")]
    public DateTimeOffset InstalledAtUtc { get; set; }
}

[DataContract]
internal sealed class GitHubRelease
{
    [DataMember(Name = "tag_name")]
    public string TagName { get; set; } = string.Empty;
    [DataMember(Name = "target_commitish")]
    public string TargetCommitish { get; set; } = string.Empty;
    [DataMember(Name = "assets")]
    public List<GitHubReleaseAsset> Assets { get; set; } = [];
}

[DataContract]
internal sealed class GitHubReleaseAsset
{
    [DataMember(Name = "name")]
    public string Name { get; set; } = string.Empty;
    [DataMember(Name = "browser_download_url")]
    public string DownloadUrl { get; set; } = string.Empty;
}

internal static class ExceptionFormatter
{
    public static string Format(string operation, Exception exception)
    {
        var builder = new StringBuilder();
        builder.AppendLine(operation);
        builder.AppendLine();
        AppendException(builder, exception, 0);
        return builder.ToString().TrimEnd();
    }

    private static void AppendException(StringBuilder builder, Exception exception, int depth)
    {
        if (depth > 0)
        {
            builder.AppendLine();
            builder.AppendLine($"--- Inner exception {depth} ---");
        }

        builder.AppendLine(exception.GetType().FullName);
        builder.AppendLine(exception.Message);
        if (!string.IsNullOrWhiteSpace(exception.StackTrace))
        {
            builder.AppendLine(exception.StackTrace);
        }

        if (exception.InnerException is not null)
        {
            AppendException(builder, exception.InnerException, depth + 1);
        }
    }
}
