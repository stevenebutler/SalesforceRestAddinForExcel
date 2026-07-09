using System.Net;
using System.Net.Http;
using SalesforceRestAddin.Core;
using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Rest;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Tests.Rest;

public sealed class MetadataCacheTests
{
    [Test]
    public async Task T_META_01_Describe_Miss_Then_Hit_Skips_Second_Http()
    {
        var cacheDir = CreateTempCacheDir();
        try
        {
            var disk = new FileMetadataCache(cacheDir);
            var parsed = new ParsedMetadataCache();
            var (client, handler) = CreateClient(disk, parsed, h =>
            {
                h.Enqueue(HttpStatusCode.OK, LoadFixture("AccountDescribe.json"));
            });

            var first = await client.DescribeAsync("Account");
            var second = await client.DescribeAsync("Account");

            await Assert.That(first.Name).IsEqualTo("Account");
            await Assert.That(second.Name).IsEqualTo("Account");
            await Assert.That(handler.Requests.Count).IsEqualTo(1);
        }
        finally
        {
            TryDeleteDirectory(cacheDir);
        }
    }

    [Test]
    public async Task T_META_02_ClearInstance_Forces_Redownload()
    {
        var cacheDir = CreateTempCacheDir();
        try
        {
            var disk = new FileMetadataCache(cacheDir);
            var parsed = new ParsedMetadataCache();
            var (client, handler) = CreateClient(disk, parsed, h =>
            {
                h.Enqueue(HttpStatusCode.OK, LoadFixture("AccountDescribe.json"));
                h.Enqueue(HttpStatusCode.OK, LoadFixture("AccountDescribe.json"));
            });

            await client.DescribeAsync("Account");
            parsed.ClearInstance("https://example.my.salesforce.com");
            disk.ClearInstance("https://example.my.salesforce.com");
            await client.DescribeAsync("Account");

            await Assert.That(handler.Requests.Count).IsEqualTo(2);
        }
        finally
        {
            TryDeleteDirectory(cacheDir);
        }
    }

    [Test]
    public async Task T_META_03_Separate_Entries_Per_Api_Version_And_Instance()
    {
        var cacheDir = CreateTempCacheDir();
        try
        {
            var cache = new FileMetadataCache(cacheDir);
            const string jsonA = "{\"sobjects\":[]}";
            const string jsonB = "{\"sobjects\":[{\"name\":\"Account\",\"label\":\"Account\",\"queryable\":true,\"custom\":false}]}";

            cache.SetObjectList("https://a.my.salesforce.com", "66.0", jsonA);
            cache.SetObjectList("https://a.my.salesforce.com", "67.0", jsonB);
            cache.SetObjectList("https://b.my.salesforce.com", "66.0", jsonB);

            await Assert.That(cache.TryGetObjectList("https://a.my.salesforce.com", "66.0", out var a66)).IsTrue();
            await Assert.That(a66).IsEqualTo(jsonA);
            await Assert.That(cache.TryGetObjectList("https://a.my.salesforce.com", "67.0", out var a67)).IsTrue();
            await Assert.That(a67).IsEqualTo(jsonB);
            await Assert.That(cache.TryGetObjectList("https://b.my.salesforce.com", "66.0", out var b66)).IsTrue();
            await Assert.That(b66).IsEqualTo(jsonB);
            await Assert.That(cache.TryGetObjectList("https://missing.my.salesforce.com", "66.0", out _)).IsFalse();

            cache.SetDescribe("https://a.my.salesforce.com", "66.0", "Account", "{\"name\":\"Account\"}");
            await Assert.That(cache.TryGetDescribe("https://a.my.salesforce.com", "66.0", "Account", out var describe)).IsTrue();
            await Assert.That(describe).Contains("Account");
            await Assert.That(cache.TryGetDescribe("https://a.my.salesforce.com", "67.0", "Account", out _)).IsFalse();
        }
        finally
        {
            TryDeleteDirectory(cacheDir);
        }
    }

    [Test]
    public async Task T_META_04_Path_Policy_Rejects_Forbidden_Roots_Without_Io()
    {
        // Pure string policy — never constructs FileMetadataCache against a real profile path.
        var forbidden = new[]
        {
            Path.Combine(Path.DirectorySeparatorChar.ToString(), "fake-home"),
            Path.Combine(Path.DirectorySeparatorChar.ToString(), "fake-local-app-data"),
        };

        await Assert.That(FileMetadataCache.IsSafeCacheRoot(forbidden[0], forbidden)).IsFalse();
        await Assert.That(FileMetadataCache.IsSafeCacheRoot(forbidden[1], forbidden)).IsFalse();
        await Assert.That(
            FileMetadataCache.IsSafeCacheRoot(
                Path.Combine(forbidden[0], "SalesforceRestAddin", "metadata-cache"),
                forbidden)).IsTrue();

        // Drive / filesystem root is always unsafe regardless of forbidden list.
        var driveRoot = Path.GetPathRoot(Path.GetFullPath(Path.DirectorySeparatorChar.ToString()));
        await Assert.That(driveRoot).IsNotNull();
        await Assert.That(FileMetadataCache.IsSafeCacheRoot(driveRoot!, Array.Empty<string>())).IsFalse();
    }

    [Test]
    public async Task T_META_05_Default_Forbidden_List_Includes_Special_Folders()
    {
        // Equality checks only — never pass these paths to ClearInstance().
        var forbidden = FileMetadataCache.GetDefaultForbiddenRoots();
        var profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

        if (!string.IsNullOrWhiteSpace(profile))
        {
            await Assert.That(
                forbidden.Any(f => string.Equals(
                    Path.GetFullPath(f).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                    Path.GetFullPath(profile).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                    StringComparison.OrdinalIgnoreCase))).IsTrue();
            await Assert.That(FileMetadataCache.IsSafeCacheRoot(profile)).IsFalse();
        }

        if (!string.IsNullOrWhiteSpace(localAppData))
        {
            await Assert.That(FileMetadataCache.IsSafeCacheRoot(localAppData)).IsFalse();
            // Production cache path is a subdirectory — allowed.
            await Assert.That(FileMetadataCache.IsSafeCacheRoot(
                Path.Combine(localAppData, "SalesforceRestAddin", "metadata-cache"))).IsTrue();
        }
    }

    [Test]
    public async Task T_META_06_ClearInstance_Without_Marker_Leaves_Contents()
    {
        var cacheDir = CreateTempCacheDir();
        try
        {
            Directory.CreateDirectory(cacheDir);
            var decoy = Path.Combine(cacheDir, "decoy.txt");
            File.WriteAllText(decoy, "must-survive");

            var cache = new FileMetadataCache(cacheDir);
            cache.ClearInstance("https://a.my.salesforce.com");

            await Assert.That(File.Exists(decoy)).IsTrue();
            await Assert.That(File.ReadAllText(decoy)).IsEqualTo("must-survive");
            await Assert.That(File.Exists(Path.Combine(cacheDir, FileMetadataCache.MarkerFileName))).IsFalse();
        }
        finally
        {
            TryDeleteDirectory(cacheDir);
        }
    }

    [Test]
    public async Task T_META_07_ClearInstance_Removes_Only_That_Instance()
    {
        var cacheDir = CreateTempCacheDir();
        try
        {
            var cache = new FileMetadataCache(cacheDir);
            const string prod = "https://kjr.my.salesforce.com";
            const string sandbox = "https://kjr--kjr2026.sandbox.my.salesforce.com";
            cache.SetObjectList(prod, "66.0", "{\"sobjects\":[]}");
            cache.SetObjectList(sandbox, "66.0", "{\"sobjects\":[{\"name\":\"Account\"}]}");

            var marker = Path.Combine(cacheDir, FileMetadataCache.MarkerFileName);
            await Assert.That(File.Exists(marker)).IsTrue();

            cache.ClearInstance(sandbox);

            await Assert.That(Directory.Exists(cacheDir)).IsTrue();
            await Assert.That(File.Exists(marker)).IsTrue();
            await Assert.That(cache.TryGetObjectList(sandbox, "66.0", out _)).IsFalse();
            await Assert.That(cache.TryGetObjectList(prod, "66.0", out var prodJson)).IsTrue();
            await Assert.That(prodJson).IsEqualTo("{\"sobjects\":[]}");
        }
        finally
        {
            TryDeleteDirectory(cacheDir);
        }
    }

    [Test]
    public async Task T_META_08_Constructor_Rejects_Forbidden_Root_Via_Injected_Policy()
    {
        var parent = CreateTempCacheDir();
        var unsafeRoot = Path.Combine(parent, "pretend-home");
        try
        {
            Directory.CreateDirectory(unsafeRoot);
            // Simulate a forbidden root without using the real user profile.
            await Assert.That(FileMetadataCache.IsSafeCacheRoot(unsafeRoot, new[] { unsafeRoot })).IsFalse();

            // Construction uses the default forbidden list; a unique temp subdir is allowed.
            var allowed = Path.Combine(parent, "metadata-cache");
            var cache = new FileMetadataCache(allowed);
            cache.SetObjectList("https://a.my.salesforce.com", "66.0", "{}");
            await Assert.That(File.Exists(Path.Combine(allowed, FileMetadataCache.MarkerFileName))).IsTrue();
        }
        finally
        {
            TryDeleteDirectory(parent);
        }
    }

    [Test]
    public async Task T_META_09_Blank_Instance_Url_Does_Not_Read_Or_Write_Cache()
    {
        var cacheDir = CreateTempCacheDir();
        try
        {
            var cache = new FileMetadataCache(cacheDir);
            cache.SetObjectList("", "66.0", "{\"sobjects\":[]}");
            cache.SetObjectList("   ", "66.0", "{\"sobjects\":[]}");
            cache.SetDescribe(null!, "66.0", "Account", "{}");

            await Assert.That(cache.TryGetObjectList("", "66.0", out _)).IsFalse();
            await Assert.That(cache.TryGetDescribe("not-a-url", "66.0", "Account", out _)).IsFalse();
            await Assert.That(Directory.Exists(cacheDir) && Directory.EnumerateFileSystemEntries(cacheDir).Any()).IsFalse();
        }
        finally
        {
            TryDeleteDirectory(cacheDir);
        }
    }

    [Test]
    public async Task T_META_10_Sandbox_And_Production_Hosts_Do_Not_Share_Cache()
    {
        var cacheDir = CreateTempCacheDir();
        try
        {
            var cache = new FileMetadataCache(cacheDir);
            const string prod = "https://kjr.my.salesforce.com";
            const string sandbox = "https://kjr--kjr2026.sandbox.my.salesforce.com";
            const string prodJson = "{\"sobjects\":[{\"name\":\"Account\",\"label\":\"Prod\",\"queryable\":true,\"custom\":false}]}";
            const string sandboxJson = "{\"sobjects\":[{\"name\":\"Account\",\"label\":\"Sandbox\",\"queryable\":true,\"custom\":false}]}";

            cache.SetObjectList(prod, "67.0", prodJson);
            cache.SetObjectList(sandbox, "67.0", sandboxJson);

            await Assert.That(cache.TryGetObjectList(prod, "67.0", out var fromProd)).IsTrue();
            await Assert.That(fromProd).IsEqualTo(prodJson);
            await Assert.That(cache.TryGetObjectList(sandbox, "67.0", out var fromSandbox)).IsTrue();
            await Assert.That(fromSandbox).IsEqualTo(sandboxJson);

            // Trailing slash / path noise must not create a second bucket for the same host.
            await Assert.That(cache.TryGetObjectList(prod + "/", "67.0", out var fromProdSlash)).IsTrue();
            await Assert.That(fromProdSlash).IsEqualTo(prodJson);

            await Assert.That(FileMetadataCache.TryNormalizeInstanceKey(prod, out var prodKey)).IsTrue();
            await Assert.That(FileMetadataCache.TryNormalizeInstanceKey(sandbox, out var sandboxKey)).IsTrue();
            await Assert.That(prodKey).IsNotEqualTo(sandboxKey);
        }
        finally
        {
            TryDeleteDirectory(cacheDir);
        }
    }

    [Test]
    public async Task T_META_11_Typed_L1_Serves_Second_Describe_Without_Http()
    {
        var cacheDir = CreateTempCacheDir();
        try
        {
            var disk = new FileMetadataCache(cacheDir);
            var parsed = new ParsedMetadataCache();
            var (client, handler) = CreateClient(disk, parsed, h =>
            {
                h.Enqueue(HttpStatusCode.OK, LoadFixture("AccountDescribe.json"));
            });

            var first = await client.DescribeAsync("Account");
            await Assert.That(handler.Requests.Count).IsEqualTo(1);
            await Assert.That(
                parsed.TryGetDescribe("https://example.my.salesforce.com", "66.0", "Account", out var fromL1)).IsTrue();
            await Assert.That(fromL1!.Name).IsEqualTo(first.Name);
            await Assert.That(
                parsed.TryGetFieldCatalog("https://example.my.salesforce.com", "66.0", "Account", out var catalog)).IsTrue();
            await Assert.That(catalog).IsNotNull();

            var second = await client.DescribeAsync("Account");
            await Assert.That(second.Name).IsEqualTo("Account");
            await Assert.That(handler.Requests.Count).IsEqualTo(1);
            await Assert.That(ReferenceEquals(first, second)).IsTrue();
        }
        finally
        {
            TryDeleteDirectory(cacheDir);
        }
    }

    [Test]
    public async Task T_META_12_ClearInstance_Leaves_Other_Org_Intact_In_L1_And_L2()
    {
        var cacheDir = CreateTempCacheDir();
        try
        {
            var disk = new FileMetadataCache(cacheDir);
            var parsed = new ParsedMetadataCache();
            const string prod = "https://kjr.my.salesforce.com";
            const string sandbox = "https://kjr--kjr2026.sandbox.my.salesforce.com";

            disk.SetObjectList(prod, "66.0", "{\"sobjects\":[]}");
            disk.SetObjectList(sandbox, "66.0", "{\"sobjects\":[{\"name\":\"Account\",\"label\":\"A\",\"queryable\":true,\"custom\":false}]}");
            parsed.SetObjectList(prod, "66.0", Array.Empty<SObjectSummary>());
            parsed.SetObjectList(
                sandbox,
                "66.0",
                new[]
                {
                    new SObjectSummary { Name = "Account", Label = "A", Queryable = true, Custom = false },
                });

            parsed.ClearInstance(sandbox);
            disk.ClearInstance(sandbox);

            await Assert.That(parsed.TryGetObjectList(sandbox, "66.0", out _)).IsFalse();
            await Assert.That(disk.TryGetObjectList(sandbox, "66.0", out _)).IsFalse();
            await Assert.That(parsed.TryGetObjectList(prod, "66.0", out _)).IsTrue();
            await Assert.That(disk.TryGetObjectList(prod, "66.0", out var prodJson)).IsTrue();
            await Assert.That(prodJson).IsEqualTo("{\"sobjects\":[]}");
        }
        finally
        {
            TryDeleteDirectory(cacheDir);
        }
    }

    private static (SalesforceDataClient client, SequentialMockHttpHandler handler) CreateClient(
        IMetadataCache disk,
        ParsedMetadataCache parsed,
        Action<SequentialMockHttpHandler>? configure = null)
    {
        var handler = new SequentialMockHttpHandler();
        configure?.Invoke(handler);
        var http = new HttpClient(handler) { BaseAddress = new Uri("https://example.my.salesforce.com") };
        var session = new SessionContext
        {
            AccessToken = "token",
            InstanceUrl = "https://example.my.salesforce.com",
            Id = "https://login.salesforce.com/id/00D/005",
            IssuedAtUnixMs = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(),
        };
        session.SetApiVersion("66.0");
        var creds = new InMemorySessionCredentialStore();
        var prefs = new InMemoryUserLoginPreferencesStore();
        var oauth = SalesforceOAuthOptions.Default;
        var authenticator = new SessionAuthenticator(creds, new SalesforceOAuthClient(http), oauth);
        var orchestrator = new SessionLoginOrchestrator(session, authenticator, new FakeInteractiveLoginHandler());
        var authClient = new SalesforceAuthenticatedClient(
            http,
            session,
            prefs,
            new SessionAccessTokenRefresher(authenticator),
            orchestrator);
        return (new SalesforceDataClient(authClient, session, disk, parsed), handler);
    }

    private static string LoadFixture(string name)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Fixtures", name);
        return File.ReadAllText(path);
    }

    private static string CreateTempCacheDir() =>
        Path.Combine(Path.GetTempPath(), "SalesforceRestAddin.Tests.MetadataCache." + Guid.NewGuid().ToString("N"));

    private static void TryDeleteDirectory(string path)
    {
        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, recursive: true);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}
