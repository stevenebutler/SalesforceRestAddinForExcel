using System;
using System.IO;
using SalesforceRestAddin.Core;
using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Tests;

public sealed class JsonUserLoginPreferencesStoreTests
{
    [Test]
    public async Task SaveAndLoad_RoundTripsApiVersionAndCachedSupportedVersions()
    {
        var path = Path.Combine(Path.GetTempPath(), $"forceconnector-prefs-{Guid.NewGuid():N}.json");
        var store = new JsonUserLoginPreferencesStore(path);

        store.Save(new UserLoginPreferences
        {
            Tenant = "kjr",
            Environment = SalesforceEnvironment.Sandbox,
            SandboxId = "dev",
            ApiVersion = "66.0",
            CachedSupportedApiVersions = ["67.0", "66.0", "65.0"],
            ShowLoginOptionsOnNextUse = true,
        });

        var loaded = store.TryLoadValid();

        await Assert.That(loaded).IsNotNull();
        await Assert.That(loaded!.Tenant).IsEqualTo("kjr");
        await Assert.That(loaded.Environment).IsEqualTo(SalesforceEnvironment.Sandbox);
        await Assert.That(loaded.SandboxId).IsEqualTo("dev");
        await Assert.That(loaded.ApiVersion).IsEqualTo("66.0");
        await Assert.That(loaded.CachedSupportedApiVersions).IsEquivalentTo(new[] { "67.0", "66.0", "65.0" });
        await Assert.That(loaded.ShowLoginOptionsOnNextUse).IsTrue();

        File.Delete(path);
    }

    [Test]
    public async Task Load_InvalidJson_ReturnsNull()
    {
        var path = Path.Combine(Path.GetTempPath(), $"forceconnector-prefs-{Guid.NewGuid():N}.json");
        File.WriteAllText(path, "not-json");

        await Assert.That(new JsonUserLoginPreferencesStore(path).TryLoadValid()).IsNull();

        File.Delete(path);
    }

    [Test]
    public async Task Load_MissingEnvironment_ReturnsNull()
    {
        var path = Path.Combine(Path.GetTempPath(), $"forceconnector-prefs-{Guid.NewGuid():N}.json");
        File.WriteAllText(path, """{"tenant":"kjr"}""");

        await Assert.That(new JsonUserLoginPreferencesStore(path).TryLoadValid()).IsNull();

        File.Delete(path);
    }

    [Test]
    public async Task Load_InvalidEnvironment_ReturnsNull()
    {
        var path = Path.Combine(Path.GetTempPath(), $"forceconnector-prefs-{Guid.NewGuid():N}.json");
        File.WriteAllText(path, """{"environment":"Staging"}""");

        await Assert.That(new JsonUserLoginPreferencesStore(path).TryLoadValid()).IsNull();

        File.Delete(path);
    }

    [Test]
    public async Task Load_ValidProduction_ReturnsPrefs()
    {
        var path = Path.Combine(Path.GetTempPath(), $"forceconnector-prefs-{Guid.NewGuid():N}.json");
        File.WriteAllText(path, """{"tenant":"kjr","environment":"Production"}""");

        var loaded = new JsonUserLoginPreferencesStore(path).TryLoadValid();

        await Assert.That(loaded).IsNotNull();
        await Assert.That(loaded!.Environment).IsEqualTo(SalesforceEnvironment.Production);
        await Assert.That(loaded.ApiVersion).IsNull();
        await Assert.That(loaded.CachedSupportedApiVersions).IsEmpty();

        File.Delete(path);
    }

    [Test]
    public async Task Load_MissingApiVersionFields_DefaultsToLatestSemantics()
    {
        var path = Path.Combine(Path.GetTempPath(), $"forceconnector-prefs-{Guid.NewGuid():N}.json");
        File.WriteAllText(path, """{"tenant":"kjr","environment":"Production"}""");

        var loaded = new JsonUserLoginPreferencesStore(path).TryLoadValid();

        await Assert.That(loaded!.ApiVersion).IsNull();
        await Assert.That(loaded.CachedSupportedApiVersions).IsEmpty();

        File.Delete(path);
    }

    [Test]
    public async Task SaveAndLoad_RoundTripsJwtAccessToken()
    {
        var path = Path.Combine(Path.GetTempPath(), $"forceconnector-prefs-{Guid.NewGuid():N}.json");
        var store = new JsonUserLoginPreferencesStore(path);

        store.Save(new UserLoginPreferences
        {
            Tenant = "kjr",
            Environment = SalesforceEnvironment.Production,
            JwtAccessToken = " eyJ.debug ",
        });

        var json = File.ReadAllText(path);
        await Assert.That(json).Contains("jwtAccessToken");

        var loaded = store.TryLoadValid();
        await Assert.That(loaded!.JwtAccessToken).IsEqualTo("eyJ.debug");
        await Assert.That(loaded.HasJwtAccessToken).IsTrue();

        File.Delete(path);
    }

    [Test]
    public async Task Save_OmitsJwtAccessToken_WhenUnset()
    {
        var path = Path.Combine(Path.GetTempPath(), $"forceconnector-prefs-{Guid.NewGuid():N}.json");
        var store = new JsonUserLoginPreferencesStore(path);

        store.Save(new UserLoginPreferences
        {
            Tenant = "kjr",
            Environment = SalesforceEnvironment.Production,
        });

        var json = File.ReadAllText(path);
        await Assert.That(json).DoesNotContain("jwtAccessToken");

        File.Delete(path);
    }

    [Test]
    public async Task Load_WhitespaceJwtAccessToken_TreatedAsUnset()
    {
        var path = Path.Combine(Path.GetTempPath(), $"forceconnector-prefs-{Guid.NewGuid():N}.json");
        File.WriteAllText(path, """{"tenant":"kjr","environment":"Production","jwtAccessToken":"   "}""");

        var loaded = new JsonUserLoginPreferencesStore(path).TryLoadValid();

        await Assert.That(loaded!.JwtAccessToken).IsNull();
        await Assert.That(loaded.HasJwtAccessToken).IsFalse();

        File.Delete(path);
    }
}
