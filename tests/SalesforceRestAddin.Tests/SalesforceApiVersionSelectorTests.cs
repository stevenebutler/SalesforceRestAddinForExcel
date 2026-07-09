using SalesforceRestAddin.Core;
using SalesforceRestAddin.Core.Rest;

namespace SalesforceRestAddin.Tests;

public sealed class SalesforceApiVersionSelectorTests
{
    private const string VersionsJson = """
        [
          {"label":"Summer '25","url":"/services/data/v65.0/","version":"65.0"},
          {"label":"Winter '26","url":"/services/data/v66.0/","version":"66.0"},
          {"label":"Spring '26","url":"/services/data/v67.0/","version":"67.0"}
        ]
        """;

    [Test]
    public async Task Resolve_WithoutPreference_UsesLatestFromOrg()
    {
        var supported = SalesforceApiVersions.ParseSupportedVersionsFromJson(VersionsJson);
        var resolution = SalesforceApiVersionSelector.Resolve(null, supported);

        await Assert.That(resolution.SelectedVersion).IsEqualTo("67.0");
        await Assert.That(resolution.UsedLatestFromOrg).IsTrue();
        await Assert.That(resolution.PreferredVersionWasInvalid).IsFalse();
    }

    [Test]
    public async Task Resolve_WithSupportedPreference_UsesPreference()
    {
        var supported = SalesforceApiVersions.ParseSupportedVersionsFromJson(VersionsJson);
        var resolution = SalesforceApiVersionSelector.Resolve("66.0", supported);

        await Assert.That(resolution.SelectedVersion).IsEqualTo("66.0");
        await Assert.That(resolution.UsedLatestFromOrg).IsFalse();
        await Assert.That(resolution.PreferredVersionWasInvalid).IsFalse();
    }

    [Test]
    public async Task Resolve_WithUnsupportedPreference_FallsBackToLatest()
    {
        var supported = SalesforceApiVersions.ParseSupportedVersionsFromJson(VersionsJson);
        var resolution = SalesforceApiVersionSelector.Resolve("51.0", supported);

        await Assert.That(resolution.SelectedVersion).IsEqualTo("67.0");
        await Assert.That(resolution.PreferredVersionWasInvalid).IsTrue();
        await Assert.That(resolution.UsedLatestFromOrg).IsTrue();
    }

    [Test]
    public async Task Resolve_WithoutOrgVersions_UsesPreferredVersion()
    {
        var resolution = SalesforceApiVersionSelector.Resolve("64.0", Array.Empty<string>());

        await Assert.That(resolution.SelectedVersion).IsEqualTo("64.0");
    }

    [Test]
    public async Task Resolve_WithoutOrgVersionsOrPreference_Throws()
    {
        await Assert.That(() => SalesforceApiVersionSelector.Resolve(null, Array.Empty<string>()))
            .Throws<InvalidOperationException>();
    }

    [Test]
    public async Task ParseSupportedVersionsFromJson_ReturnsDescendingOrder()
    {
        var supported = SalesforceApiVersions.ParseSupportedVersionsFromJson(VersionsJson);

        await Assert.That(supported).IsEquivalentTo(new[] { "67.0", "66.0", "65.0" });
    }

    [Test]
    public async Task Normalize_FormatsVersionLiterals()
    {
        await Assert.That(SalesforceApiVersions.Normalize("66")).IsEqualTo("66.0");
        await Assert.That(SalesforceApiVersions.AreEqual("66", "66.0")).IsTrue();
    }
}
