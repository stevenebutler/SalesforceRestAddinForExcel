using System.IO;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Tests;

public sealed class SalesforceRestAddinDataPathsTests
{
    [Test]
    public async Task Data_paths_are_absolute_and_under_force_connector()
    {
        await Assert.That(Path.IsPathRooted(SalesforceRestAddinDataPaths.DataDirectory)).IsTrue();
        await Assert.That(Path.IsPathRooted(SalesforceRestAddinDataPaths.PreferencesFile)).IsTrue();
        await Assert.That(Path.IsPathRooted(SalesforceRestAddinDataPaths.WebView2UserDataFolder)).IsTrue();

        await Assert.That(SalesforceRestAddinDataPaths.DataDirectory).EndsWith("SalesforceRestAddin");
        await Assert.That(SalesforceRestAddinDataPaths.PreferencesFile).EndsWith("login-preferences.json");
        await Assert.That(SalesforceRestAddinDataPaths.ConnectorOptionsFile).EndsWith("connector-options.json");
        await Assert.That(SalesforceRestAddinDataPaths.WebView2UserDataFolder).EndsWith(
            Path.Combine("SalesforceRestAddin", "WebView2"));
    }
}
