using SalesforceRestAddin.Core.Rest;

namespace SalesforceRestAddin.Tests;

public sealed class SalesforceApiVersionsTests
{
    [Test]
    public async Task ParseLatestFromVersionsJson_Picks_highest_version()
    {
        const string json = """
            [
              {"label":"Summer '25","url":"/services/data/v65.0/","version":"65.0"},
              {"label":"Winter '26","url":"/services/data/v66.0/","version":"66.0"},
              {"label":"Spring '26","url":"/services/data/v67.0/","version":"67.0"}
            ]
            """;

        await Assert.That(SalesforceApiVersions.ParseLatestFromVersionsJson(json)).IsEqualTo("67.0");
    }

    [Test]
    public async Task ResolveLatestAsync_Uses_org_version_list()
    {
        using var client = Mock.HttpClient("https://kjr.my.salesforce.com");
        client.Handler.OnGet("/services/data/")
            .RespondWithJson("""
                [
                  {"label":"Winter '26","url":"/services/data/v66.0/","version":"66.0"},
                  {"label":"Spring '26","url":"/services/data/v67.0/","version":"67.0"}
                ]
                """);

        var resolver = new SalesforceApiVersionResolver(client);
        var version = await resolver.ResolveLatestAsync(
            "https://kjr.my.salesforce.com",
            "00Dtoken");

        await Assert.That(version).IsEqualTo("67.0");
    }

    [Test]
    public async Task ResolveForSessionAsync_WithoutPreferenceAndNoOrgList_Throws()
    {
        using var client = Mock.HttpClient("https://kjr.my.salesforce.com");
        client.Handler.OnGet("/services/data/")
            .RespondWithJson("""{"error":"invalid_token"}""", System.Net.HttpStatusCode.Unauthorized);

        var resolver = new SalesforceApiVersionResolver(client);

        await Assert.That(async () => await resolver.ResolveForSessionAsync(
                "https://kjr.my.salesforce.com",
                "bad-token"))
            .Throws<InvalidOperationException>();
    }
}
