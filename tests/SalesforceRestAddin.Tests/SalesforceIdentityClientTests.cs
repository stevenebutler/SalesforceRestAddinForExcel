using SalesforceRestAddin.Core.Rest;
using TUnit.Mocks;

namespace SalesforceRestAddin.Tests;

public sealed class SalesforceIdentityClientTests
{
    [Test]
    public async Task GetIdentityAsync_RequestsVersionLatestAndFormatsResponse()
    {
        using var client = Mock.HttpClient("https://login.salesforce.com");
        client.Handler.OnGet("/id/00Dxx/005yy?version=latest")
            .RespondWithJson(
                """
                {
                  "id": "https://login.salesforce.com/id/00Dxx/005yy",
                  "display_name": "Test User",
                  "username": "user@example.com",
                  "organization_id": "00Dxx",
                  "urls": {
                    "rest": "https://kjr.my.salesforce.com/services/data/v67.0/",
                    "partner": "https://kjr.my.salesforce.com/services/Soap/u/67.0"
                  }
                }
                """);

        var identityClient = new SalesforceIdentityClient(client);
        var identity = await identityClient.GetIdentityAsync(
            "https://login.salesforce.com/id/00Dxx/005yy",
            accessToken: "token");

        await Assert.That(client.Handler.Requests[0].RequestUri!.Query).IsEqualTo("?version=latest");
        await Assert.That(identity.FormatSummary()).Contains("display_name: Test User");
        await Assert.That(identity.FormatSummary()).Contains("rest: https://kjr.my.salesforce.com/services/data/v67.0/");
        await Assert.That(identity.TryGetDisplayName()).IsEqualTo("Test User");
    }

    [Test]
    public async Task TryGetDisplayName_ReturnsNull_WhenMissing()
    {
        var identity = SalesforceIdentityResponse.ParseJson("""{"id":"https://login.salesforce.com/id/00D/005"}""");

        await Assert.That(identity.TryGetDisplayName()).IsNull();
    }
}
