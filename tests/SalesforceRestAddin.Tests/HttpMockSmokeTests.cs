using SalesforceRestAddin.Core;
using TUnit.Mocks;

namespace SalesforceRestAddin.Tests;

/// <summary>
/// Proves the TUnit HTTP mock harness works for upcoming Salesforce REST port slices.
/// </summary>
public class HttpMockSmokeTests
{
    [Test]
    public async Task MockHttpClient_Returns_Configured_Json()
    {
        using var client = Mock.HttpClient("https://instance.my.salesforce.com");
        client.Handler.OnGet("/services/data/v59.0/sobjects/")
            .RespondWithJson("""{"sobjects":[],"encoding":"UTF-8","maxBatchSize":200}""");

        var response = await client.GetAsync("/services/data/v59.0/sobjects/");
        var body = await response.Content.ReadAsStringAsync();

        await Assert.That(response.IsSuccessStatusCode).IsTrue();
        await Assert.That(body).Contains("sobjects");
    }
}
