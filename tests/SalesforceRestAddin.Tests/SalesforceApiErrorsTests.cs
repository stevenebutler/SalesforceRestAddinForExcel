using System.Net;
using System.Net.Http;
using SalesforceRestAddin.Core.Rest;

namespace SalesforceRestAddin.Tests;

public sealed class SalesforceApiErrorsTests
{
    [Test]
    public async Task IsSessionAuthFailure_TreatsMissingOAuthToken403_AsAuthFailure()
    {
        using var response = new HttpResponseMessage(HttpStatusCode.Forbidden)
        {
            Content = new StringContent(SalesforceMockResponses.MissingOAuthToken403),
        };

        await Assert.That(SalesforceApiErrors.IsSessionAuthFailure(response)).IsTrue();
        await Assert.That(SalesforceApiErrors.IsSessionAuthFailure(SalesforceMockResponses.MissingOAuthToken403)).IsTrue();
    }

    [Test]
    public async Task IsSessionAuthFailure_Unrelated403_IsNotAuthFailure()
    {
        using var response = new HttpResponseMessage(HttpStatusCode.Forbidden)
        {
            Content = new StringContent("""[{"message":"Insufficient access","errorCode":"INSUFFICIENT_ACCESS"}]"""),
        };

        await Assert.That(SalesforceApiErrors.IsSessionAuthFailure(response)).IsFalse();
    }
}
