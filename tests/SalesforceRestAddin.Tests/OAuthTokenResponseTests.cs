using SalesforceRestAddin.Core;
using SalesforceRestAddin.Core.OAuth;

namespace SalesforceRestAddin.Tests;

public sealed class OAuthTokenResponseTests
{
    [Test]
    public async Task ParseJson_CapturesNonSecretMetadataAndAdditionalProperties()
    {
        const string json = """
            {
              "access_token": "secret-access",
              "refresh_token": "secret-refresh",
              "instance_url": "https://kjr.my.salesforce.com",
              "id": "https://login.salesforce.com/id/00Dxx/005yy",
              "token_type": "Bearer",
              "issued_at": "1710000000000",
              "scope": "refresh_token api offline_access",
              "signature": "abc123",
              "sfdc_community_url": "https://community.example.com",
              "sfdc_community_id": "0DBxx",
              "custom_field": "extra"
            }
            """;

        var token = OAuthTokenResponse.ParseJson(json);

        await Assert.That(token.TokenType).IsEqualTo("Bearer");
        await Assert.That(token.IssuedAtUnixMs).IsEqualTo(1_710_000_000_000);
        await Assert.That(token.Scope).IsEqualTo("refresh_token api offline_access");
        await Assert.That(token.Signature).IsEqualTo("abc123");
        await Assert.That(token.SfdcCommunityUrl).IsEqualTo("https://community.example.com");
        await Assert.That(token.SfdcCommunityId).IsEqualTo("0DBxx");
        await Assert.That(token.AdditionalProperties["custom_field"]).IsEqualTo("extra");
    }

    [Test]
    public async Task FormatMetadataSummary_OmitsSecretsAndExplainsLifetime()
    {
        var token = OAuthTokenResponse.ParseJson(
            """
            {
              "access_token": "secret-access",
              "refresh_token": "secret-refresh",
              "instance_url": "https://kjr.my.salesforce.com",
              "id": "https://login.salesforce.com/id/00Dxx/005yy",
              "token_type": "Bearer",
              "issued_at": "1710000000000",
              "scope": "refresh_token api"
            }
            """);

        var summary = token.FormatMetadataSummary(
            DateTimeOffset.FromUnixTimeMilliseconds(1_710_003_600_000));

        await Assert.That(summary).Contains("access_token: present");
        await Assert.That(summary).Contains("refresh_token: present");
        await Assert.That(summary).DoesNotContain("secret-access");
        await Assert.That(summary).DoesNotContain("secret-refresh");
        await Assert.That(summary).Contains("scope: refresh_token api");
        await Assert.That(summary).Contains("do not include expires_in");
        await Assert.That(summary).Contains("Access token age");
    }

    [Test]
    public async Task ApplyTo_PreservesExistingRefreshTokenWhenOmitted()
    {
        var session = new SessionContext { RefreshToken = "existing-refresh" };
        var token = OAuthTokenResponse.ParseJson(
            """
            {
              "access_token": "new-access",
              "instance_url": "https://kjr.my.salesforce.com",
              "id": "https://login.salesforce.com/id/00Dxx/005yy"
            }
            """);

        token.ApplyTo(session);

        await Assert.That(session.AccessToken).IsEqualTo("new-access");
        await Assert.That(session.RefreshToken).IsEqualTo("existing-refresh");
    }
}
