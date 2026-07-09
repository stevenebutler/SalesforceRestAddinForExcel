namespace SalesforceRestAddin.Tests;

public static class SalesforceMockResponses
{
    public const string InvalidSessionId401 =
        """[{"message":"Session expired or invalid","errorCode":"INVALID_SESSION_ID"}]""";

    public const string InvalidAuthHeader401 =
        """[{"message":"INVALID_AUTH_HEADER","errorCode":"INVALID_AUTH_HEADER"}]""";

    public const string InvalidGrant400 =
        """{"error":"invalid_grant","error_description":"expired access/refresh token"}""";

    public static string RefreshSuccess =>
        $$"""
        {
          "access_token": "refreshed-access",
          "instance_url": "https://kjr.my.salesforce.com",
          "id": "https://login.salesforce.com/id/00Dxx/005yy",
          "issued_at": "{{DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()}}"
        }
        """;

    public const string MissingOAuthToken403 = "Missing_OAuth_Token";

    public const string IdentitySuccess =
        """
        {
          "id": "https://login.salesforce.com/id/00Dxx/005yy",
          "display_name": "Test User",
          "username": "user@example.com"
        }
        """;
}
