namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Outcome of attempting to persist a refresh token for a login host.
/// </summary>
public sealed class SessionPersistResult
{
    public bool Saved { get; init; }

    public bool RefreshTokenPresentInResponse { get; init; }

    public string? ErrorMessage { get; init; }

    public static SessionPersistResult NoRefreshTokenInResponse() => new()
    {
        RefreshTokenPresentInResponse = false,
        ErrorMessage = "Salesforce did not return a refresh token. Saved sign-in will not be available until offline access is granted for this connected app.",
    };

    public static SessionPersistResult SavedSuccessfully() => new()
    {
        Saved = true,
        RefreshTokenPresentInResponse = true,
    };

    public static SessionPersistResult Failed(string message) => new()
    {
        RefreshTokenPresentInResponse = true,
        ErrorMessage = message,
    };
}
