namespace SalesforceRestAddin.Core;

/// <summary>
/// Salesforce session state shared by the Excel-DNA host and Core library.
/// Replaces the static fields previously held on VSTO <c>ThisAddIn</c>.
/// </summary>
public sealed class SessionContext
{
    public static SessionContext Current { get; } = new();

    private string? _apiVersion;

    public string? AccessToken { get; set; }

    public string? RefreshToken { get; set; }

    public string? InstanceUrl { get; set; }

    public string? Id { get; set; }

    /// <summary>
    /// Salesforce identity <c>display_name</c> for the ribbon label; session-only (not persisted).
    /// </summary>
    public string? DisplayName { get; set; }

    /// <summary>
    /// True when the session was hydrated from a hand-edited <c>jwtAccessToken</c> in login preferences.
    /// Disables refresh-token recovery; Credential Manager is left untouched.
    /// </summary>
    public bool IsJwtAccessTokenMode { get; set; }

    /// <summary>
    /// REST API version for this session (e.g. <c>"66.0"</c>), set after sign-in.
    /// </summary>
    public string ApiVersion =>
        _apiVersion ?? throw new InvalidOperationException(
            "Salesforce API version is not available until sign-in completes.");

    public bool IsApiVersionResolved => !string.IsNullOrWhiteSpace(_apiVersion);

    public long? IssuedAtUnixMs { get; set; }

    public bool IsLoggedIn =>
        !string.IsNullOrWhiteSpace(AccessToken) &&
        !string.IsNullOrWhiteSpace(InstanceUrl) &&
        !IsExpired();

    public void SetApiVersion(string apiVersion)
    {
        _apiVersion = Rest.SalesforceApiVersions.Normalize(apiVersion);
    }

    public void ClearApiVersion() => _apiVersion = null;

    public bool IsExpired()
    {
        if (IssuedAtUnixMs is null)
        {
            return false;
        }

        // Salesforce access tokens are typically short-lived; treat > 2 hours as expired
        // until refresh-token handling is ported.
        const long twoHoursMs = 2 * 60 * 60 * 1000;
        var nowMs = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        return nowMs - IssuedAtUnixMs.Value > twoHoursMs;
    }

    public void Invalidate()
    {
        AccessToken = null;
        RefreshToken = null;
        InstanceUrl = null;
        Id = null;
        DisplayName = null;
        IsJwtAccessTokenMode = false;
        IssuedAtUnixMs = null;
        ClearApiVersion();
    }
}
