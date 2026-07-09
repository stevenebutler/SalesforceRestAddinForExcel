namespace SalesforceRestAddin.Core.OAuth;

/// <summary>
/// Parses Salesforce OAuth redirect callback URLs.
/// </summary>
public static class OAuthRedirectParser
{
    public static bool TryGetAuthorizationCode(string callbackUrl, string expectedRedirectUri, out string? code, out string? error)
    {
        code = null;
        error = null;

        if (!callbackUrl.StartsWith(expectedRedirectUri, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (!Uri.TryCreate(callbackUrl, UriKind.Absolute, out var uri))
        {
            return false;
        }

        var query = ParseQuery(uri.Query);
        if (query.TryGetValue("error", out var errorValue) && !string.IsNullOrWhiteSpace(errorValue))
        {
            error = errorValue;
            return false;
        }

        if (query.TryGetValue("code", out var codeValue) && !string.IsNullOrWhiteSpace(codeValue))
        {
            code = codeValue;
            return true;
        }

        return false;
    }

    private static Dictionary<string, string> ParseQuery(string query)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrEmpty(query))
        {
            return result;
        }

        var text = query.StartsWith('?') ? query[1..] : query;
        foreach (var pair in text.Split('&', StringSplitOptions.RemoveEmptyEntries))
        {
            var separator = pair.IndexOf('=');
            if (separator < 0)
            {
                result[Uri.UnescapeDataString(pair)] = string.Empty;
                continue;
            }

            var name = Uri.UnescapeDataString(pair[..separator]);
            var value = Uri.UnescapeDataString(pair[(separator + 1)..]);
            result[name] = value;
        }

        return result;
    }
}
