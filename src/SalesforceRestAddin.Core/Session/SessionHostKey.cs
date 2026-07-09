using System.Text.Json;
using System;
using SalesforceRestAddin.Core.OAuth;

namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Stable key for a Salesforce login host (tenant + environment + sandbox).
/// Each host stores its own refresh token.
/// </summary>
public static class SessionHostKey
{
    public static string FromLoginOptions(SalesforceLoginOptions options)
    {
        if (options is null)
        {
            throw new ArgumentNullException(nameof(options));
        }

        var baseUrl = SalesforceLoginUrlBuilder.BuildLoginBaseUrl(options);
        return new Uri(baseUrl).Host;
    }
}
