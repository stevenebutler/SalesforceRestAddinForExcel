using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Core.Rest;

/// <summary>
/// Fetches Salesforce user/org identity metadata from the OAuth <c>id</c> URL.
/// </summary>
public sealed class SalesforceIdentityClient
{
    private readonly HttpClient _http;

    public SalesforceIdentityClient(HttpClient http)
    {
        _http = http ?? throw new ArgumentNullException(nameof(http));
    }

    public async Task<SalesforceIdentityResponse> GetIdentityAsync(
        string identityUrl,
        string accessToken,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(identityUrl))
        {
            throw new ArgumentException("Identity URL is required.", nameof(identityUrl));
        }

        if (string.IsNullOrWhiteSpace(accessToken))
        {
            throw new ArgumentException("Access token is required.", nameof(accessToken));
        }

        var requestUri = identityUrl.IndexOf('?') >= 0
            ? $"{identityUrl}&version=latest"
            : $"{identityUrl}?version=latest";

        using var request = new HttpRequestMessage(HttpMethod.Get, requestUri);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

        using var response = await _http.SendAsync(request, cancellationToken).ConfigureAwait(false);
        var json = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);

        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException(
                $"Salesforce identity request failed ({(int)response.StatusCode}): {json}");
        }

        return SalesforceIdentityResponse.ParseJson(json);
    }

    public async Task<SalesforceIdentityResponse> GetIdentityForSessionAsync(
        SalesforceAuthenticatedClient authenticatedClient,
        SessionContext session,
        CancellationToken cancellationToken = default)
    {
        if (authenticatedClient is null)
        {
            throw new ArgumentNullException(nameof(authenticatedClient));
        }

        if (session is null)
        {
            throw new ArgumentNullException(nameof(session));
        }

        var identityUrl = session.Id;
        if (identityUrl is null || string.IsNullOrWhiteSpace(identityUrl))
        {
            throw new InvalidOperationException("Salesforce identity URL is not available in the session.");
        }

        var requestUri = identityUrl.IndexOf('?') >= 0
            ? $"{identityUrl}&version=latest"
            : $"{identityUrl}?version=latest";

        using var request = new HttpRequestMessage(HttpMethod.Get, requestUri);
        using var response = await authenticatedClient.SendAsync(request, cancellationToken).ConfigureAwait(false);
        var json = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);

        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException(
                $"Salesforce identity request failed ({(int)response.StatusCode}): {json}");
        }

        return SalesforceIdentityResponse.ParseJson(json);
    }
}

public sealed class SalesforceIdentityResponse
{
    public required string RawJson { get; init; }

    public static SalesforceIdentityResponse ParseJson(string json) =>
        new() { RawJson = json };

    /// <summary>Returns the identity <c>display_name</c> when present.</summary>
    public string? TryGetDisplayName()
    {
        using var document = JsonDocument.Parse(RawJson);
        if (!document.RootElement.TryGetProperty("display_name", out var property))
        {
            return null;
        }

        var value = property.GetString();
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    public string FormatSummary()
    {
        using var document = JsonDocument.Parse(RawJson);
        var builder = new StringBuilder();
        builder.AppendLine("Salesforce identity (GET id?version=latest):");
        AppendElement(builder, document.RootElement, indent: 0);
        return builder.ToString().TrimEnd();
    }

    private static void AppendElement(StringBuilder builder, JsonElement element, int indent)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
                foreach (var property in element.EnumerateObject())
                {
                    AppendIndented(builder, indent, $"{property.Name}:");
                    if (property.Value.ValueKind is JsonValueKind.Object or JsonValueKind.Array)
                    {
                        builder.AppendLine();
                        AppendElement(builder, property.Value, indent + 1);
                    }
                    else
                    {
                        builder.Append(' ');
                        AppendScalar(builder, property.Value);
                        builder.AppendLine();
                    }
                }

                break;

            case JsonValueKind.Array:
                var index = 0;
                foreach (var item in element.EnumerateArray())
                {
                    AppendIndented(builder, indent, $"[{index}]:");
                    if (item.ValueKind is JsonValueKind.Object or JsonValueKind.Array)
                    {
                        builder.AppendLine();
                        AppendElement(builder, item, indent + 1);
                    }
                    else
                    {
                        builder.Append(' ');
                        AppendScalar(builder, item);
                        builder.AppendLine();
                    }

                    index++;
                }

                break;

            default:
                AppendScalar(builder, element);
                builder.AppendLine();
                break;
        }
    }

    private static void AppendScalar(StringBuilder builder, JsonElement element)
    {
        builder.Append(element.ValueKind switch
        {
            JsonValueKind.String => element.GetString() ?? string.Empty,
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            JsonValueKind.Null => "(null)",
            _ => element.GetRawText(),
        });
    }

    private static void AppendIndented(StringBuilder builder, int indent, string text)
    {
        builder.Append(' ', indent * 2);
        builder.Append(text);
    }
}
