using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;

namespace SalesforceRestAddin.Core.Rest;

/// <summary>
/// Reads <c>GET /services/data/</c> to discover supported REST versions for the org.
/// </summary>
public sealed class SalesforceApiVersionResolver
{
    private readonly HttpClient _http;

    public SalesforceApiVersionResolver(HttpClient http)
    {
        _http = http ?? throw new ArgumentNullException(nameof(http));
    }

    public async Task<string> ResolveLatestAsync(
        string instanceUrl,
        string accessToken,
        CancellationToken cancellationToken = default)
    {
        var resolution = await ResolveForSessionAsync(instanceUrl, accessToken, null, cancellationToken)
            .ConfigureAwait(false);
        return resolution.SelectedVersion;
    }

    public async Task<ApiVersionResolution> ResolveForSessionAsync(
        string instanceUrl,
        string accessToken,
        string? preferredVersion = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(instanceUrl) || string.IsNullOrWhiteSpace(accessToken))
        {
            return SalesforceApiVersionSelector.Resolve(preferredVersion, Array.Empty<string>());
        }

        try
        {
            var supportedVersions = await GetSupportedVersionsAsync(instanceUrl, accessToken, cancellationToken)
                .ConfigureAwait(false);
            return SalesforceApiVersionSelector.Resolve(preferredVersion, supportedVersions);
        }
        catch (InvalidOperationException)
        {
            throw;
        }
        catch
        {
            return SalesforceApiVersionSelector.Resolve(preferredVersion, Array.Empty<string>());
        }
    }

    public async Task<IReadOnlyList<string>> GetSupportedVersionsAsync(
        string instanceUrl,
        string accessToken,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(instanceUrl) || string.IsNullOrWhiteSpace(accessToken))
        {
            return Array.Empty<string>();
        }

        var versionsUrl = $"{instanceUrl.TrimEnd('/')}/services/data/";
        using var request = new HttpRequestMessage(HttpMethod.Get, versionsUrl);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

        using var response = await _http.SendAsync(request, cancellationToken).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            return Array.Empty<string>();
        }

        var json = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
        return SalesforceApiVersions.ParseSupportedVersionsFromJson(json);
    }
}
