using System.Net;
using System.Net.Http;

namespace SalesforceRestAddin.Core.Net;

/// <summary>
/// Shared <see cref="HttpClient"/> construction for Salesforce REST/OAuth.
/// Enables response decompression so Salesforce may return gzip/deflate bodies
/// (<c>Accept-Encoding</c> is added by the handler).
/// </summary>
public static class SalesforceHttpClientFactory
{
    public static HttpClient Create() => new(CreateHandler());

    public static HttpClientHandler CreateHandler() =>
        new()
        {
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,
        };
}
