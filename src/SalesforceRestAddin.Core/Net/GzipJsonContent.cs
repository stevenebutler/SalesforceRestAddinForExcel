using System.IO.Compression;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;

namespace SalesforceRestAddin.Core.Net;

/// <summary>
/// Builds JSON <see cref="HttpContent"/> for Salesforce REST, optionally gzip-compressed
/// with <c>Content-Encoding: gzip</c> (see Salesforce REST compression headers).
/// </summary>
public static class GzipJsonContent
{
    /// <summary>
    /// Minimum uncompressed UTF-8 byte length before gzip is applied.
    /// Smaller bodies stay plain JSON (gzip can expand tiny payloads).
    /// </summary>
    public const int DefaultMinBytesForGzip = 512;

    public static HttpContent Create(string json, int minBytesForGzip = DefaultMinBytesForGzip)
    {
        if (json is null)
        {
            throw new ArgumentNullException(nameof(json));
        }

        var utf8 = Encoding.UTF8.GetBytes(json);
        if (utf8.Length < minBytesForGzip)
        {
            return new StringContent(json, Encoding.UTF8, "application/json");
        }

        using var output = new MemoryStream();
        using (var gzip = new GZipStream(output, CompressionLevel.Optimal, leaveOpen: true))
        {
            gzip.Write(utf8, 0, utf8.Length);
        }

        var content = new ByteArrayContent(output.ToArray());
        content.Headers.ContentType = new MediaTypeHeaderValue("application/json")
        {
            CharSet = "utf-8",
        };
        content.Headers.ContentEncoding.Add("gzip");
        return content;
    }
}
