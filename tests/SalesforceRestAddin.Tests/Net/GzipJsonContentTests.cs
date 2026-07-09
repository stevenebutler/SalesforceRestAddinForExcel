using System.IO.Compression;
using System.Net;
using System.Net.Http;
using System.Text;
using SalesforceRestAddin.Core.Net;

namespace SalesforceRestAddin.Tests.Net;

public sealed class GzipJsonContentTests
{
    [Test]
    public async Task Create_Below_Threshold_Is_Plain_Json()
    {
        var json = "{\"a\":1}";
        using var content = GzipJsonContent.Create(json, minBytesForGzip: 512);

        await Assert.That(content.Headers.ContentEncoding).IsEmpty();
        await Assert.That(content.Headers.ContentType!.MediaType).IsEqualTo("application/json");
        await Assert.That(await content.ReadAsStringAsync()).IsEqualTo(json);
    }

    [Test]
    public async Task Create_At_Or_Above_Threshold_Is_Gzip_With_Content_Encoding()
    {
        var json = new string('x', 600);
        using var content = GzipJsonContent.Create(json, minBytesForGzip: 512);

        await Assert.That(content.Headers.ContentEncoding.ToString()).IsEqualTo("gzip");
        await Assert.That(content.Headers.ContentType!.MediaType).IsEqualTo("application/json");

        var compressed = await content.ReadAsByteArrayAsync();
        await using var input = new MemoryStream(compressed);
        await using var gzip = new GZipStream(input, CompressionMode.Decompress);
        using var reader = new StreamReader(gzip, Encoding.UTF8);
        var decompressed = await reader.ReadToEndAsync();
        await Assert.That(decompressed).IsEqualTo(json);
    }

    [Test]
    public async Task Factory_Handler_Enables_Automatic_Decompression()
    {
        var handler = SalesforceHttpClientFactory.CreateHandler();
        await Assert.That(handler.AutomaticDecompression)
            .IsEqualTo(DecompressionMethods.GZip | DecompressionMethods.Deflate);
    }
}
