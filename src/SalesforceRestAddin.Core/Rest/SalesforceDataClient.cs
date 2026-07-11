using System;
using System.Diagnostics;
using System.Net.Http;
using System.Text.Json;
using SalesforceRestAddin.Core;
using SalesforceRestAddin.Core.Net;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Core.Tables;

namespace SalesforceRestAddin.Core.Rest;

public sealed class SalesforceDataClient
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    };

    private readonly SalesforceAuthenticatedClient _client;
    private readonly SessionContext _session;
    private readonly IMetadataCache _metadataCache;
    private readonly ParsedMetadataCache _parsedCache;

    public SalesforceDataClient(
        SalesforceAuthenticatedClient client,
        SessionContext session,
        IMetadataCache? metadataCache = null,
        ParsedMetadataCache? parsedCache = null)
    {
        _client = client ?? throw new ArgumentNullException(nameof(client));
        _session = session ?? throw new ArgumentNullException(nameof(session));
        _metadataCache = metadataCache ?? NoOpMetadataCache.Instance;
        _parsedCache = parsedCache ?? new ParsedMetadataCache();
    }

    public async Task<SObjectDescribe> DescribeAsync(string objectApiName, CancellationToken cancellationToken = default)
    {
        var instanceUrl = _session.InstanceUrl ?? string.Empty;
        var apiVersion = _session.ApiVersion ?? string.Empty;

        if (_parsedCache.TryGetDescribe(instanceUrl, apiVersion, objectApiName, out var parsed)
            && parsed is not null)
        {
            SessionFlowTrace.Log($"REST describe {objectApiName}: memory cache hit");
            return parsed;
        }

        if (_metadataCache.TryGetDescribe(instanceUrl, apiVersion, objectApiName, out var cached)
            && !string.IsNullOrWhiteSpace(cached))
        {
            SessionFlowTrace.Log($"REST describe {objectApiName}: disk cache hit");
            var fromDisk = SObjectDescribeJsonParser.Parse(cached!);
            _parsedCache.SetDescribe(instanceUrl, apiVersion, objectApiName, fromDisk);
            return fromDisk;
        }

        var url = $"{instanceUrl}/services/data/v{apiVersion}/sobjects/{objectApiName}/describe";
        using var request = new HttpRequestMessage(HttpMethod.Get, new Uri(url));
        var json = await SendAndReadAsync(request, $"describe {objectApiName}", cancellationToken).ConfigureAwait(false);
        var describe = SObjectDescribeJsonParser.Parse(json);
        _parsedCache.SetDescribe(instanceUrl, apiVersion, objectApiName, describe);
        _metadataCache.SetDescribe(instanceUrl, apiVersion, objectApiName, json);
        return describe;
    }

    public async Task<IReadOnlyList<SObjectSummary>> ListObjectsAsync(CancellationToken cancellationToken = default)
    {
        var instanceUrl = _session.InstanceUrl ?? string.Empty;
        var apiVersion = _session.ApiVersion ?? string.Empty;

        if (_parsedCache.TryGetObjectList(instanceUrl, apiVersion, out var parsedList)
            && parsedList is not null)
        {
            SessionFlowTrace.Log("REST list sobjects: memory cache hit");
            return parsedList;
        }

        if (_metadataCache.TryGetObjectList(instanceUrl, apiVersion, out var cached)
            && !string.IsNullOrWhiteSpace(cached))
        {
            SessionFlowTrace.Log("REST list sobjects: disk cache hit");
            var fromDisk = SObjectListJsonParser.Parse(cached!);
            _parsedCache.SetObjectList(instanceUrl, apiVersion, fromDisk);
            return fromDisk;
        }

        var url = $"{instanceUrl}/services/data/v{apiVersion}/sobjects/";
        using var request = new HttpRequestMessage(HttpMethod.Get, new Uri(url));
        var json = await SendAndReadAsync(request, "list sobjects", cancellationToken).ConfigureAwait(false);
        var objects = SObjectListJsonParser.Parse(json);
        _parsedCache.SetObjectList(instanceUrl, apiVersion, objects);
        _metadataCache.SetObjectList(instanceUrl, apiVersion, json);
        return objects;
    }

    public async Task<QueryResultPage> QueryAsync(
        string soql,
        string? nextRecordsUrl = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(nextRecordsUrl))
        {
            SessionFlowTrace.Log($"SOQL: {soql}");
        }
        else
        {
            SessionFlowTrace.Log($"SOQL query-more: {nextRecordsUrl}");
        }

        var url = BuildQueryUrl(soql, nextRecordsUrl);
        using var request = new HttpRequestMessage(HttpMethod.Get, new Uri(url));
        var json = await SendAndReadAsync(request, "query", cancellationToken).ConfigureAwait(false);
        if (string.IsNullOrWhiteSpace(nextRecordsUrl) && IsAggregateCountSoql(soql))
        {
            SessionFlowTrace.Log($"SOQL response: {json}");
        }

        return ParseQueryPage(json);
    }

    /// <summary>
    /// Aggregate COUNT queries return <c>totalSize: 1</c> with the real count in
    /// <c>records[0].expr0</c> (or an aliased field). Prefer that over <see cref="QueryResultPage.TotalSize"/>.
    /// </summary>
    public static int ResolveAggregateCount(QueryResultPage page)
    {
        if (page.Records.Count == 1)
        {
            var record = page.Records[0];
            foreach (var key in new[] { "expr0", "cnt", "count" })
            {
                if (record.TryGetValue(key, out var value) && TryToInt(value, out var n))
                {
                    return n;
                }
            }

            // Unaliased COUNT may appear as the sole non-attributes numeric property.
            foreach (var pair in record)
            {
                if (TryToInt(pair.Value, out var n))
                {
                    return n;
                }
            }
        }

        return page.TotalSize;
    }

    private static bool IsAggregateCountSoql(string soql) =>
        soql.IndexOf("COUNT(", StringComparison.OrdinalIgnoreCase) >= 0;

    private static bool TryToInt(object? value, out int n)
    {
        switch (value)
        {
            case int i:
                n = i;
                return true;
            case long l when l >= int.MinValue && l <= int.MaxValue:
                n = (int)l;
                return true;
            case double d when d >= int.MinValue && d <= int.MaxValue && Math.Abs(d % 1) < double.Epsilon:
                n = (int)d;
                return true;
            case string s when int.TryParse(s, out n):
                return true;
            default:
                n = 0;
                return false;
        }
    }

    public async Task<IReadOnlyList<Dictionary<string, object?>>> RetrieveAsync(
        string objectApiName,
        IReadOnlyList<string> ids,
        IReadOnlyList<string> fields,
        CancellationToken cancellationToken = default)
    {
        var fieldList = string.Join(",", fields);
        var idList = string.Join(",", ids);
        var url =
            $"{_session.InstanceUrl}/services/data/v{_session.ApiVersion}/composite/sobjects/{objectApiName}?ids={idList}&fields={fieldList}";
        using var request = new HttpRequestMessage(HttpMethod.Get, new Uri(url));
        var json = await SendAndReadAsync(request, $"retrieve {objectApiName}", cancellationToken).ConfigureAwait(false);
        using var document = JsonDocument.Parse(json);
        return document.RootElement.EnumerateArray()
            .Select(ParseRecord)
            .ToList();
    }

    public async Task<IReadOnlyList<SaveResult>> CreateAsync(
        IReadOnlyList<Dictionary<string, object?>> records,
        bool sendPreventAutoAssignHeader,
        CancellationToken cancellationToken = default)
    {
        return await SaveAsync("POST", records, sendPreventAutoAssignHeader, cancellationToken).ConfigureAwait(false);
    }

    public async Task<IReadOnlyList<SaveResult>> UpdateAsync(
        IReadOnlyList<Dictionary<string, object?>> records,
        bool sendPreventAutoAssignHeader,
        CancellationToken cancellationToken = default)
    {
        return await SaveAsync("PATCH", records, sendPreventAutoAssignHeader, cancellationToken).ConfigureAwait(false);
    }

    public async Task<IReadOnlyList<DeleteResult>> DeleteAsync(
        IReadOnlyList<string> ids,
        CancellationToken cancellationToken = default)
    {
        var idParam = string.Join(",", ids);
        var url = $"{_session.InstanceUrl}/services/data/v{_session.ApiVersion}/composite/sobjects?ids={idParam}&allOrNone=false";
        using var request = new HttpRequestMessage(HttpMethod.Delete, new Uri(url));
        var json = await SendAndReadAsync(request, "delete", cancellationToken).ConfigureAwait(false);
        using var document = JsonDocument.Parse(json);
        return document.RootElement.EnumerateArray().Select(element =>
        {
            var success = element.GetProperty("success").GetBoolean();
            var id = element.TryGetProperty("id", out var idEl) ? idEl.GetString() : null;
            var errors = ParseErrors(element);
            return new DeleteResult { Success = success, Id = id, Errors = errors };
        }).ToList();
    }

    private string BuildQueryUrl(string soql, string? nextRecordsUrl)
    {
        if (nextRecordsUrl is null || string.IsNullOrWhiteSpace(nextRecordsUrl))
        {
            return $"{_session.InstanceUrl}/services/data/v{_session.ApiVersion}/query?q={Uri.EscapeDataString(soql)}";
        }

        if (nextRecordsUrl.StartsWith("http", StringComparison.OrdinalIgnoreCase))
        {
            return nextRecordsUrl;
        }

        return $"{_session.InstanceUrl?.TrimEnd('/')}{nextRecordsUrl}";
    }

    private async Task<string> SendAndReadAsync(
        HttpRequestMessage request,
        string operation,
        CancellationToken cancellationToken)
    {
        SessionFlowTrace.Log($"REST {operation}: {request.Method} {request.RequestUri}");
        var stopwatch = Stopwatch.StartNew();
        var responseCompleted = false;
        try
        {
            using var response = await _client.SendAsync(request, cancellationToken).ConfigureAwait(false);
            var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            stopwatch.Stop();
            SessionFlowTrace.Log(
                $"REST {operation} completed: {(int)response.StatusCode} {response.ReasonPhrase} " +
                $"bodyChars={body.Length} elapsed={stopwatch.Elapsed.TotalMilliseconds:0}ms");
            responseCompleted = true;

            if (!response.IsSuccessStatusCode)
            {
                SessionFlowTrace.Log(
                    $"REST {operation} failed: {(int)response.StatusCode} {response.ReasonPhrase}");
                if (!string.IsNullOrWhiteSpace(body))
                {
                    SessionFlowTrace.Log($"Response body: {TrimForLog(body)}");
                }
            }

            response.EnsureSuccessStatusCode();
            return body;
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            if (!responseCompleted)
            {
                SessionFlowTrace.Log(
                    $"REST {operation} transport failed: {ex.GetType().Name} " +
                    $"elapsed={stopwatch.Elapsed.TotalMilliseconds:0}ms");
            }
            throw;
        }
    }

    private static string TrimForLog(string body, int maxLength = 2000) =>
        body.Length <= maxLength ? body : body.Substring(0, maxLength) + "...";

    private async Task<IReadOnlyList<SaveResult>> SaveAsync(
        string method,
        IReadOnlyList<Dictionary<string, object?>> records,
        bool sendPreventAutoAssignHeader,
        CancellationToken cancellationToken)
    {
        var url = $"{_session.InstanceUrl}/services/data/v{_session.ApiVersion}/composite/sobjects";
        var httpMethod = method == "PATCH" ? new HttpMethod("PATCH") : HttpMethod.Post;
        using var request = new HttpRequestMessage(httpMethod, new Uri(url));
        if (sendPreventAutoAssignHeader)
        {
            request.Headers.TryAddWithoutValidation("Sforce-Auto-Assign", "FALSE");
        }

        var payload = JsonSerializer.Serialize(new { records, allOrNone = false }, JsonOptions);
        SessionFlowTrace.Log($"REST {method.ToLowerInvariant()} request body: {TrimForLog(payload)}");
        request.Content = GzipJsonContent.Create(payload);
        var json = await SendAndReadAsync(request, method.ToLowerInvariant(), cancellationToken).ConfigureAwait(false);
        using var document = JsonDocument.Parse(json);
        return document.RootElement.EnumerateArray().Select(element =>
        {
            var success = element.GetProperty("success").GetBoolean();
            var id = element.TryGetProperty("id", out var idEl) ? idEl.GetString() : null;
            return new SaveResult { Success = success, Id = id, Errors = ParseErrors(element) };
        }).ToList();
    }

    private static QueryResultPage ParseQueryPage(string json)
    {
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        var records = root.GetProperty("records").EnumerateArray().Select(ParseRecord).ToList();
        return new QueryResultPage
        {
            TotalSize = root.GetProperty("totalSize").GetInt32(),
            Done = root.GetProperty("done").GetBoolean(),
            NextRecordsUrl = root.TryGetProperty("nextRecordsUrl", out var next) ? next.GetString() : null,
            Records = records,
        };
    }

    private static Dictionary<string, object?> ParseRecord(JsonElement element)
    {
        var record = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in element.EnumerateObject())
        {
            if (property.NameEquals("attributes"))
            {
                continue;
            }

            record[property.Name] = ParseJsonValue(property.Value);
        }

        return record;
    }

    private static object? ParseJsonValue(JsonElement value) =>
        value.ValueKind switch
        {
            JsonValueKind.Null => null,
            JsonValueKind.String => value.GetString(),
            JsonValueKind.Number => value.TryGetInt64(out var l) ? l : value.GetDouble(),
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Object => value.EnumerateObject()
                .ToDictionary(p => p.Name, p => ParseJsonValue(p.Value), StringComparer.OrdinalIgnoreCase),
            _ => value.GetRawText(),
        };

    private static IReadOnlyList<string> ParseErrors(JsonElement element)
    {
        if (!element.TryGetProperty("errors", out var errors) || errors.ValueKind != JsonValueKind.Array)
        {
            return Array.Empty<string>();
        }

        return errors.EnumerateArray()
            .Select(e =>
            {
                var status = e.TryGetProperty("statusCode", out var s) ? s.GetString() : null;
                var message = e.TryGetProperty("message", out var m) ? m.GetString() : null;
                return $"{status}: {message}";
            })
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .ToList()!;
    }
}
