using System.Net.Http;
using System.Diagnostics;
using SalesforceRestAddin.Core;
using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Rest;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Core.Soql;
using SalesforceRestAddin.Core.Tables;
using SalesforceRestAddin.Core.Values;

namespace SalesforceRestAddin.Core.DataPlane;

public sealed class QueryTableInput
{
    public required ForceTableSnapshot Snapshot { get; init; }

    public ConnectorOptions Options { get; init; } = ConnectorOptions.Default;

    public bool ConfirmQueryTableDownload { get; init; }

    public IReferenceResolver? ReferenceResolver { get; init; }

    /// <summary>Optional progress sink for pagination (FR-QTD-7).</summary>
    public Action<OperationProgress>? Progress { get; init; }
}

/// <summary>
/// A single, ordered page prepared for writing to a worksheet. This type deliberately
/// contains only Core table data; the Excel host owns the STA hand-off and COM writes.
/// </summary>
internal sealed class QueryTablePage
{
    public int PageNumber { get; init; }

    public required SheetProjection Projection { get; init; }

    /// <summary>Present only on the first page, before its values are written.</summary>
    public ClearBodyRegion? ClearBody { get; init; }

    /// <summary>
    /// Present only on the first non-empty page. This is the complete result row count
    /// used by the Excel host to format the body before subsequent pages arrive.
    /// </summary>
    public int? TotalRowCount { get; init; }

    public int RecordsProcessed { get; init; }
}

internal sealed class PagedQueryResult
{
    public required DataOperationResult Result { get; init; }

    /// <summary>Combined duration of the SOQL/count HTTP requests in this query operation.</summary>
    public TimeSpan ApiElapsed { get; init; }
}

public static class QueryTable
{
    public static async Task<DataOperationResult> RunAsync(
        SalesforceDataClient client,
        QueryTableInput input,
        CancellationToken cancellationToken = default)
    {
        SessionFlowTrace.Log($"QueryTable: object={input.Snapshot.ObjectApiName}");

        var describe = await client.DescribeAsync(input.Snapshot.ObjectApiName, cancellationToken)
            .ConfigureAwait(false);
        var bindResult = ForceTableBinder.Bind(input.Snapshot, describe);
        if (!bindResult.Succeeded)
        {
            return new DataOperationResult
            {
                ErrorSummary = bindResult.Errors[0].Message,
            };
        }

        var binding = bindResult.Binding!;
        var criteria = await SoqlCriteriaParser.ParseAsync(
            input.Snapshot.CriteriaRow,
            binding.Catalog,
            input.Options,
            input.ReferenceResolver,
            cancellationToken,
            input.Snapshot.CriteriaReferenceIds).ConfigureAwait(false);

        if (!criteria.Succeeded)
        {
            return new DataOperationResult { ErrorSummary = criteria.Errors[0].Message };
        }

        var where = criteria.WhereClause ?? string.Empty;
        if (criteria.ReferenceJoinIds is { Count: > 0 } && criteria.ReferenceJoinField is not null)
        {
            var batches = SoqlQueryBuilder.BuildReferenceInBatches(
                criteria.ReferenceJoinIds,
                criteria.ReferenceJoinField,
                input.Options.CompositeBatchSize);
            return await RunBatchedReferenceQueriesAsync(
                client,
                binding,
                batches,
                where,
                input,
                cancellationToken).ConfigureAwait(false);
        }

        var soql = SoqlQueryBuilder.BuildSelectQuery(binding, where);
        if (input.ConfirmQueryTableDownload)
        {
            var countError = await ValidateCountAsync(client, binding, where, cancellationToken).ConfigureAwait(false);
            if (countError is not null)
            {
                return countError;
            }
        }

        return await RunQueryAndProjectAsync(client, binding, soql, input, cancellationToken)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Downloads a normal SOQL cursor one page at a time. The caller is responsible for
    /// handing pages to Excel (or another UI) and may return as soon as a page is queued,
    /// allowing the next queryMore request to overlap with the prior page write.
    /// Reference-join queries deliberately retain their existing accumulated-result path.
    /// </summary>
    internal static async Task<PagedQueryResult> RunPagedAsync(
        SalesforceDataClient client,
        QueryTableInput input,
        Func<QueryTablePage, CancellationToken, Task> enqueuePageAsync,
        CancellationToken cancellationToken = default)
    {
        if (enqueuePageAsync is null)
        {
            throw new ArgumentNullException(nameof(enqueuePageAsync));
        }

        SessionFlowTrace.Log($"QueryTable paged: object={input.Snapshot.ObjectApiName}");

        var describe = await client.DescribeAsync(input.Snapshot.ObjectApiName, cancellationToken)
            .ConfigureAwait(false);
        var bindResult = ForceTableBinder.Bind(input.Snapshot, describe);
        if (!bindResult.Succeeded)
        {
            return Complete(new DataOperationResult { ErrorSummary = bindResult.Errors[0].Message }, TimeSpan.Zero);
        }

        var binding = bindResult.Binding!;
        var criteria = await SoqlCriteriaParser.ParseAsync(
            input.Snapshot.CriteriaRow,
            binding.Catalog,
            input.Options,
            input.ReferenceResolver,
            cancellationToken,
            input.Snapshot.CriteriaReferenceIds).ConfigureAwait(false);
        if (!criteria.Succeeded)
        {
            return Complete(new DataOperationResult { ErrorSummary = criteria.Errors[0].Message }, TimeSpan.Zero);
        }

        if (criteria.ReferenceJoinIds is { Count: > 0 } && criteria.ReferenceJoinField is not null)
        {
            var batches = SoqlQueryBuilder.BuildReferenceInBatches(
                criteria.ReferenceJoinIds,
                criteria.ReferenceJoinField,
                input.Options.CompositeBatchSize);
            var referenceQueryStopwatch = Stopwatch.StartNew();
            var result = await RunBatchedReferenceQueriesAsync(
                client,
                binding,
                batches,
                criteria.WhereClause ?? string.Empty,
                input,
                cancellationToken).ConfigureAwait(false);
            referenceQueryStopwatch.Stop();
            return Complete(result, referenceQueryStopwatch.Elapsed);
        }

        var where = criteria.WhereClause ?? string.Empty;
        var soql = SoqlQueryBuilder.BuildSelectQuery(binding, where);
        var apiElapsed = TimeSpan.Zero;

        async Task<QueryResultPage> QueryAsyncTimed(string? nextRecordsUrl)
        {
            var stopwatch = Stopwatch.StartNew();
            try
            {
                return await client.QueryAsync(soql, nextRecordsUrl, cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                stopwatch.Stop();
                apiElapsed += stopwatch.Elapsed;
            }
        }

        if (input.ConfirmQueryTableDownload)
        {
            var countStopwatch = Stopwatch.StartNew();
            var countError = await ValidateCountAsync(client, binding, where, cancellationToken).ConfigureAwait(false);
            countStopwatch.Stop();
            apiElapsed += countStopwatch.Elapsed;
            if (countError is not null)
            {
                return Complete(countError, apiElapsed);
            }
        }

        var processed = 0;
        var isFirstPage = true;
        var pageNumber = 0;
        QueryResultPage? page = null;
        try
        {
            page = await QueryAsyncTimed(null).ConfigureAwait(false);
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (page.Records.Count == 0 && isFirstPage)
                {
                    var noRows = BuildNoRowsResult(binding);
                    pageNumber++;
                    await enqueuePageAsync(
                        new QueryTablePage
                        {
                            PageNumber = pageNumber,
                            Projection = noRows.Projection!,
                            ClearBody = noRows.ClearBody,
                            RecordsProcessed = 0,
                        },
                        cancellationToken).ConfigureAwait(false);
                    SessionFlowTrace.Log($"QueryTable page {pageNumber}: queued records=0 downloaded=0");
                    return Complete(new DataOperationResult { RecordsProcessed = 0 }, apiElapsed);
                }

                if (page.Records.Count > 0)
                {
                    int? totalRowCount = isFirstPage
                        ? Math.Max(page.TotalSize, page.Records.Count)
                        : null;
                    var pageResult = ProjectPage(binding, page.Records, input, processed, totalRowCount);
                    processed += page.Records.Count;
                    pageNumber++;
                    var queueStopwatch = Stopwatch.StartNew();
                    await enqueuePageAsync(
                        new QueryTablePage
                        {
                            PageNumber = pageNumber,
                            Projection = pageResult.Projection!,
                            ClearBody = pageResult.ClearBody,
                            TotalRowCount = totalRowCount,
                            RecordsProcessed = processed,
                        },
                        cancellationToken).ConfigureAwait(false);
                    queueStopwatch.Stop();
                    SessionFlowTrace.Log(
                        $"QueryTable page {pageNumber}: queued records={page.Records.Count} " +
                        $"downloaded={processed} queueWait={FormatElapsed(queueStopwatch.Elapsed)}");
                    ReportDownload(input, processed, page.TotalSize);
                }

                if (page.Done || page.NextRecordsUrl is null)
                {
                    return Complete(new DataOperationResult { RecordsProcessed = processed }, apiElapsed);
                }

                isFirstPage = false;
                cancellationToken.ThrowIfCancellationRequested();
                // enqueuePageAsync has only queued the current page. Start queryMore now so
                // its HTTP request can run while the UI thread writes that page.
                page = await QueryAsyncTimed(page.NextRecordsUrl).ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException)
        {
            return Complete(new DataOperationResult { WasCancelled = true, RecordsProcessed = processed }, apiElapsed);
        }
        catch (HttpRequestException ex)
        {
            SessionFlowTrace.LogException("QueryTable paged: HTTP request failed", ex);
            return Complete(BuildErrorProjection(binding, ExceptionChainText.Format(ex), processed), apiElapsed);
        }
    }

    private static PagedQueryResult Complete(DataOperationResult result, TimeSpan apiElapsed) => new()
    {
        Result = result,
        ApiElapsed = apiElapsed,
    };

    private static string FormatElapsed(TimeSpan elapsed) => $"{Math.Max(0, elapsed.TotalMilliseconds):0}ms";

    public static async Task<QueryTableCountResult> GetMatchCountAsync(
        SalesforceDataClient client,
        QueryTableInput input,
        CancellationToken cancellationToken = default)
    {
        var bindingResult = await TryBindAsync(client, input, cancellationToken).ConfigureAwait(false);
        if (!bindingResult.Succeeded)
        {
            return new QueryTableCountResult { ErrorSummary = bindingResult.ErrorSummary };
        }

        if (bindingResult.ReferenceJoin)
        {
            return new QueryTableCountResult { Count = null };
        }

        var countSoql = SoqlQueryBuilder.BuildCountQuery(bindingResult.Binding!.Describe.Name, bindingResult.Where!);
        var countPage = await client.QueryAsync(countSoql, null, cancellationToken).ConfigureAwait(false);
        return new QueryTableCountResult { Count = SalesforceDataClient.ResolveAggregateCount(countPage) };
    }

    private static async Task<BindingPlanResult> TryBindAsync(
        SalesforceDataClient client,
        QueryTableInput input,
        CancellationToken cancellationToken)
    {
        var describe = await client.DescribeAsync(input.Snapshot.ObjectApiName, cancellationToken)
            .ConfigureAwait(false);
        var bindResult = ForceTableBinder.Bind(input.Snapshot, describe);
        if (!bindResult.Succeeded)
        {
            return new BindingPlanResult { ErrorSummary = bindResult.Errors[0].Message };
        }

        var binding = bindResult.Binding!;
        var criteria = await SoqlCriteriaParser.ParseAsync(
            input.Snapshot.CriteriaRow,
            binding.Catalog,
            input.Options,
            input.ReferenceResolver,
            cancellationToken,
            input.Snapshot.CriteriaReferenceIds).ConfigureAwait(false);

        if (!criteria.Succeeded)
        {
            return new BindingPlanResult { ErrorSummary = criteria.Errors[0].Message };
        }

        return new BindingPlanResult
        {
            Binding = binding,
            Where = criteria.WhereClause ?? string.Empty,
            ReferenceJoin = criteria.ReferenceJoinIds is { Count: > 0 } && criteria.ReferenceJoinField is not null,
        };
    }

    private static async Task<DataOperationResult?> ValidateCountAsync(
        SalesforceDataClient client,
        ForceTableBinding binding,
        string where,
        CancellationToken cancellationToken)
    {
        var countSoql = SoqlQueryBuilder.BuildCountQuery(binding.Describe.Name, where);
        var countPage = await client.QueryAsync(countSoql, null, cancellationToken).ConfigureAwait(false);
        var count = SalesforceDataClient.ResolveAggregateCount(countPage);
        if (count > SoqlQueryBuilder.ExcelRowLimit)
        {
            return new DataOperationResult
            {
                ErrorSummary = $"Query would return {count} rows, exceeding the Excel limit.",
            };
        }

        return null;
    }

    private sealed class BindingPlanResult
    {
        public ForceTableBinding? Binding { get; init; }

        public string? Where { get; init; }

        public bool ReferenceJoin { get; init; }

        public string? ErrorSummary { get; init; }

        public bool Succeeded => string.IsNullOrEmpty(ErrorSummary);
    }

    private static async Task<DataOperationResult> RunBatchedReferenceQueriesAsync(
        SalesforceDataClient client,
        ForceTableBinding binding,
        IReadOnlyList<string> inClauses,
        string baseWhere,
        QueryTableInput input,
        CancellationToken cancellationToken)
    {
        var allRecords = new List<Dictionary<string, object?>>();
        foreach (var inClause in inClauses)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var where = string.IsNullOrWhiteSpace(baseWhere) ? inClause : $"{baseWhere} and {inClause}";
            var soql = SoqlQueryBuilder.BuildSelectQuery(binding, where);
            var page = await client.QueryAsync(soql, null, cancellationToken).ConfigureAwait(false);
            allRecords.AddRange(page.Records);
            ReportDownload(input, allRecords.Count, page.TotalSize);
            while (!page.Done && page.NextRecordsUrl is not null)
            {
                cancellationToken.ThrowIfCancellationRequested();
                page = await client.QueryAsync(soql, page.NextRecordsUrl, cancellationToken).ConfigureAwait(false);
                allRecords.AddRange(page.Records);
                ReportDownload(input, allRecords.Count, page.TotalSize);
            }
        }

        return ProjectRecords(binding, allRecords, input);
    }

    private static async Task<DataOperationResult> RunQueryAndProjectAsync(
        SalesforceDataClient client,
        ForceTableBinding binding,
        string soql,
        QueryTableInput input,
        CancellationToken cancellationToken)
    {
        var allRecords = new List<Dictionary<string, object?>>();
        QueryResultPage page;
        try
        {
            page = await client.QueryAsync(soql, null, cancellationToken).ConfigureAwait(false);
        }
        catch (HttpRequestException ex)
        {
            SessionFlowTrace.LogException("QueryTable: HTTP request failed", ex);
            return BuildErrorProjection(binding, ExceptionChainText.Format(ex));
        }

        allRecords.AddRange(page.Records);
        ReportDownload(input, allRecords.Count, page.TotalSize);
        var processed = page.Records.Count;
        while (!page.Done && page.NextRecordsUrl is not null)
        {
            cancellationToken.ThrowIfCancellationRequested();
            page = await client.QueryAsync(soql, page.NextRecordsUrl, cancellationToken).ConfigureAwait(false);
            allRecords.AddRange(page.Records);
            processed += page.Records.Count;
            ReportDownload(input, allRecords.Count, page.TotalSize);
        }

        if (allRecords.Count == 0)
        {
            return BuildNoRowsResult(binding);
        }

        var result = ProjectRecords(binding, allRecords, input);
        return new DataOperationResult
        {
            Projection = result.Projection,
            ClearBody = result.ClearBody,
            RecordsProcessed = processed,
            RowOutcomes = result.RowOutcomes,
        };
    }

    private static void ReportDownload(QueryTableInput input, int downloaded, int totalSize)
    {
        if (input.Progress is null)
        {
            return;
        }

        if (totalSize > 0)
        {
            input.Progress(OperationProgress.Download(downloaded, totalSize));
        }
        else
        {
            input.Progress(OperationProgress.DownloadUnknownTotal(downloaded));
        }
    }

    public static DataOperationResult ProjectRecords(
        ForceTableBinding binding,
        IReadOnlyList<Dictionary<string, object?>> records,
        QueryTableInput input)
    {
        if (records.Count == 0)
        {
            return BuildNoRowsResult(binding);
        }

        return new DataOperationResult
        {
            Projection = CreateProjection(binding, records, input, bodyRowOffset: 0),
            ClearBody = CreateClearBody(binding, records.Count),
            RecordsProcessed = records.Count,
        };
    }

    private static DataOperationResult ProjectPage(
        ForceTableBinding binding,
        IReadOnlyList<Dictionary<string, object?>> records,
        QueryTableInput input,
        int bodyRowOffset,
        int? totalSize)
    {
        return new DataOperationResult
        {
            Projection = CreateProjection(binding, records, input, bodyRowOffset),
            ClearBody = totalSize is int total
                ? CreateClearBody(binding, Math.Max(total, records.Count))
                : null,
            RecordsProcessed = records.Count,
        };
    }

    private static SheetProjection CreateProjection(
        ForceTableBinding binding,
        IReadOnlyList<Dictionary<string, object?>> records,
        QueryTableInput input,
        int bodyRowOffset)
    {
        var dataColumns = binding.Columns.ToList();
        var values = new object?[records.Count, dataColumns.Count];
        var formats = new List<ColumnFormat>();

        for (var c = 0; c < dataColumns.Count; c++)
        {
            var column = dataColumns[c];
            formats.Add(FieldValueConverter.CreateColumnFormat(column.Field));

            for (var r = 0; r < records.Count; r++)
            {
                records[r].TryGetValue(column.Field.Name, out var raw);
                values[r, c] = FieldValueConverter.ToDisplayValue(column.Field, raw, input.Options);
            }
        }

        return new SheetProjection
        {
            Values = values,
            StartRow = binding.Snapshot.StartRow + 2 + bodyRowOffset,
            StartColumn = binding.Snapshot.StartColumn,
            ColumnFormats = formats,
        };
    }

    private static DataOperationResult BuildNoRowsResult(ForceTableBinding binding)
    {
        var values = new object?[1, 1] { { "#N/F" } };
        return new DataOperationResult
        {
            Projection = new SheetProjection
            {
                Values = values,
                StartRow = binding.Snapshot.StartRow + 2,
                StartColumn = binding.Snapshot.StartColumn + binding.IdColumnIndex,
            },
            // Clear the full existing body (bugs.md #5), then write #N/F in the Id cell.
            ClearBody = CreateClearBody(binding, projectionRowCount: 1),
        };
    }

    /// <summary>
    /// Clears the table body within the entity column span: at least the existing body rows
    /// (so a smaller query does not leave stale rows), and enough rows for the new projection.
    /// </summary>
    private static ClearBodyRegion CreateClearBody(ForceTableBinding binding, int projectionRowCount)
    {
        var bodyStart = binding.Snapshot.StartRow + 2;
        var existingEnd = binding.Snapshot.BodyRowCount > 0
            ? bodyStart + binding.Snapshot.BodyRowCount - 1
            : bodyStart - 1;
        var projectionEnd = bodyStart + Math.Max(projectionRowCount, 1) - 1;

        return new ClearBodyRegion
        {
            StartRow = bodyStart,
            StartColumn = binding.Snapshot.StartColumn,
            EndColumn = binding.Snapshot.StartColumn + binding.Snapshot.ColumnCount - 1,
            EndRow = Math.Max(existingEnd, projectionEnd),
        };
    }

    private static DataOperationResult BuildErrorProjection(
        ForceTableBinding binding,
        string message,
        int bodyRowOffset = 0)
    {
        var values = new object?[1, 1] { { "#Err" } };
        return new DataOperationResult
        {
            Projection = new SheetProjection
            {
                Values = values,
                StartRow = binding.Snapshot.StartRow + 2 + bodyRowOffset,
                StartColumn = binding.Snapshot.StartColumn + binding.IdColumnIndex,
            },
            ErrorSummary = message,
        };
    }
}

public sealed class QueryTableCountResult
{
    public int? Count { get; init; }

    public string? ErrorSummary { get; init; }
}
