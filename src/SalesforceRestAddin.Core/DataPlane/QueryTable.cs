using System.Net.Http;
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
            cancellationToken).ConfigureAwait(false);

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
            cancellationToken).ConfigureAwait(false);

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
                values[r, c] = FieldValueConverter.ToDisplayValue(
                    column.Field,
                    raw,
                    input.Options);
            }
        }

        return new DataOperationResult
        {
            Projection = new SheetProjection
            {
                Values = values,
                StartRow = binding.Snapshot.StartRow + 2,
                StartColumn = binding.Snapshot.StartColumn,
                ColumnFormats = formats,
            },
            ClearBody = CreateClearBody(binding, records.Count),
            RecordsProcessed = records.Count,
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

    private static DataOperationResult BuildErrorProjection(ForceTableBinding binding, string message)
    {
        var values = new object?[1, 1] { { "#Err" } };
        return new DataOperationResult
        {
            Projection = new SheetProjection
            {
                Values = values,
                StartRow = binding.Snapshot.StartRow + 2,
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
