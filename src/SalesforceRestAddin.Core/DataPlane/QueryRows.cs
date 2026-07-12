using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Rest;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Core.Tables;
using SalesforceRestAddin.Core.Values;

namespace SalesforceRestAddin.Core.DataPlane;

public sealed class QueryRowsInput
{
    public required ForceTableBinding Binding { get; init; }

    public required ForceTableSelection Selection { get; init; }

    public ConnectorOptions Options { get; init; } = ConnectorOptions.Default;

    public bool RefreshAll { get; init; }

    public IReferenceResolver? ReferenceResolver { get; init; }
}

public static class QueryRows
{
    public static async Task<DataOperationResult> RunAsync(
        SalesforceDataClient client,
        QueryRowsInput input,
        CancellationToken cancellationToken = default)
    {
        var limitError = SelectionLimits.ValidateSelection(input.Selection, input.Options);
        if (limitError is not null)
        {
            return new DataOperationResult { ErrorSummary = limitError };
        }

        var binding = input.Binding;
        var rowIndices = input.RefreshAll
            ? ExpandContiguousIdRows(binding)
            : input.Selection.BodyRowIndices;

        var ids = new List<string>();
        foreach (var row in rowIndices)
        {
            var id = binding.Snapshot.Body[row, binding.IdColumnIndex]?.ToString();
            if (!SalesforceId.IsValid(id))
            {
                continue;
            }

            ids.Add(SalesforceId.Normalize(id)!);
        }

        if (ids.Count == 0)
        {
            return new DataOperationResult { ErrorSummary = "No valid Salesforce Ids found in selection." };
        }

        var fields = binding.Columns
            .Where(c => !c.Field.IsId)
            .Select(c => c.Field.Name)
            .ToList();

        var batches = RecordBatchSplitter.Split(ids, input.Options.CompositeBatchSize);
        var recordsById = new Dictionary<string, Dictionary<string, object?>>(StringComparer.OrdinalIgnoreCase);
        foreach (var batch in batches)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var retrieved = await client.RetrieveAsync(
                binding.Describe.Name,
                batch,
                fields,
                cancellationToken).ConfigureAwait(false);
            var returnedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var record in retrieved)
            {
                if (record.TryGetValue("Id", out var idObj) && idObj is string id)
                {
                    recordsById[id] = record;
                    returnedIds.Add(id);
                }
            }

            var missingIds = batch.Where(id => !returnedIds.Contains(id)).ToList();
            if (missingIds.Count > 0)
            {
                SessionFlowTrace.Log(
                    $"Query Selected Rows: Salesforce returned no row(s) for Id(s): {string.Join(", ", missingIds)}");
            }
        }

        var dataColumns = binding.Columns.Where(c => !c.Field.IsId).ToList();
        var values = new object?[rowIndices.Count, dataColumns.Count];
        var formats = dataColumns.Select(c => FieldValueConverter.CreateColumnFormat(c.Field)).ToList();
        for (var r = 0; r < rowIndices.Count; r++)
        {
            var id = binding.Snapshot.Body[rowIndices[r], binding.IdColumnIndex]?.ToString();
            recordsById.TryGetValue(id ?? string.Empty, out var record);
            for (var c = 0; c < dataColumns.Count; c++)
            {
                object? raw = null;
                record?.TryGetValue(dataColumns[c].Field.Name, out raw);
                values[r, c] = FieldValueConverter.ToDisplayValue(
                    dataColumns[c].Field,
                    raw,
                    input.Options);
            }
        }

        return new DataOperationResult
        {
            Projection = new SheetProjection
            {
                Values = values,
                StartRow = binding.Snapshot.StartRow + 2 + rowIndices[0],
                StartColumn = binding.Snapshot.StartColumn + dataColumns[0].ColumnIndex,
                ColumnFormats = formats,
            },
            RecordsProcessed = ids.Count,
        };
    }

    private static IReadOnlyList<int> ExpandContiguousIdRows(ForceTableBinding binding)
    {
        var rows = new List<int>();
        for (var r = 0; r < binding.Snapshot.BodyRowCount; r++)
        {
            var id = binding.Snapshot.Body[r, binding.IdColumnIndex]?.ToString();
            if (SalesforceId.IsValid(id))
            {
                rows.Add(r);
            }
            else if (rows.Count > 0)
            {
                break;
            }
        }

        return rows;
    }
}
