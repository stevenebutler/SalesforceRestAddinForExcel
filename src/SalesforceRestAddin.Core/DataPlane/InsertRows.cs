using SalesforceRestAddin.Core.Rest;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Core.Tables;
using SalesforceRestAddin.Core.Values;

namespace SalesforceRestAddin.Core.DataPlane;

public sealed class InsertRowsInput
{
    public required ForceTableBinding Binding { get; init; }

    public required ForceTableSelection Selection { get; init; }

    public ConnectorOptions Options { get; init; } = ConnectorOptions.Default;
}

public static class InsertRows
{
    private static readonly StringComparer NewComparer = StringComparer.OrdinalIgnoreCase;

    public static async Task<DataOperationResult> RunAsync(
        SalesforceDataClient client,
        InsertRowsInput input,
        CancellationToken cancellationToken = default)
    {
        var limitError = SelectionLimits.ValidateSelection(input.Selection, new ConnectorOptions());
        if (limitError is not null)
        {
            return new DataOperationResult { ErrorSummary = limitError };
        }

        var records = new List<(int BodyRowIndex, Dictionary<string, object?> Record)>();
        foreach (var bodyRow in input.Selection.BodyRowIndices)
        {
            var idCell = input.Binding.Snapshot.Body[bodyRow, input.Binding.IdColumnIndex]?.ToString();
            if (!IsNewRow(idCell))
            {
                continue;
            }

            var record = new Dictionary<string, object?>
            {
                ["attributes"] = new Dictionary<string, object?> { ["type"] = input.Binding.Describe.Name },
            };
            var hasCreateable = false;
            foreach (var column in input.Binding.Columns.Where(c => !c.Field.IsId))
            {
                if (!FieldValueConverter.IsWritable(column.Field, forCreate: true))
                {
                    continue;
                }

                hasCreateable = true;
                var value = FieldValueConverter.ToSalesforceValue(
                    column.Field,
                    input.Binding.Snapshot.Body[bodyRow, column.ColumnIndex],
                    input.Options,
                    forCreate: true);
                if (value is not null)
                {
                    record[column.Field.Name] = value;
                }
            }

            if (!hasCreateable)
            {
                return new DataOperationResult
                {
                    ErrorSummary = $"Row {bodyRow + 3} has no createable fields.",
                };
            }

            records.Add((bodyRow, record));
        }

        if (records.Count == 0)
        {
            return new DataOperationResult { ErrorSummary = "No rows marked New for insert." };
        }

        var payloads = records.Select(r => r.Record).ToList();
        var batches = RecordBatchSplitter.Split(payloads, input.Options.CompositeBatchSize);
        var outcomes = new List<RowOutcome>();
        var index = 0;

        foreach (var batch in batches)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var results = await client.CreateAsync(
                batch,
                input.Options.SendPreventAutoAssignHeader,
                cancellationToken).ConfigureAwait(false);

            for (var i = 0; i < results.Count; i++)
            {
                var (bodyRow, _) = records[index];
                var result = results[i];
                outcomes.Add(new RowOutcome
                {
                    BodyRowIndex = bodyRow,
                    Succeeded = result.Success,
                    ErrorMessages = result.Errors,
                    IdWriteback = result.Success ? result.Id : null,
                });
                index++;
            }
        }

        return new DataOperationResult
        {
            RecordsProcessed = records.Count,
            RowOutcomes = outcomes,
        };
    }

    public static bool IsNewRow(string? idCell) =>
        idCell is not null && !string.IsNullOrWhiteSpace(idCell) && NewComparer.Equals(idCell.Trim(), "new");
}
