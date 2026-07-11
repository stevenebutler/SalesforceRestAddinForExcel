using SalesforceRestAddin.Core.Rest;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Core.Tables;
using SalesforceRestAddin.Core.Values;

namespace SalesforceRestAddin.Core.DataPlane;

public sealed class UpdateCellsInput
{
    public required ForceTableBinding Binding { get; init; }

    public required ForceTableSelection Selection { get; init; }

    public ConnectorOptions Options { get; init; } = ConnectorOptions.Default;

    public IReferenceResolver? ReferenceResolver { get; init; }
}

public static class UpdateCells
{
    public static DataOperationResult Plan(ForceTableBinding binding, ForceTableSelection selection, ConnectorOptions options)
    {
        var limitError = SelectionLimits.ValidateSelection(selection, options, allowMultiArea: true);
        if (limitError is not null)
        {
            return new DataOperationResult { ErrorSummary = limitError };
        }

        var outOfTable = ValidateSelectionWithinTable(binding, selection);
        if (outOfTable is not null)
        {
            return new DataOperationResult { ErrorSummary = outOfTable };
        }

        var records = BuildRecords(binding, selection, options, out var hasUpdateable);
        if (!hasUpdateable)
        {
            return new DataOperationResult { ErrorSummary = "No updatable columns selected." };
        }

        if (records.Count == 0)
        {
            return new DataOperationResult { ErrorSummary = "No rows to update." };
        }

        return new DataOperationResult { RecordsProcessed = records.Count };
    }

    /// <summary>
    /// Ensures every selected body row/column falls inside the captured table snapshot.
    /// </summary>
    public static string? ValidateSelectionWithinTable(ForceTableBinding binding, ForceTableSelection selection)
    {
        var columnCount = binding.Snapshot.ColumnCount;
        var bodyRowCount = binding.Snapshot.BodyRowCount;

        foreach (var bodyRow in selection.BodyRowIndices)
        {
            if (bodyRow < 0 || bodyRow >= bodyRowCount)
            {
                return OutsideTableMessage;
            }

            foreach (var col in selection.ColumnsForBodyRow(bodyRow))
            {
                if (col < 0 || col >= columnCount)
                {
                    return OutsideTableMessage;
                }
            }
        }

        return null;
    }

    public const string OutsideTableMessage =
        "Selection includes cells outside the ForceConnector table. Select cells within a single table only.";

    public static async Task<DataOperationResult> RunAsync(
        SalesforceDataClient client,
        UpdateCellsInput input,
        CancellationToken cancellationToken = default)
    {
        var plan = Plan(input.Binding, input.Selection, input.Options);
        if (!string.IsNullOrEmpty(plan.ErrorSummary))
        {
            return plan;
        }

        var records = BuildRecords(input.Binding, input.Selection, input.Options, out _);
        var batches = RecordBatchSplitter.Split(records, input.Options.CompositeBatchSize);
        var outcomes = new List<RowOutcome>();
        var rowIndex = 0;

        foreach (var batch in batches)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var results = await client.UpdateAsync(
                batch,
                input.Options.SendPreventAutoAssignHeader,
                cancellationToken).ConfigureAwait(false);

            for (var i = 0; i < results.Count; i++)
            {
                var result = results[i];
                if (!result.Success)
                {
                    outcomes.Add(new RowOutcome
                    {
                        BodyRowIndex = input.Selection.BodyRowIndices[rowIndex],
                        Succeeded = false,
                        ErrorMessages = result.Errors,
                    });
                }

                rowIndex++;
            }
        }

        return new DataOperationResult
        {
            RecordsProcessed = records.Count,
            RowOutcomes = outcomes,
            ErrorSummary = RowOutcomeErrorSummary.FromFailures(
                outcomes,
                input.Binding.Snapshot.StartRow,
                "Update",
                records.Count),
        };
    }

    public static List<Dictionary<string, object?>> BuildRecords(
        ForceTableBinding binding,
        ForceTableSelection selection,
        ConnectorOptions options,
        out bool hasUpdateable)
    {
        hasUpdateable = false;
        var records = new List<Dictionary<string, object?>>();
        // Default: omit AutoFilter/manually hidden rows and columns unless IncludeHiddenCells.
        var visibleRows = options.IncludeHiddenCells
            ? selection.BodyRowIndices.ToList()
            : HiddenSelectionFilter.VisibleBodyRows(binding.Snapshot, selection.BodyRowIndices).ToList();
        SessionFlowTrace.Log(
            $"UpdateCells: includeHidden={options.IncludeHiddenCells} " +
            $"selectedBodyRows={selection.BodyRowIndices.Count} " +
            $"capturedHiddenBodyRows={binding.Snapshot.HiddenRowIndices?.Count ?? 0} " +
            $"rowsForPayload={visibleRows.Count}");

        foreach (var bodyRow in visibleRows)
        {
            var id = SalesforceId.Normalize(binding.Snapshot.Body[bodyRow, binding.IdColumnIndex]?.ToString());
            if (!SalesforceId.IsValid(id))
            {
                continue;
            }

            var selectedCols = selection.ColumnsForBodyRow(bodyRow);
            var visibleCols = options.IncludeHiddenCells
                ? selectedCols.ToList()
                : HiddenSelectionFilter.VisibleColumns(binding.Snapshot, selectedCols).ToList();

            // composite/sobjects PATCH requires attributes.type on every record (same as create).
            var record = new Dictionary<string, object?>
            {
                ["attributes"] = new Dictionary<string, object?> { ["type"] = binding.Describe.Name },
                ["Id"] = id,
            };
            foreach (var column in binding.Columns.Where(col => visibleCols.Contains(col.ColumnIndex)))
            {
                if (!FieldValueConverter.IsWritable(column.Field, forCreate: false))
                {
                    continue;
                }

                hasUpdateable = true;
                // Include null so cleared cells become Salesforce null (not omitted / not "").
                record[column.Field.Name] = FieldValueConverter.ToSalesforceValue(
                    column.Field,
                    binding.Snapshot.Body[bodyRow, column.ColumnIndex],
                    options,
                    forCreate: false);
            }

            // attributes + Id alone is not an update; need at least one field value.
            if (record.Count > 2)
            {
                records.Add(record);
            }
        }

        return records;
    }
}
