using SalesforceRestAddin.Core.Rest;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Core.Tables;

namespace SalesforceRestAddin.Core.DataPlane;

public sealed class DeleteRecordsInput
{
    public required ForceTableBinding Binding { get; init; }

    public required ForceTableSelection Selection { get; init; }

    public ConnectorOptions Options { get; init; } = ConnectorOptions.Default;
}

public static class DeleteRecords
{
    public static async Task<DataOperationResult> RunAsync(
        SalesforceDataClient client,
        DeleteRecordsInput input,
        CancellationToken cancellationToken = default)
    {
        var limitError = SelectionLimits.ValidateSelection(input.Selection, input.Options);
        if (limitError is not null)
        {
            return new DataOperationResult { ErrorSummary = limitError };
        }

        var ids = input.Selection.BodyRowIndices
            .Select(r => SalesforceId.Normalize(input.Binding.Snapshot.Body[r, input.Binding.IdColumnIndex]?.ToString()))
            .Where(id => SalesforceId.IsValid(id))
            .Select(id => id!)
            .ToList();

        if (ids.Count == 0)
        {
            return new DataOperationResult { ErrorSummary = "No Salesforce Ids found to delete." };
        }

        var batches = RecordBatchSplitter.Split(ids, input.Options.CompositeBatchSize);
        var outcomes = new List<RowOutcome>();
        var idIndex = 0;

        foreach (var batch in batches)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var results = await client.DeleteAsync(batch, cancellationToken).ConfigureAwait(false);
            for (var i = 0; i < results.Count; i++)
            {
                var bodyRow = input.Selection.BodyRowIndices[idIndex];
                var result = results[i];
                outcomes.Add(new RowOutcome
                {
                    BodyRowIndex = bodyRow,
                    Succeeded = result.Success,
                    ErrorMessages = result.Success
                        ? result.Errors
                        : new[] { "Delete Row Failed" }.Concat(result.Errors).ToList(),
                    IdDisplayOverride = result.Success ? "deleted" : null,
                });
                idIndex++;
            }
        }

        return new DataOperationResult
        {
            RecordsProcessed = ids.Count,
            RowOutcomes = outcomes,
        };
    }
}
