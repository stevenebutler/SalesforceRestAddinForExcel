using System.Text;

namespace SalesforceRestAddin.Core.Tables;

/// <summary>
/// Builds a scrollable <see cref="DataOperationResult.ErrorSummary"/> from per-row failures
/// (Id-cell comments remain the sheet annotation).
/// </summary>
public static class RowOutcomeErrorSummary
{
    public static string? FromFailures(
        IReadOnlyList<RowOutcome> outcomes,
        int tableStartRow,
        string operationLabel,
        int? recordsProcessed = null)
    {
        var failures = outcomes.Where(o => !o.Succeeded).ToList();
        if (failures.Count == 0)
        {
            return null;
        }

        // Update only stores failed outcomes; prefer RecordsProcessed for the denominator.
        var total = recordsProcessed ?? outcomes.Count;
        if (total < failures.Count)
        {
            total = failures.Count;
        }

        var builder = new StringBuilder();
        builder.Append(operationLabel);
        builder.Append(" failed for ");
        builder.Append(failures.Count);
        builder.Append(" of ");
        builder.Append(total);
        builder.Append(" row(s). Details are also in Id cell comments.");
        builder.AppendLine();

        foreach (var outcome in failures)
        {
            var excelRow = tableStartRow + 2 + outcome.BodyRowIndex;
            builder.AppendLine();
            builder.Append("Row ");
            builder.Append(excelRow);
            builder.Append(':');
            if (outcome.ErrorMessages.Count == 0)
            {
                builder.Append(" Operation Failed");
                continue;
            }

            foreach (var message in outcome.ErrorMessages)
            {
                builder.AppendLine();
                builder.Append("  ");
                builder.Append(message);
            }
        }

        return builder.ToString();
    }
}
