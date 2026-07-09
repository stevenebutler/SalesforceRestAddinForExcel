using SalesforceRestAddin.Core.Tables;

namespace SalesforceRestAddin.Core.Soql;

public static class SoqlQueryBuilder
{
    public const int ExcelRowLimit = 1_048_570;

    public static string BuildSelectQuery(
        ForceTableBinding binding,
        string? whereClause,
        IReadOnlyList<string>? additionalReferenceIds = null,
        string? referenceJoinField = null)
    {
        var fields = binding.Columns
            .Select(c => c.Field.Name)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var where = BuildEffectiveWhere(whereClause, additionalReferenceIds, referenceJoinField);
        var whereSuffix = string.IsNullOrWhiteSpace(where) ? string.Empty : $" WHERE {where}";
        return $"SELECT {string.Join(", ", fields)} FROM {binding.Describe.Name}{whereSuffix}";
    }

    public static string BuildCountQuery(string objectApiName, string? whereClause)
    {
        var whereSuffix = string.IsNullOrWhiteSpace(whereClause) ? string.Empty : $" WHERE {whereClause}";
        return $"SELECT COUNT(Id) FROM {objectApiName}{whereSuffix}";
    }

    public static IReadOnlyList<string> BuildReferenceInBatches(
        IReadOnlyList<string> referenceIds,
        string referenceJoinField,
        int batchSize)
    {
        var batches = new List<string>();
        for (var i = 0; i < referenceIds.Count; i += batchSize)
        {
            var chunk = referenceIds.Skip(i).Take(batchSize).ToList();
            var inList = string.Join(", ", chunk.Select(id => $"'{SoqlEscape.EscapeLiteral(id)}'"));
            batches.Add($"{referenceJoinField} IN ({inList})");
        }

        return batches;
    }

    private static string BuildEffectiveWhere(
        string? whereClause,
        IReadOnlyList<string>? additionalReferenceIds,
        string? referenceJoinField)
    {
        var parts = new List<string>();
        if (whereClause is not null && !string.IsNullOrWhiteSpace(whereClause))
        {
            parts.Add(whereClause);
        }

        if (additionalReferenceIds is { Count: > 0 }
            && referenceJoinField is not null
            && !string.IsNullOrWhiteSpace(referenceJoinField))
        {
            var inList = string.Join(
                ", ",
                additionalReferenceIds.Select(id => $"'{SoqlEscape.EscapeLiteral(id)}'"));
            parts.Add($"{referenceJoinField} IN ({inList})");
        }

        return string.Join(" and ", parts);
    }
}
