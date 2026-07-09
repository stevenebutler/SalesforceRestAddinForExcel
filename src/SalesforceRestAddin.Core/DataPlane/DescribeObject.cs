using SalesforceRestAddin.Core.Rest;
using SalesforceRestAddin.Core.Tables;

namespace SalesforceRestAddin.Core.DataPlane;

public static class DescribeObject
{
    private static readonly HashSet<string> NamedStandardFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "Id", "Name", "OwnerId", "CreatedDate", "CreatedById", "LastModifiedDate", "LastModifiedById",
    };

    private static readonly string[] GridHeaders =
    [
        "Label", "API Name", "Type", "Length", "Precision", "Scale", "Nillable", "Createable", "Updateable",
        "Custom", "Reference Targets", "Picklist Values",
    ];

    public static async Task<DataOperationResult> RunAsync(
        SalesforceDataClient client,
        IReadOnlyList<string> objectApiNames,
        CancellationToken cancellationToken = default)
    {
        if (objectApiNames.Count != 1)
        {
            return new DataOperationResult { ErrorSummary = "DescribeObject.RunAsync expects a single object name." };
        }

        try
        {
            var describe = await client.DescribeAsync(objectApiNames[0], cancellationToken).ConfigureAwait(false);
            return new DataOperationResult
            {
                Projection = BuildProjection(describe),
                RecordsProcessed = 1,
            };
        }
        catch
        {
            return new DataOperationResult { ErrorSummary = $"Describe failed for {objectApiNames[0]}." };
        }
    }

    public static SheetProjection BuildProjection(SObjectDescribe describe)
    {
        var rows = BuildRows(describe);
        var values = new object?[rows.Count + 2, GridHeaders.Length];
        values[0, 0] = describe.Label;
        values[0, 1] = describe.Name;
        for (var c = 0; c < GridHeaders.Length; c++)
        {
            values[1, c] = GridHeaders[c];
        }

        for (var r = 0; r < rows.Count; r++)
        {
            for (var c = 0; c < rows[r].Length; c++)
            {
                values[r + 2, c] = rows[r][c];
            }
        }

        return new SheetProjection
        {
            Values = values,
            StartRow = 1,
            StartColumn = 1,
        };
    }

    public static IReadOnlyList<object?[]> BuildRows(SObjectDescribe describe)
    {
        var ordered = describe.Fields
            .OrderBy(f => GetSortKey(f))
            .ThenBy(f => f.Name, StringComparer.OrdinalIgnoreCase);

        return ordered.Select(f => new object?[]
        {
            f.Label,
            f.Name,
            f.Type,
            f.Length,
            f.Precision,
            f.Scale,
            f.Nillable,
            f.Createable,
            f.Updateable,
            f.Custom,
            string.Join(";", f.ReferenceTo),
            string.Join(";", f.PicklistValues),
        }).ToList();
    }

    private static int GetSortKey(FieldDescriptor field)
    {
        if (field.IsId)
        {
            return 0;
        }

        if (NamedStandardFields.Contains(field.Name))
        {
            return 1;
        }

        if (!field.Name.EndsWith("__c", StringComparison.Ordinal))
        {
            return 2;
        }

        return 3;
    }
}
