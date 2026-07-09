namespace SalesforceRestAddin.Core.Tables;

public sealed class BoundColumn
{
    public required int ColumnIndex { get; init; }

    public required FieldDescriptor Field { get; init; }
}

public sealed class ForceTableBinding
{
    public required ForceTableSnapshot Snapshot { get; init; }

    public required SObjectDescribe Describe { get; init; }

    public required FieldCatalog Catalog { get; init; }

    public required IReadOnlyList<BoundColumn> Columns { get; init; }

    public int IdColumnIndex => Snapshot.IdColumnIndex;
}
