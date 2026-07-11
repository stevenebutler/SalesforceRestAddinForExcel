namespace SalesforceRestAddin.Core.DataPlane;

public enum ObjectPickerSortColumn
{
    Group = 0,
    Label = 1,
    ApiName = 2,
}

public enum ObjectPickerSortDirection
{
    Ascending = 0,
    Descending = 1,
}

public sealed record ObjectPickerSortKey(ObjectPickerSortColumn Column, ObjectPickerSortDirection Direction);

public sealed class ObjectPickerSortState
{
    private readonly GridSortState<ObjectPickerSortColumn> _sortState;

    public ObjectPickerSortState()
    {
        _sortState = new GridSortState<ObjectPickerSortColumn>(
            [
                new GridSortKey<ObjectPickerSortColumn>(ObjectPickerSortColumn.Group, GridSortDirection.Ascending),
                new GridSortKey<ObjectPickerSortColumn>(ObjectPickerSortColumn.Label, GridSortDirection.Ascending),
            ]);
    }

    public IReadOnlyList<ObjectPickerSortKey> Keys =>
        _sortState.Keys
            .Select(key =>
                new ObjectPickerSortKey(
                    key.Column,
                    key.Direction == GridSortDirection.Ascending
                        ? ObjectPickerSortDirection.Ascending
                        : ObjectPickerSortDirection.Descending))
            .ToArray();

    public void ResetDefault()
    {
        _sortState.ResetDefault();
    }

    public void Promote(ObjectPickerSortColumn column)
    {
        _sortState.Click(column);
    }

    public IReadOnlyList<ObjectPickerRow> Sort(IEnumerable<ObjectPickerRow> rows) =>
        rows.OrderBy(row => row, Comparer<ObjectPickerRow>.Create(Compare)).ToList();

    private int Compare(ObjectPickerRow? left, ObjectPickerRow? right)
    {
        if (ReferenceEquals(left, right))
        {
            return 0;
        }

        if (left is null)
        {
            return -1;
        }

        if (right is null)
        {
            return 1;
        }

        foreach (var key in _sortState.Keys)
        {
            var result = key.Column switch
            {
                ObjectPickerSortColumn.Group => string.Compare(left.Group, right.Group, StringComparison.OrdinalIgnoreCase),
                ObjectPickerSortColumn.Label => string.Compare(left.Label, right.Label, StringComparison.OrdinalIgnoreCase),
                ObjectPickerSortColumn.ApiName => string.Compare(left.ApiName, right.ApiName, StringComparison.OrdinalIgnoreCase),
                _ => 0,
            };

            if (result == 0)
            {
                continue;
            }

            return key.Direction == GridSortDirection.Ascending ? result : -result;
        }

        return 0;
    }
}
