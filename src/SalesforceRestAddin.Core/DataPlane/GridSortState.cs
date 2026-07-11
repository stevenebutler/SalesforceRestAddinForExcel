namespace SalesforceRestAddin.Core.DataPlane;

public enum GridSortDirection
{
    Ascending = 0,
    Descending = 1,
}

public sealed record GridSortKey<TColumn>(TColumn Column, GridSortDirection Direction);

public sealed class GridSortState<TColumn> where TColumn : notnull
{
    private readonly List<GridSortKey<TColumn>> _keys = new();
    private readonly List<GridSortKey<TColumn>> _defaultKeys = new();
    private readonly IEqualityComparer<TColumn> _comparer;

    public GridSortState(IEnumerable<GridSortKey<TColumn>> defaultKeys, IEqualityComparer<TColumn>? comparer = null)
    {
        _comparer = comparer ?? EqualityComparer<TColumn>.Default;
        _defaultKeys.AddRange(defaultKeys);
        _keys.AddRange(_defaultKeys);
    }

    public IReadOnlyList<GridSortKey<TColumn>> Keys => _keys;

    public void ResetDefault()
    {
        _keys.Clear();
        _keys.AddRange(_defaultKeys);
    }

    public void Click(TColumn column)
    {
        var existingIndex = _keys.FindIndex(key => _comparer.Equals(key.Column, column));
        if (existingIndex == 0)
        {
            var current = _keys[0];
            _keys[0] = current with
            {
                Direction = current.Direction == GridSortDirection.Ascending
                    ? GridSortDirection.Descending
                    : GridSortDirection.Ascending,
            };
            return;
        }

        if (existingIndex > 0)
        {
            _keys.RemoveAt(existingIndex);
        }

        _keys.Insert(0, new GridSortKey<TColumn>(column, GridSortDirection.Ascending));
    }
}
