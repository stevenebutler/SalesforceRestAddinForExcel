using System;
using System.Collections.Generic;
using System.Linq;
using SalesforceRestAddin.Core.DataPlane;

namespace SalesforceRestAddin.Windows.Ui;

internal enum FieldPickerSortColumn
{
    Category = 0,
    Label = 1,
    ApiName = 2,
}

internal sealed class FieldPickerSortState
{
    private readonly GridSortState<FieldPickerSortColumn> _sortState;

    public FieldPickerSortState()
    {
        _sortState = new GridSortState<FieldPickerSortColumn>(
            [
                new GridSortKey<FieldPickerSortColumn>(FieldPickerSortColumn.Category, GridSortDirection.Ascending),
            ]);
    }

    public IReadOnlyList<GridSortKey<FieldPickerSortColumn>> Keys => _sortState.Keys;

    public void Promote(FieldPickerSortColumn column) => _sortState.Click(column);

    public IReadOnlyList<FieldListItem> Sort(IEnumerable<FieldListItem> rows) =>
        rows.OrderBy(row => row, Comparer<FieldListItem>.Create(Compare)).ToList();

    private int Compare(FieldListItem? left, FieldListItem? right)
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
                FieldPickerSortColumn.Category => left.Bucket.CompareTo(right.Bucket),
                FieldPickerSortColumn.Label => string.Compare(left.LabelText, right.LabelText, StringComparison.OrdinalIgnoreCase),
                FieldPickerSortColumn.ApiName => string.Compare(left.ApiName, right.ApiName, StringComparison.OrdinalIgnoreCase),
                _ => 0,
            };

            if (result == 0)
            {
                continue;
            }

            return key.Direction == GridSortDirection.Ascending ? result : -result;
        }

        return left.DescribeIndex.CompareTo(right.DescribeIndex);
    }
}
