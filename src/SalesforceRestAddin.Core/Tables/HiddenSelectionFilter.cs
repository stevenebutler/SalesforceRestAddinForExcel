namespace SalesforceRestAddin.Core.Tables;

public static class HiddenSelectionFilter
{
    public static bool IsHiddenRow(ForceTableSnapshot snapshot, int bodyRowIndex) =>
        snapshot.HiddenRowIndices?.Contains(bodyRowIndex) == true;

    public static bool IsHiddenColumn(ForceTableSnapshot snapshot, int columnIndex) =>
        snapshot.HiddenColumnIndices?.Contains(columnIndex) == true;

    public static IEnumerable<int> VisibleBodyRows(ForceTableSnapshot snapshot, IEnumerable<int> bodyRowIndices)
    {
        foreach (var row in bodyRowIndices)
        {
            if (!IsHiddenRow(snapshot, row))
            {
                yield return row;
            }
        }
    }

    public static IEnumerable<int> VisibleColumns(ForceTableSnapshot snapshot, int startColumn, int endColumn)
    {
        for (var col = startColumn; col <= endColumn; col++)
        {
            if (!IsHiddenColumn(snapshot, col))
            {
                yield return col;
            }
        }
    }

    public static IEnumerable<int> VisibleColumns(ForceTableSnapshot snapshot, IEnumerable<int> columnIndices)
    {
        foreach (var col in columnIndices)
        {
            if (!IsHiddenColumn(snapshot, col))
            {
                yield return col;
            }
        }
    }
}
