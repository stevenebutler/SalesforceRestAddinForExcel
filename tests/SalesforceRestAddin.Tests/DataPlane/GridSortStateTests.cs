using SalesforceRestAddin.Core.DataPlane;

namespace SalesforceRestAddin.Tests.DataPlane;

public sealed class GridSortStateTests
{
    private enum SampleColumn
    {
        First = 0,
        Second = 1,
        Third = 2,
    }

    [Test]
    public async Task T_GRID_SORT_01_Promoting_A_New_Column_Moves_It_To_The_Front_Ascending()
    {
        var state = new GridSortState<SampleColumn>(
            [
                new GridSortKey<SampleColumn>(SampleColumn.First, GridSortDirection.Ascending),
                new GridSortKey<SampleColumn>(SampleColumn.Second, GridSortDirection.Ascending),
            ]);

        state.Click(SampleColumn.Third);

        await Assert.That(state.Keys.SequenceEqual(
            [
                new GridSortKey<SampleColumn>(SampleColumn.Third, GridSortDirection.Ascending),
                new GridSortKey<SampleColumn>(SampleColumn.First, GridSortDirection.Ascending),
                new GridSortKey<SampleColumn>(SampleColumn.Second, GridSortDirection.Ascending),
            ])).IsTrue();
    }

    [Test]
    public async Task T_GRID_SORT_02_Clicking_The_Current_First_Column_Toggles_Only_Its_Direction()
    {
        var state = new GridSortState<SampleColumn>(
            [
                new GridSortKey<SampleColumn>(SampleColumn.First, GridSortDirection.Ascending),
                new GridSortKey<SampleColumn>(SampleColumn.Second, GridSortDirection.Ascending),
            ]);

        state.Click(SampleColumn.First);

        await Assert.That(state.Keys.SequenceEqual(
            [
                new GridSortKey<SampleColumn>(SampleColumn.First, GridSortDirection.Descending),
                new GridSortKey<SampleColumn>(SampleColumn.Second, GridSortDirection.Ascending),
            ])).IsTrue();
    }

    [Test]
    public async Task T_GRID_SORT_03_Promoting_A_Non_First_Existing_Column_Preserves_Remaining_Order()
    {
        var state = new GridSortState<SampleColumn>(
            [
                new GridSortKey<SampleColumn>(SampleColumn.First, GridSortDirection.Ascending),
                new GridSortKey<SampleColumn>(SampleColumn.Second, GridSortDirection.Ascending),
                new GridSortKey<SampleColumn>(SampleColumn.Third, GridSortDirection.Ascending),
            ]);

        state.Click(SampleColumn.Second);

        await Assert.That(state.Keys.SequenceEqual(
            [
                new GridSortKey<SampleColumn>(SampleColumn.Second, GridSortDirection.Ascending),
                new GridSortKey<SampleColumn>(SampleColumn.First, GridSortDirection.Ascending),
                new GridSortKey<SampleColumn>(SampleColumn.Third, GridSortDirection.Ascending),
            ])).IsTrue();
    }
}
