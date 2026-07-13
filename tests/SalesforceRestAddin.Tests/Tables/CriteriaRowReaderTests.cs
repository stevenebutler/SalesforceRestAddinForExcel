using SalesforceRestAddin.Core.Tables;
using SalesforceRestAddin.Tests.Fixtures;

namespace SalesforceRestAddin.Tests.Tables;

public sealed class CriteriaRowReaderTests
{
    private static readonly FieldCatalog Catalog = new(DescribeFixtures.LoadAccountDescribe());

    [Test]
    public async Task T_TBL_CRIT_01_Expands_In_Four_Triplet_Chunks_Until_Blank_Terminator()
    {
        var criteria = BuildCriteriaTriplets(("Name", "equals", "A"), ("Industry", "equals", "B"), ("Id", "equals", "C"), ("Name", "equals", "D"), ("Industry", "equals", "E"));
        var calls = new List<(int StartColumn, int Width)>();

        var result = CriteriaRowReader.Read(
            7,
            2,
            (startColumn, width) =>
            {
                calls.Add((startColumn, width));
                return Slice(criteria, startColumn, width);
            },
            Catalog);

        await Assert.That(result.Succeeded).IsTrue();
        await Assert.That(result.CriteriaRow.Length).IsEqualTo(15);
        await Assert.That(calls.Count).IsEqualTo(2);
        await Assert.That(calls[0].StartColumn).IsEqualTo(2);
        await Assert.That(calls[0].Width).IsEqualTo(13);
        await Assert.That(calls[1].StartColumn).IsEqualTo(14);
        await Assert.That(calls[1].Width).IsEqualTo(13);
    }

    [Test]
    public async Task T_TBL_CRIT_02_Bad_Field_Stops_Without_Reading_Later_Chunks()
    {
        var criteria = BuildCriteriaTriplets(
            ("Name", "equals", "A"),
            ("Industry", "equals", "B"),
            ("Id", "equals", "C"),
            ("Name", "equals", "D"),
            ("Bogus Field", "equals", "E"),
            ("Industry", "equals", "F"));
        var calls = new List<(int StartColumn, int Width)>();

        var result = CriteriaRowReader.Read(
            7,
            2,
            (startColumn, width) =>
            {
                calls.Add((startColumn, width));
                return Slice(criteria, startColumn, width);
            },
            Catalog);

        await Assert.That(result.Succeeded).IsFalse();
        await Assert.That(result.Error!.Cell.Row).IsEqualTo(7);
        await Assert.That(result.Error!.Cell.Column).IsEqualTo(14);
        await Assert.That(result.Error.Message).Contains("Bogus Field");
        await Assert.That(calls.Count).IsEqualTo(2);
    }

    private static object?[] BuildCriteriaTriplets(params (string Field, string Operator, string Value)[] triplets)
    {
        var values = new List<object?>();
        foreach (var triplet in triplets)
        {
            values.Add(triplet.Field);
            values.Add(triplet.Operator);
            values.Add(triplet.Value);
        }

        values.Add(null);
        return values.ToArray();
    }

    private static object?[] Slice(object?[] criteria, int startColumn, int width)
    {
        var slice = new object?[width];
        var offset = startColumn - 2;
        for (var i = 0; i < width; i++)
        {
            var index = offset + i;
            slice[i] = index >= 0 && index < criteria.Length ? criteria[index] : null;
        }

        return slice;
    }
}
