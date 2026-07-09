using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Tests.Fixtures;

namespace SalesforceRestAddin.Tests.DataPlane;

public sealed class DescribeObjectTests
{
    [Test]
    public async Task T_DSO_02_Field_Ordering_Standard_Before_Custom()
    {
        var describe = DescribeFixtures.LoadAccountDescribe();
        var rows = DescribeObject.BuildRows(describe);
        var names = rows.Select(r => r[1]?.ToString()).ToList();
        await Assert.That(names[0]).IsEqualTo("Id");
        await Assert.That(names.Any(n => n == "Custom__c")).IsTrue();
        await Assert.That(names.IndexOf("Custom__c")).IsGreaterThan(names.IndexOf("Name"));
    }

    [Test]
    public async Task T_DSO_01_BuildProjection_Has_Header_Row()
    {
        var describe = DescribeFixtures.LoadAccountDescribe();
        var projection = DescribeObject.BuildProjection(describe);
        await Assert.That(projection.Values[0, 0]).IsEqualTo("Account");
        await Assert.That(projection.Values[1, 0]).IsEqualTo("Label");
    }
}
