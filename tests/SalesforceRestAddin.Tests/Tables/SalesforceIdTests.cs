using SalesforceRestAddin.Core.Tables;

namespace SalesforceRestAddin.Tests.Tables;

public sealed class SalesforceIdTests
{
    [Test]
    public async Task Normalize_15Char_Pads_To_18()
    {
        var id = SalesforceId.Normalize("001000000000000");
        await Assert.That(id).IsEqualTo("001000000000000AAA");
    }

    [Test]
    public async Task IsValid_Accepts_15_And_18()
    {
        await Assert.That(SalesforceId.IsValid("001000000000000")).IsTrue();
        await Assert.That(SalesforceId.IsValid("001000000000000AAA")).IsTrue();
        await Assert.That(SalesforceId.IsValid("bad")).IsFalse();
    }
}
