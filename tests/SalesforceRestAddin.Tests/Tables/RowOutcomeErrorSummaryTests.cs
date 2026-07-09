using SalesforceRestAddin.Core.Tables;

namespace SalesforceRestAddin.Tests.Tables;

public sealed class RowOutcomeErrorSummaryTests
{
    [Test]
    public async Task FromFailures_Uses_RecordsProcessed_When_Only_Failures_Are_Stored()
    {
        var outcomes = new[]
        {
            new RowOutcome
            {
                BodyRowIndex = 1,
                Succeeded = false,
                ErrorMessages = new[] { "FIELD_CUSTOM_VALIDATION_EXCEPTION: bad value" },
            },
        };

        var summary = RowOutcomeErrorSummary.FromFailures(outcomes, tableStartRow: 1, "Update", recordsProcessed: 5);

        await Assert.That(summary).IsNotNull();
        await Assert.That(summary!).Contains("Update failed for 1 of 5 row(s)");
        await Assert.That(summary!).Contains("Row 4:");
        await Assert.That(summary!).Contains("bad value");
    }

    [Test]
    public async Task FromFailures_Returns_Null_When_All_Succeeded()
    {
        var outcomes = new[]
        {
            new RowOutcome { BodyRowIndex = 0, Succeeded = true },
        };

        await Assert.That(RowOutcomeErrorSummary.FromFailures(outcomes, 1, "Insert", 1)).IsNull();
    }
}
