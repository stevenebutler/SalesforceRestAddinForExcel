using SalesforceRestAddin.Core.Tables;

namespace SalesforceRestAddin.Tests.Fixtures;

public static class ForceTableSnapshots
{
    public static ForceTableSnapshot ValidAccountTable(int bodyRows = 2)
    {
        var headers = new object?[] { "Account ID", "Account Name", "Industry" };
        var apiNames = new string?[] { "Id", "Name", "Industry" };
        var body = new object?[bodyRows, headers.Length];
        for (var r = 0; r < bodyRows; r++)
        {
            body[r, 0] = $"00100000000000{r:D1}";
            body[r, 1] = $"Acme {r}";
            body[r, 2] = "Technology";
        }

        return new ForceTableSnapshot
        {
            ObjectApiName = "Account",
            CriteriaRow = Array.Empty<object?>(),
            HeaderLabels = headers,
            HeaderApiNames = apiNames,
            Body = body,
            IdColumnIndex = 0,
        };
    }

    public static ForceTableSnapshot ValidAccountTableAt(int startRow, int startColumn, int bodyRows = 2)
    {
        var baseSnapshot = ValidAccountTable(bodyRows);
        return new ForceTableSnapshot
        {
            ObjectApiName = baseSnapshot.ObjectApiName,
            CriteriaRow = baseSnapshot.CriteriaRow,
            HeaderLabels = baseSnapshot.HeaderLabels,
            HeaderApiNames = baseSnapshot.HeaderApiNames,
            Body = baseSnapshot.Body,
            IdColumnIndex = baseSnapshot.IdColumnIndex,
            StartRow = startRow,
            StartColumn = startColumn,
        };
    }

    public static ForceTableSnapshot AccountFromA1Value()
    {
        var snapshot = ValidAccountTable();
        return new ForceTableSnapshot
        {
            ObjectApiName = snapshot.ObjectApiName,
            CriteriaRow = snapshot.CriteriaRow,
            HeaderLabels = snapshot.HeaderLabels,
            HeaderApiNames = snapshot.HeaderApiNames,
            Body = snapshot.Body,
            IdColumnIndex = snapshot.IdColumnIndex,
        };
    }

    public static ForceTableSnapshot MissingIdColumn()
    {
        var headers = new object?[] { "Account Name", "Industry" };
        var apiNames = new string?[] { "Name", "Industry" };
        return new ForceTableSnapshot
        {
            ObjectApiName = "Account",
            CriteriaRow = Array.Empty<object?>(),
            HeaderLabels = headers,
            HeaderApiNames = apiNames,
            Body = new object?[0, headers.Length],
            IdColumnIndex = -1,
        };
    }

    public static ForceTableSnapshot WithCustomFieldComment()
    {
        var headers = new object?[] { "Account ID", "Custom Field" };
        var apiNames = new string?[] { "Id", "Custom__c" };
        return new ForceTableSnapshot
        {
            ObjectApiName = "Account",
            CriteriaRow = Array.Empty<object?>(),
            HeaderLabels = headers,
            HeaderApiNames = apiNames,
            Body = new object?[0, headers.Length],
            IdColumnIndex = 0,
        };
    }

    public static ForceTableSnapshot IdNotFirstColumn()
    {
        var headers = new object?[] { "Account Name", "Account ID", "Industry" };
        var apiNames = new string?[] { "Name", "Id", "Industry" };
        return new ForceTableSnapshot
        {
            ObjectApiName = "Account",
            CriteriaRow = Array.Empty<object?>(),
            HeaderLabels = headers,
            HeaderApiNames = apiNames,
            Body = new object?[0, headers.Length],
            IdColumnIndex = 1,
        };
    }
}
