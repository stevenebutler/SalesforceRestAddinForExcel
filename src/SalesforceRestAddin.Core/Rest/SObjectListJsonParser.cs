using System.Text.Json;

namespace SalesforceRestAddin.Core.Rest;

public static class SObjectListJsonParser
{
    public static IReadOnlyList<SObjectSummary> Parse(string json)
    {
        using var document = JsonDocument.Parse(json);
        var list = new List<SObjectSummary>();
        foreach (var element in document.RootElement.GetProperty("sobjects").EnumerateArray())
        {
            list.Add(new SObjectSummary
            {
                Name = element.GetProperty("name").GetString() ?? string.Empty,
                Label = element.GetProperty("label").GetString() ?? string.Empty,
                Queryable = element.TryGetProperty("queryable", out var q) && q.GetBoolean(),
                Custom = element.TryGetProperty("custom", out var c) && c.GetBoolean(),
                CustomSetting = element.TryGetProperty("customSetting", out var cs) && cs.GetBoolean(),
                DeprecatedAndHidden = element.TryGetProperty("deprecatedAndHidden", out var d) && d.GetBoolean(),
                AssociateEntityType = element.TryGetProperty("associateEntityType", out var aet) ? aet.GetString() : null,
                AssociateParentEntity = element.TryGetProperty("associateParentEntity", out var ape) ? ape.GetString() : null,
                KeyPrefix = element.TryGetProperty("keyPrefix", out var kp) ? kp.GetString() : null,
            });
        }

        return list;
    }
}
