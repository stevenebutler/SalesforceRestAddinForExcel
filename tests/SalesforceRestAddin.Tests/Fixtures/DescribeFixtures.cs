using System.Text.Json;
using SalesforceRestAddin.Core.Tables;

namespace SalesforceRestAddin.Tests.Fixtures;

public static class DescribeFixtures
{
    public static SObjectDescribe LoadAccountDescribe()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "AccountDescribe.json");
        if (!File.Exists(path))
        {
            path = Path.Combine(
                Directory.GetCurrentDirectory(),
                "src",
                "SalesforceRestAddin.Tests",
                "Fixtures",
                "AccountDescribe.json");
        }

        var json = File.ReadAllText(path);
        return ParseDescribe(json);
    }

    public static SObjectDescribe ParseDescribe(string json)
    {
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        var fields = new List<FieldDescriptor>();

        foreach (var fieldElement in root.GetProperty("fields").EnumerateArray())
        {
            var picklist = new List<string>();
            if (fieldElement.TryGetProperty("picklistValues", out var picklistElement))
            {
                foreach (var value in picklistElement.EnumerateArray())
                {
                    picklist.Add(value.GetString() ?? string.Empty);
                }
            }

            var referenceTo = new List<string>();
            if (fieldElement.TryGetProperty("referenceTo", out var refElement))
            {
                foreach (var value in refElement.EnumerateArray())
                {
                    referenceTo.Add(value.GetString() ?? string.Empty);
                }
            }

            fields.Add(new FieldDescriptor
            {
                Name = fieldElement.GetProperty("name").GetString() ?? string.Empty,
                Label = fieldElement.GetProperty("label").GetString() ?? string.Empty,
                Type = fieldElement.GetProperty("type").GetString() ?? "string",
                Createable = fieldElement.TryGetProperty("createable", out var c) && c.GetBoolean(),
                Updateable = fieldElement.TryGetProperty("updateable", out var u) && u.GetBoolean(),
                Nillable = !fieldElement.TryGetProperty("nillable", out var n) || n.GetBoolean(),
                Custom = fieldElement.TryGetProperty("custom", out var custom) && custom.GetBoolean(),
                Length = fieldElement.TryGetProperty("length", out var len) && len.TryGetInt32(out var length) ? length : null,
                Precision = fieldElement.TryGetProperty("precision", out var prec) && prec.TryGetInt32(out var precision) ? precision : null,
                Scale = fieldElement.TryGetProperty("scale", out var scaleEl) && scaleEl.TryGetInt32(out var scale) ? scale : null,
                PicklistValues = picklist,
                ReferenceTo = referenceTo,
            });
        }

        return new SObjectDescribe
        {
            Name = root.GetProperty("name").GetString() ?? string.Empty,
            Label = root.GetProperty("label").GetString() ?? string.Empty,
            Fields = fields,
        };
    }
}
