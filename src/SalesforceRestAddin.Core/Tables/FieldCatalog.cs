namespace SalesforceRestAddin.Core.Tables;

public sealed class FieldCatalog
{
    private readonly Dictionary<string, FieldDescriptor> _byApiName;
    private readonly Dictionary<string, FieldDescriptor> _byLabel;

    public FieldCatalog(SObjectDescribe describe)
    {
        if (describe is null)
        {
            throw new ArgumentNullException(nameof(describe));
        }

        _byApiName = new Dictionary<string, FieldDescriptor>(StringComparer.OrdinalIgnoreCase);
        _byLabel = new Dictionary<string, FieldDescriptor>(StringComparer.OrdinalIgnoreCase);

        foreach (var field in describe.Fields)
        {
            _byApiName[field.Name] = field;
            if (!_byLabel.ContainsKey(field.Label))
            {
                _byLabel[field.Label] = field;
            }
        }

        Describe = describe;
    }

    public SObjectDescribe Describe { get; }

    public bool TryGetByApiName(string apiName, out FieldDescriptor? field) =>
        _byApiName.TryGetValue(apiName, out field);

    public bool TryGetByLabel(string label, out FieldDescriptor? field) =>
        _byLabel.TryGetValue(label, out field);

    public FieldDescriptor? ResolveHeaderField(string? label, string? apiNameFromComment)
    {
        if (apiNameFromComment is not null
            && !string.IsNullOrWhiteSpace(apiNameFromComment)
            && TryGetByApiName(apiNameFromComment, out var byApi))
        {
            return byApi;
        }

        if (label is not null && !string.IsNullOrWhiteSpace(label) && TryGetByLabel(label.Trim(), out var byLabel))
        {
            return byLabel;
        }

        if (label is not null && !string.IsNullOrWhiteSpace(label) && TryGetByApiName(label.Trim(), out var byApiName))
        {
            return byApiName;
        }

        return null;
    }
}
