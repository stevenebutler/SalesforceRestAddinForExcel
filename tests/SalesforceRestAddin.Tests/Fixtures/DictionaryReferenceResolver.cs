using SalesforceRestAddin.Core.Values;

namespace SalesforceRestAddin.Tests.Fixtures;

public sealed class DictionaryReferenceResolver : IReferenceResolver
{
    private readonly Dictionary<(string Object, string Name), string> _nameToId;
    private readonly Dictionary<(string Object, string Id), string> _idToName;

    public DictionaryReferenceResolver(
        Dictionary<(string Object, string Name), string>? nameToId = null,
        Dictionary<(string Object, string Id), string>? idToName = null)
    {
        _nameToId = nameToId ?? new Dictionary<(string, string), string>();
        _idToName = idToName ?? new Dictionary<(string, string), string>();
    }

    public Task<string?> ResolveNameToIdAsync(string targetObject, string displayName, CancellationToken cancellationToken = default)
    {
        _nameToId.TryGetValue((targetObject, displayName), out var id);
        return Task.FromResult<string?>(id);
    }

    public Task<string?> ResolveIdToNameAsync(string targetObject, string id, CancellationToken cancellationToken = default)
    {
        _idToName.TryGetValue((targetObject, id), out var name);
        return Task.FromResult<string?>(name);
    }
}
