namespace SalesforceRestAddin.Core.Values;

public interface IReferenceResolver
{
    Task<string?> ResolveNameToIdAsync(string targetObject, string displayName, CancellationToken cancellationToken = default);

    Task<string?> ResolveIdToNameAsync(string targetObject, string id, CancellationToken cancellationToken = default);
}
