using System;
using System.Collections.Generic;

namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// In-memory credential store for tests and headless scenarios.
/// </summary>
public sealed class InMemorySessionCredentialStore : ISessionCredentialStore
{
    private readonly Dictionary<string, StoredSessionCredentials> _credentials =
        new(StringComparer.OrdinalIgnoreCase);

    public StoredSessionCredentials? Load(string hostKey) =>
        _credentials.TryGetValue(hostKey, out var credentials) ? credentials : null;

    public void Save(StoredSessionCredentials credentials)
    {
        if (credentials is null)
        {
            throw new ArgumentNullException(nameof(credentials));
        }

        _credentials[credentials.HostKey] = credentials;
    }

    public void Delete(string hostKey) => _credentials.Remove(hostKey);
}
