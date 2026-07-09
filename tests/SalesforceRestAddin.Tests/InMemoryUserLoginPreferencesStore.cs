using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Tests;

public sealed class InMemoryUserLoginPreferencesStore : IUserLoginPreferencesStore
{
    public UserLoginPreferences? Stored { get; set; }

    public UserLoginPreferences? TryLoadValid() => Stored;

    public UserLoginPreferences Load() => Stored ?? new UserLoginPreferences();

    public void Save(UserLoginPreferences preferences) => Stored = preferences;
}
