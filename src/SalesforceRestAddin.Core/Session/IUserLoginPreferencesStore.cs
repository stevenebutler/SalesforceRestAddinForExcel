namespace SalesforceRestAddin.Core.Session;

public interface IUserLoginPreferencesStore
{
    /// <summary>
    /// Returns preferences only when the on-disk file is valid JSON with an explicit Production or Sandbox environment.
    /// </summary>
    UserLoginPreferences? TryLoadValid();

    UserLoginPreferences Load();

    void Save(UserLoginPreferences preferences);
}
