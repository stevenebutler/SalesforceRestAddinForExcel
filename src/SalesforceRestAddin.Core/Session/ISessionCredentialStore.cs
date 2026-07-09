namespace SalesforceRestAddin.Core.Session;

public interface ISessionCredentialStore
{
    StoredSessionCredentials? Load(string hostKey);

    void Save(StoredSessionCredentials credentials);

    void Delete(string hostKey);
}
