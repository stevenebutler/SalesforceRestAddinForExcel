namespace SalesforceRestAddin.Core.Session;

public interface IConnectorOptionsStore
{
    ConnectorOptions Load();

    void Save(ConnectorOptions options);
}
