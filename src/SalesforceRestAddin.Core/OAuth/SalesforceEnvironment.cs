namespace SalesforceRestAddin.Core.OAuth;

public enum SalesforceEnvironment
{
    Production,
    Sandbox,
}

public static class SalesforceEnvironmentExtensions
{
    public static string LoginHost(this SalesforceEnvironment environment) =>
        environment switch
        {
            SalesforceEnvironment.Production => "login.salesforce.com",
            SalesforceEnvironment.Sandbox => "test.salesforce.com",
            _ => throw new ArgumentOutOfRangeException(nameof(environment)),
        };
}
