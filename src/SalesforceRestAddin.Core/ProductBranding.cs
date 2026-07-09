namespace SalesforceRestAddin.Core;

/// <summary>User-facing product names. ProgId remains <see cref="ComApiIdentity.ComAddInProgId"/> for VBA.</summary>
public static class ProductBranding
{
    public const string ProductName = "Salesforce REST Add-in for Excel";

    public const string RibbonTabLabel = "Salesforce Rest";

    /// <summary>Short identifier for AppData folder, deploy bundle, and credential targets.</summary>
    public const string ShortName = "SalesforceRestAddin";

    public const string Tagline = "Bring Salesforce into Excel with REST with native OAUTH Login support";
}
