namespace SalesforceRestAddin.Core;

/// <summary>User-facing product names. ProgId remains <see cref="ComApiIdentity.ComAddInProgId"/> for VBA.</summary>
public static class ProductBranding
{
    public const string ProductName = "Salesforce REST Add-in for Excel";

    public const string RibbonTabLabel = "Salesforce Rest";

    /// <summary>Short identifier for AppData folder, deploy bundle, and credential targets.</summary>
    public const string ShortName = "SalesforceRestAddin";

    public const string Tagline = "Bring Salesforce into Excel with REST and native OAuth Login support";

    public const string Author = "Steven Butler";

    public const string Copyright = "Copyright © 2026 Steven Butler";

    /// <summary>Public GitHub repository (develop branch is the active development line).</summary>
    public const string RepositoryUrl = "https://github.com/stevenebutler/SalesforceRestAddinForExcel";

    public const string RepositoryDisplayName = "stevenebutler/SalesforceRestAddinForExcel";

    /// <summary>
    /// Legacy VSTO Force.com Connector Next Generation that inspired the worksheet/ribbon workflow.
    /// The ProgId / ConnectorAdaptor VBA API is separate work (not from that original project).
    /// </summary>
    public const string InspirationProductName = "Force.com Connector Next Generation";

    public const string InspirationRepositoryUrl = "https://github.com/good-ghost/ForceConnector";

    public const string InspirationRepositoryDisplayName = "good-ghost/ForceConnector";
}
