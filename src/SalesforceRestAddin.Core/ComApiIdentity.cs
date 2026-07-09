namespace SalesforceRestAddin.Core;

/// <summary>
/// Well-known COM identity values for the VBA automation API.
/// Must remain stable for existing ConnectorAdaptor modules.
/// </summary>
public static class ComApiIdentity
{
    public const string AutomationClassId = "e5fcccc8-a685-4980-a79a-eab37f2c7caf";
    public const string AutomationInterfaceId = "bacb17f7-9b85-4137-861d-c9a2d899b564";
    public const string AutomationEventsId = "6894737e-cfa7-410c-be80-03ad4de47656";

    /// <summary>COMAddIns registry key / ProgId for the Excel add-in host.</summary>
    public const string ComAddInProgId = "ForceConnector.NextGen";

    /// <summary>Fixed GUID for the Excel-DNA COM add-in registration (VSTO project GUID).</summary>
    public const string ComAddInClassId = "051C9DB6-A956-01BF-02FF-0050C399499D";
}
