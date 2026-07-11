using SalesforceRestAddin.Core.Rest;

namespace SalesforceRestAddin.Core.DataPlane;

public sealed record ObjectPickerRow(
    SObjectSummary Summary,
    string Group,
    string Label,
    string ApiName,
    WizardObjectKind Kind);
