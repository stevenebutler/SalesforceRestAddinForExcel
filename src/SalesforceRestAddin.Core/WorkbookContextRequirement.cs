using System;

namespace SalesforceRestAddin.Core;

internal enum WorkbookContextRequirement
{
    Workbook,
    Worksheet,
    SalesforceConnectorTable,
}

internal static class WorkbookContextRequirementText
{
    internal static string GetMessage(WorkbookContextRequirement requirement) =>
        requirement switch
        {
            WorkbookContextRequirement.Workbook => "Create or open a workbook, then try again.",
            WorkbookContextRequirement.Worksheet => "Create or open a workbook, then try again.",
            WorkbookContextRequirement.SalesforceConnectorTable => "Open a workbook containing a Salesforce Connector table, then try again.",
            _ => throw new ArgumentOutOfRangeException(nameof(requirement), requirement, null),
        };
}
