using System;
using Microsoft.Office.Interop.Excel;
using SalesforceRestAddin.Core;

namespace SalesforceRestAddin;

internal static class WorkbookContextGuard
{
    internal static bool TryRequireContext(
        Application excel,
        WorkbookContextRequirement requirement)
    {
        if (excel is null)
        {
            return false;
        }

        try
        {
            if (excel.Workbooks.Count == 0 || excel.ActiveWorkbook is null)
            {
                return false;
            }

            if (requirement == WorkbookContextRequirement.Workbook)
            {
                return true;
            }

            return excel.ActiveSheet is Worksheet;
        }
        catch
        {
            return false;
        }
    }

    internal static string GetMessage(WorkbookContextRequirement requirement) =>
        WorkbookContextRequirementText.GetMessage(requirement);
}
