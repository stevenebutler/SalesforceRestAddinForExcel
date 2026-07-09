using System.Runtime.InteropServices;
using SalesforceRestAddin.Core;
using Microsoft.VisualBasic;

namespace SalesforceRestAddin;

[ComVisible(true)]
[ComClass(ComApiIdentity.AutomationClassId, ComApiIdentity.AutomationInterfaceId, ComApiIdentity.AutomationEventsId)]
public class SalesForceAddInApi
{
    public SalesForceAddInApi()
    {
    }

    public void QuerySelectedRowsApi() => AddInHost.QuerySelectedRows();

    public void QueryTableDataApi() => AddInHost.QueryTableData();

    public void UpdateSelectedCellsApi() => AddInHost.UpdateSelectedCells();

    public void InsertSelectedRowsApi() => AddInHost.InsertSelectedRows();

    public void DeleteSelectedRecordsApi() => AddInHost.DeleteSelectedRecords();

    public void RefreshTableDataApi() => AddInHost.RefreshTableData();
}
