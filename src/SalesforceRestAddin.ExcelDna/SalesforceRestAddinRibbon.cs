using System;
using System.IO;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SalesforceRestAddin.Core;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin;

/// <summary>
/// Ribbon callbacks — registered by Excel-DNA <c>LoadCustomUI</c> after <see cref="ExcelAddIn.AutoOpen"/>.
/// </summary>
[ComVisible(true)]
public sealed class SalesforceRestAddinRibbon : ExcelRibbon
{
    public const string ConnectorGroupControlId = "grpSalesforceRest";
    public const string LogoutControlId = "btnLogout";

    private static IRibbonUI? _ribbon;
    private static IUserLoginPreferencesStore? _preferencesStore;

    public override string GetCustomUI(string ribbonId) =>
        ribbonId == "Microsoft.Excel.Workbook"
            ? LoadRibbonXml()
            : string.Empty;

    private static string LoadRibbonXml()
    {
        const string resourceName = "SalesforceRestAddin.SalesforceRestAddinRibbon.xml";
        using var stream = typeof(SalesforceRestAddinRibbon).Assembly.GetManifestResourceStream(resourceName);
        if (stream is null)
        {
            throw new InvalidOperationException($"Embedded ribbon XML not found: {resourceName}");
        }

        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    public void OnRibbonLoad(IRibbonUI ribbon) => _ribbon = ribbon;

    internal static void Configure(IUserLoginPreferencesStore preferencesStore) =>
        _preferencesStore = preferencesStore ?? throw new ArgumentNullException(nameof(preferencesStore));

    /// <summary>Invalidates group label and Logout enablement after session/target changes.</summary>
    internal static void RefreshSessionDependentUi()
    {
        _ribbon?.InvalidateControl(ConnectorGroupControlId);
        _ribbon?.InvalidateControl(LogoutControlId);
    }

    private static SessionUiState CurrentUiState()
    {
        var prefs = _preferencesStore
            ?? new JsonUserLoginPreferencesStore(SalesforceRestAddinDataPaths.PreferencesFile);
        return SessionUiState.From(prefs, SessionContext.Current);
    }

    public string GetConnectorGroupLabel(IRibbonControl control) => CurrentUiState().GroupLabel;

    public bool GetOptionsEnabled(IRibbonControl control) => true;

    public bool GetLogoutEnabled(IRibbonControl control) => CurrentUiState().LogoutEnabled;

    public void OnAbout(IRibbonControl control) =>
        ExcelAsyncUtil.QueueAsMacro(AddInHost.ShowAbout);

    public void OnTableWizard(IRibbonControl control) =>
        ExcelAsyncUtil.QueueAsMacro(AddInHost.QueryTableWizard);

    public void OnUpdateSelectedCells(IRibbonControl control) =>
        ExcelAsyncUtil.QueueAsMacro(AddInHost.UpdateSelectedCells);

    public void OnInsertSelectedRows(IRibbonControl control) =>
        ExcelAsyncUtil.QueueAsMacro(AddInHost.InsertSelectedRows);

    public void OnQuerySelectedRows(IRibbonControl control) =>
        ExcelAsyncUtil.QueueAsMacro(AddInHost.QuerySelectedRows);

    public void OnQueryTableData(IRibbonControl control) =>
        ExcelAsyncUtil.QueueAsMacro(AddInHost.QueryTableData);

    public void OnDeleteSelectedRecords(IRibbonControl control) =>
        ExcelAsyncUtil.QueueAsMacro(AddInHost.DeleteSelectedRecords);

    public void OnDescribeSforceObject(IRibbonControl control) =>
        ExcelAsyncUtil.QueueAsMacro(AddInHost.DescribeSforceObject);

    public void OnOptions(IRibbonControl control) =>
        ExcelAsyncUtil.QueueAsMacro(AddInHost.ShowOptions);

    public void OnLogout(IRibbonControl control) =>
        ExcelAsyncUtil.QueueAsMacro(AddInHost.Logout);
}
