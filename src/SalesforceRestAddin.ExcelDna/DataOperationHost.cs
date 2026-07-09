using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using SalesforceRestAddin.Core;
using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Rest;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Core.Soql;
using SalesforceRestAddin.Core.Tables;
using SalesforceRestAddin.Tables;
using SalesforceRestAddin.Windows.Ui;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;
using ExcelWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
using QueryTableOperation = SalesforceRestAddin.Core.DataPlane.QueryTable;

namespace SalesforceRestAddin;

public static class DataOperationHost
{
    private static readonly object GateLock = new();
    private static SessionGateComposer.SessionServices? _services;
    private static IConnectorOptionsStore? _optionsStore;

    public static void Configure(
        SessionGateComposer.SessionServices? services = null,
        IConnectorOptionsStore? optionsStore = null)
    {
        lock (GateLock)
        {
            _services = services;
            _optionsStore = optionsStore;
        }
    }

    private static SessionGateComposer.SessionServices Services =>
        _services ?? SessionGateComposer.Create();

    private static ConnectorOptions LoadOptions() =>
        (_optionsStore ?? new JsonConnectorOptionsStore(SalesforceRestAddinDataPaths.ConnectorOptionsFile)).Load();

    public static Task RunQueryTableDataAsync()
    {
        RunQueryTableData();
        return Task.CompletedTask;
    }

    public static void RunQueryTableData()
    {
        SessionFlowTrace.BeginScope("Query Table Data");

        var services = Services;
        AddInHost.EnsureLoggedInAndRefreshUi(services.Gate);

        var session = SessionContext.Current;
        SessionFlowTrace.Log(
            $"Session instance={session.InstanceUrl} apiVersion={session.ApiVersion} object from active cell");

        var excel = (ExcelApplication)ExcelDnaUtil.Application;
        var sheet = (ExcelWorksheet)excel.ActiveSheet;
        var activeCell = (ExcelRange)excel.ActiveCell;
        ForceTableSnapshot snapshot;
        try
        {
            snapshot = ForceTableReader.Capture(excel, sheet, activeCell);
        }
        catch (InvalidOperationException ex)
        {
            ErrorDialogWindow.Show(AddInHost.AddInTitle, ex.Message);
            return;
        }

        var options = LoadOptions();
        var client = services.CreateDataClient();
        var input = new QueryTableInput
        {
            Snapshot = snapshot,
            Options = options,
        };

        if (options.ConfirmLargeQuery)
        {
            var countResult = ExcelStaAsyncHost.Run(
                "Counting matches…",
                excel,
                (ct, report) =>
                {
                    report("Counting matching rows…");
                    return QueryTableOperation.GetMatchCountAsync(client, input, ct);
                });
            var countError = countResult.ErrorSummary;
            if (countError is not null && countError.Length > 0)
            {
                ErrorDialogWindow.Show(AddInHost.AddInTitle, countError);
                return;
            }

            if (countResult.Count is int count)
            {
                if (count > SoqlQueryBuilder.ExcelRowLimit)
                {
                    ErrorDialogWindow.Show(
                        AddInHost.AddInTitle,
                        $"Query would return {count} rows, exceeding the Excel limit.");
                    return;
                }

                if (!ConfirmationDialogWindow.Show(
                        AddInHost.AddInTitle,
                        $"Download {count} rows from Salesforce?"))
                {
                    return;
                }
            }
        }

        DataOperationResult result;
        try
        {
            result = ExcelStaAsyncHost.Run(
                "Query Table Data",
                excel,
                (ct, report) =>
                {
                    report(OperationProgress.Status("Querying Salesforce..."));
                    return QueryTableOperation.RunAsync(
                        client,
                        new QueryTableInput
                        {
                            Snapshot = snapshot,
                            Options = options,
                            ConfirmLargeQuery = options.ConfirmLargeQuery,
                            Progress = report,
                        },
                        ct);
                });
        }
        catch (OperationCanceledException)
        {
            // Cancelled mid-download: leave the sheet unchanged (bugs.md #4).
            SessionFlowTrace.Log("Query Table Data: cancelled by user.");
            return;
        }

        ApplyResult(sheet, result, binding: null, excel);
    }

    public static void ApplyResult(
        ExcelWorksheet sheet,
        DataOperationResult result,
        ForceTableBinding? binding = null,
        ExcelApplication? excel = null)
    {
        try
        {
            if (result.ClearBody is not null)
            {
                SetStatusBar(excel, "Clearing cells");
                SheetProjectionWriter.ClearBody(sheet, result.ClearBody);
            }

            if (result.Projection is not null)
            {
                SetStatusBar(excel, "Updating cells");
                SheetProjectionWriter.Apply(sheet, result.Projection);
            }

            if (binding is not null && result.RowOutcomes.Count > 0)
            {
                SetStatusBar(excel, "Writing results");
                RowOutcomeWriter.Apply(sheet, binding, result.RowOutcomes);
            }
        }
        finally
        {
            ClearStatusBar(excel);
        }

        if (!string.IsNullOrEmpty(result.ErrorSummary))
        {
            var details = result.ErrorSummary
                + Environment.NewLine
                + Environment.NewLine
                + SessionFlowTrace.FormatRecent();
            ErrorDialogWindow.Show(AddInHost.AddInTitle, details);
        }
    }

    private static void SetStatusBar(ExcelApplication? excel, string message)
    {
        if (excel is null)
        {
            return;
        }

        try
        {
            var truncated = message.Length > 128 ? message.Substring(0, 128) : message;
            excel.StatusBar = truncated;
        }
        catch
        {
            // Excel may reject status bar updates during shutdown.
        }
    }

    private static void ClearStatusBar(ExcelApplication? excel)
    {
        if (excel is null)
        {
            return;
        }

        try
        {
            excel.StatusBar = false;
        }
        catch
        {
            // Ignore status bar reset failures.
        }
    }
}
