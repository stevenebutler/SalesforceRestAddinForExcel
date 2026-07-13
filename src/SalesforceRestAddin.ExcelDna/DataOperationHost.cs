using System;
using System.Diagnostics;
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
    // One queued response overlaps the current Excel write with the next queryMore request
    // while keeping memory use and cancellation latency bounded. Raise only from measured data.
    private const int QueryPageQueueCapacity = 1;

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
        var excel = (ExcelApplication)ExcelDnaUtil.Application;
        if (!WorkbookContextGuard.TryRequireContext(excel, WorkbookContextRequirement.SalesforceConnectorTable))
        {
            ErrorDialogWindow.Show(
                AddInHost.AddInTitle,
                WorkbookContextGuard.GetMessage(WorkbookContextRequirement.SalesforceConnectorTable));
            return;
        }

        AddInHost.EnsureLoggedInAndRefreshUi(services.Gate);

        var sheet = (ExcelWorksheet)excel.ActiveSheet;
        var activeCell = (ExcelRange)excel.ActiveCell;
        ForceTableSnapshot shell;
        try
        {
            shell = ForceTableReader.CaptureShell(excel, sheet, activeCell);
        }
        catch (InvalidOperationException ex)
        {
            ErrorDialogWindow.Show(AddInHost.AddInTitle, ex.Message);
            return;
        }

        var client = services.CreateDataClient();
        var describe = ExcelStaAsyncHost.Run(
            "Loading metadata…",
            excel,
            (ct, report) =>
            {
                report($"Describing {shell.ObjectApiName}…");
                return client.DescribeAsync(shell.ObjectApiName, ct);
            });
        var catalog = new FieldCatalog(describe);
        var criteriaRead = CriteriaRowReader.Read(
            shell.StartRow,
            shell.StartColumn + 1,
            (startColumn, width) => ForceTableReader.ReadCriteriaRowSlice(sheet, shell.StartRow, startColumn, width),
            catalog);
        if (!criteriaRead.Succeeded)
        {
            ErrorDialogWindow.Show(AddInHost.AddInTitle, criteriaRead.Error!.Message);
            return;
        }

        var criteriaReferenceIds = ForceTableReader.ReadCriteriaReferenceIds(sheet, criteriaRead.CriteriaRow);
        var options = LoadOptions();
        var session = SessionContext.Current;
        SessionFlowTrace.Log(
            $"Session instance={session.InstanceUrl} apiVersion={session.ApiVersion} object from active cell");

        var input = new QueryTableInput
        {
            Snapshot = new ForceTableSnapshot
            {
                ObjectApiName = shell.ObjectApiName,
                CriteriaRow = criteriaRead.CriteriaRow,
                CriteriaReferenceIds = criteriaReferenceIds,
                HeaderLabels = shell.HeaderLabels,
                HeaderApiNames = shell.HeaderApiNames,
                Body = shell.Body,
                StartRow = shell.StartRow,
                StartColumn = shell.StartColumn,
                HiddenRowIndices = shell.HiddenRowIndices,
                HiddenColumnIndices = shell.HiddenColumnIndices,
            },
            Options = options,
            Describe = describe,
        };

        if (!options.NoConfirmQueryDownload)
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
        TimeSpan apiElapsed;
        var pageWrites = new PagedQueryWriteState(options.ColumnSizingMode, options.RowSizingMode);
        var queryStopwatch = Stopwatch.StartNew();
        try
        {
            var pagedResult = ExcelStaAsyncHost.Run(
                "Query Table Data",
                excel,
                (ct, report) =>
                {
                    report(OperationProgress.Status("Querying Salesforce..."));
                    return RunPagedQueryAsync(
                        client,
                        new QueryTableInput
                        {
                            Snapshot = input.Snapshot,
                            Options = options,
                            ConfirmQueryTableDownload = !options.NoConfirmQueryDownload,
                            Progress = report,
                        },
                        sheet,
                        excel,
                        pageWrites,
                        ct);
                });
            result = pagedResult.Result;
            apiElapsed = pagedResult.ApiElapsed;
        }
        catch (OperationCanceledException)
        {
            // Cancelled mid-download: leave the sheet unchanged (bugs.md #4).
            SessionFlowTrace.Log("Query Table Data: cancelled by user.");
            return;
        }
        finally
        {
            queryStopwatch.Stop();
        }

        var applyStopwatch = Stopwatch.StartNew();
        try
        {
            // A streamed success has already been applied page-by-page. Non-streamed
            // reference joins and terminal error markers retain the existing result path.
            if (!pageWrites.HasPages || result.ErrorSummary is not null)
            {
                ApplyResult(sheet, result, options, binding: null, excel);
            }
        }
        finally
        {
            applyStopwatch.Stop();
        }

        if (result.Succeeded)
        {
            SetStatusBar(
                excel,
                QueryCompletionStatus.Build(
                    result.RecordsProcessed,
                    queryStopwatch.Elapsed,
                    apiElapsed,
                    pageWrites.ApplyElapsed + applyStopwatch.Elapsed));
        }
    }

    private static async Task<PagedQueryResult> RunPagedQueryAsync(
        SalesforceDataClient client,
        QueryTableInput input,
        ExcelWorksheet sheet,
        ExcelApplication excel,
        PagedQueryWriteState pageWrites,
        CancellationToken cancellationToken)
    {
        var queue = new BoundedAsyncQueue<QueryTablePage>(QueryPageQueueCapacity);
        // Do not let SemaphoreSlim release the consumer inline on the REST producer.
        // The consumer blocks only its worker while ExcelComThread marshals COM to the STA.
        var consumer = Task.Run(() => ConsumeQueryPagesAsync(queue, sheet, excel, pageWrites));
        PagedQueryResult result;
        try
        {
            result = await QueryTableOperation.RunPagedAsync(
                client,
                input,
                (page, ct) => queue.EnqueueAsync(page, ct),
                cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            queue.Complete();
            // Always drain pages that were accepted before cancellation or a terminal
            // producer failure. The consumer intentionally has no cancellation token.
            await consumer.ConfigureAwait(false);
        }

        return result;
    }

    private static async Task ConsumeQueryPagesAsync(
        BoundedAsyncQueue<QueryTablePage> queue,
        ExcelWorksheet sheet,
        ExcelApplication excel,
        PagedQueryWriteState pageWrites)
    {
        while (true)
        {
            var next = await queue.TryDequeueAsync().ConfigureAwait(false);
            if (!next.HasItem)
            {
                return;
            }

            var page = next.Item!;
            var writeElapsed = TimeSpan.Zero;
            ExcelComThread.Invoke(() =>
            {
                SessionFlowTrace.Log(
                    $"QueryTable page {page.PageNumber}: Excel write start " +
                    $"records={page.Projection.Values.GetLength(0)}");
                var writeStopwatch = Stopwatch.StartNew();
                try
                {
                    if (page.ClearBody is not null)
                    {
                        SetStatusBar(excel, "Clearing cells");
                        SheetProjectionWriter.ClearBody(sheet, page.ClearBody);
                    }

                    if (page.TotalRowCount is int totalRowCount)
                    {
                        SetStatusBar(excel, "Formatting cells");
                        SheetProjectionWriter.ApplyColumnFormatsToBody(
                            sheet,
                            page.Projection,
                            totalRowCount);

                        if (pageWrites.ColumnSizing == ColumnSizingMode.HeadersOnly)
                        {
                            SheetProjectionWriter.ApplyHeaderColumnSizing(sheet, page.Projection);
                        }

                        if (pageWrites.RowSizing == RowSizingMode.ForceSingleLine)
                        {
                            SheetProjectionWriter.ForceRowsToSingleLine(
                                sheet,
                                page.Projection,
                                totalRowCount);
                        }
                    }

                    SetStatusBar(excel, "Updating cells");
                    SheetProjectionWriter.ApplyPage(
                        sheet,
                        page.Projection,
                        applyColumnFormats: false,
                        autoFitColumns: pageWrites.ShouldAutoFitColumns,
                        rowSizingMode: pageWrites.RowSizing == RowSizingMode.FitEachPage
                            ? RowSizingMode.FitEachPage
                            : RowSizingMode.None,
                        preserveExistingColumnWidths: pageWrites.ShouldPreserveExistingColumnWidths);
                }
                finally
                {
                    writeStopwatch.Stop();
                    writeElapsed = writeStopwatch.Elapsed;
                    SessionFlowTrace.Log(
                        $"QueryTable page {page.PageNumber}: Excel write completed " +
                        $"records={page.Projection.Values.GetLength(0)} " +
                        $"elapsed={writeElapsed.TotalMilliseconds:0}ms");
                }
            });

            pageWrites.RecordPage(writeElapsed);
        }
    }

    public static void ApplyResult(
        ExcelWorksheet sheet,
        DataOperationResult result,
        ConnectorOptions options,
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
                SheetProjectionWriter.Apply(
                    sheet,
                    result.Projection,
                    options.ColumnSizingMode,
                    options.RowSizingMode);
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

    private sealed class PagedQueryWriteState
    {
        public PagedQueryWriteState(ColumnSizingMode columnSizingMode, RowSizingMode rowSizingMode)
        {
            ColumnSizing = columnSizingMode;
            RowSizing = rowSizingMode;
        }

        public ColumnSizingMode ColumnSizing { get; }

        public RowSizingMode RowSizing { get; }

        public bool HasPages { get; private set; }

        public bool ShouldAutoFitColumns =>
            ColumnSizing == ColumnSizingMode.AllDownloadedData
            || (ColumnSizing == ColumnSizingMode.FirstDownloadedPage && !HasPages);

        public bool ShouldPreserveExistingColumnWidths =>
            ColumnSizing == ColumnSizingMode.AllDownloadedData && HasPages;

        public TimeSpan ApplyElapsed { get; private set; }

        public void RecordPage(TimeSpan elapsed)
        {
            if (!HasPages)
            {
                HasPages = true;
            }

            ApplyElapsed += elapsed;
        }
    }
}
