using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
using Microsoft.Office.Interop.Excel;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;
using System.Windows;

namespace SalesforceRestAddin;

/// <summary>
/// Application entry points invoked by ribbon, COM API, and future XLL macros.
/// </summary>
public static class AddInHost
{
    public const string AddInTitle = ProductBranding.ProductName;

    public static readonly SalesForceAddInApi Api = new();

    private static SessionGateComposer.SessionServices? _services;
    private static IConnectorOptionsStore? _optionsStore;

    public static void Initialize(SessionGateComposer.SessionServices? services = null, IConnectorOptionsStore? optionsStore = null)
    {
        _services = services ?? SessionGateComposer.Create();
        _optionsStore = optionsStore ?? new JsonConnectorOptionsStore(SalesforceRestAddinDataPaths.ConnectorOptionsFile);
        DataOperationHost.Configure(_services, _optionsStore);
        SalesforceRestAddinRibbon.Configure(_services.PreferencesStore);
        SalesforceRestAddinRibbon.RefreshSessionDependentUi();
    }

    private static SessionGateComposer.SessionServices Services =>
        _services ??= SessionGateComposer.Create();

    private static ConnectorOptions LoadOptions() =>
        (_optionsStore ?? new JsonConnectorOptionsStore(SalesforceRestAddinDataPaths.ConnectorOptionsFile)).Load();

    public static void ShowAbout()
    {
        WpfUiThread.Run(() =>
        {
            var window = new AboutWindow(BuildInfo.CommitId, SessionContext.Current.IsLoggedIn);
            window.ShowDialog();
        });
    }

    public static void QueryTableWizard() => RunGated("Table Query Wizard", RunWizard);

    public static void DescribeSforceObject() => RunGated("Describe Sforce Object", RunDescribe);

    public static void ShowOptions()
    {
        WpfUiThread.Run(() =>
        {
            var store = _optionsStore ?? new JsonConnectorOptionsStore(SalesforceRestAddinDataPaths.ConnectorOptionsFile);
            var window = new OptionsWindow(
                store.Load(),
                Services.ClearCurrentInstanceMetadataCache,
                SessionContext.Current.InstanceUrl);
            if (window.ShowDialog() == true)
            {
                store.Save(window.GetOptions());
            }
        });
    }

    public static void Logout()
    {
        var cleared = Services.Gate.Logout();
        SalesforceRestAddinRibbon.RefreshSessionDependentUi();
        if (cleared)
        {
            WpfUiThread.Run(() =>
                MessageBox.Show(
                    "You have been signed out of Salesforce.",
                    AddInTitle,
                    MessageBoxButton.OK,
                    MessageBoxImage.Information));
        }
    }

    public static void QuerySelectedRows() => RunGated("Query Selected Rows", () => RunBindingOperation(QueryRowsOperation));

    public static void QueryTableData() => RunGated("Query Table Data", DataOperationHost.RunQueryTableData);

    public static void UpdateSelectedCells() => RunGated("Update Selected Cells", () => RunBindingOperation(UpdateCellsOperation, confirmIncludeHidden: true));

    public static void InsertSelectedRows() => RunGated("Insert Selected Rows", () => RunBindingOperation(InsertRowsOperation, confirmInsert: true));

    public static void DeleteSelectedRecords() => RunGated("Delete Selected Records", () => RunBindingOperation(DeleteRecordsOperation, confirmDelete: true));

    public static void RefreshTableData() => RunGated("Refresh Table Data", () => RunBindingOperation(RefreshTableDataOperation));

    /// <summary>Ensures a live session, then refreshes ribbon label / Logout enablement.</summary>
    internal static void EnsureLoggedInAndRefreshUi(SessionGate gate)
    {
        var excel = (ExcelApplication)ExcelDnaUtil.Application;
        // Silent refresh / metadata can take seconds (bugs.md #7) — sole pump is ExcelStaAsyncHost.
        ExcelStaAsyncHost.Run(
            "Signing in…",
            excel,
            (ct, reportStatus) => gate.EnsureLoggedInAsync(ct, reportStatus));

        SalesforceRestAddinRibbon.RefreshSessionDependentUi();
    }

    private static void RunGated(string operationName, System.Action operation)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            if (!OperationGate.TryEnter())
            {
                return;
            }

            try
            {
                SessionFlowTrace.ResetForOperation(operationName);
                operation();
            }
            catch (SalesforceLoginCancelledException ex)
            {
                ShowLoginError("Sign-in was cancelled", ex.Message, Services.Gate.LastDiagnostics);
            }
            catch (SalesforceLoginFailedException ex)
            {
                ShowLoginError("Sign-in failed", ex.Message, ex.Diagnostics ?? Services.Gate.LastDiagnostics);
            }
            catch (Exception ex)
            {
                SessionFlowTrace.LogException(operationName, ex);
                var details = ExceptionDetailFormatter.Format(operationName, ex)
                    + Environment.NewLine
                    + Environment.NewLine
                    + SessionFlowTrace.FormatRecent();
                ErrorDialogWindow.Show(AddInTitle, details);
            }
            finally
            {
                OperationGate.Exit();
            }
        });
    }

    private static void ShowLoginError(string title, string message, SessionLoginDiagnostics? diagnostics)
    {
        var details = new StringBuilder();
        details.AppendLine(message);
        details.AppendLine();
        details.AppendLine("--- Gate diagnostics ---");
        details.AppendLine(SessionLoginDiagnosticsText.Format(diagnostics));
        details.AppendLine();
        details.AppendLine(SessionFlowTrace.FormatRecent());
        ErrorDialogWindow.Show(AddInTitle, $"{title}{Environment.NewLine}{Environment.NewLine}{details}");
    }

    private static void RunWizard()
    {
        var excel = (ExcelApplication)ExcelDnaUtil.Application;
        if (!WorkbookContextGuard.TryRequireContext(excel, WorkbookContextRequirement.Worksheet))
        {
            ErrorDialogWindow.Show(
                AddInTitle,
                WorkbookContextGuard.GetMessage(WorkbookContextRequirement.Worksheet));
            return;
        }

        EnsureLoggedInAndRefreshUi(Services.Gate);
        var selection = (Range)excel.Selection;
        var anchor = PromptAnchorCell(excel, selection.Row, selection.Column);
        if (anchor < 0)
        {
            SessionFlowTrace.Log("Table Query Wizard: anchor selection cancelled.");
            return;
        }

        var anchorRow = anchor / 10000;
        var anchorColumn = anchor % 10000;
        var client = Services.CreateDataClient();
        var objects = ExcelStaAsyncHost.Run(
            "Loading objects…",
            excel,
            (ct, report) =>
            {
                report("Loading Salesforce objects…");
                return client.ListObjectsAsync(ct);
            });

        // Describe is HTTP-only; do not QueueAsMacro while ShowDialog blocks the macro thread.
        Func<string, Task<SObjectDescribe>> describeLoader = apiName => client.DescribeAsync(apiName);

        TableQueryWizardResult? wizardResult = null;
        WpfUiThread.Run(() =>
        {
            var wizard = new TableQueryWizardWindow(objects, describeLoader, anchorRow, anchorColumn);
            if (wizard.ShowDialog() == true)
            {
                wizardResult = wizard.GetResult();
            }
        });

        if (wizardResult is null)
        {
            SessionFlowTrace.Log("Table Query Wizard: cancelled.");
            return;
        }

        var sheet = (Worksheet)excel.ActiveSheet;
        var startRow = wizardResult.AnchorRow;
        var startColumn = wizardResult.AnchorColumn;
        var options = LoadOptions();

        SessionFlowTrace.Log(
            $"Table Query Wizard: writing layout object={wizardResult.Describe.Name} fields={wizardResult.Fields.Count} anchor=R{startRow}C{startColumn}");

        WizardTableLayoutWriter.Apply(
            sheet,
            startRow,
            startColumn,
            wizardResult.Describe,
            wizardResult.Fields,
            wizardResult.Criteria,
            options.ColumnSizingMode == ColumnSizingMode.HeadersOnly
                ? AutomaticSizingMode.Width
                : AutomaticSizingMode.None);

        ((Range)sheet.Cells[startRow, startColumn]).Select();
        DataOperationHost.RunQueryTableData();
    }

    private static void RunDescribe()
    {
        var excel = (ExcelApplication)ExcelDnaUtil.Application;
        if (!WorkbookContextGuard.TryRequireContext(excel, WorkbookContextRequirement.Workbook))
        {
            ErrorDialogWindow.Show(
                AddInTitle,
                WorkbookContextGuard.GetMessage(WorkbookContextRequirement.Workbook));
            return;
        }

        EnsureLoggedInAndRefreshUi(Services.Gate);
        var client = Services.CreateDataClient();
        var options = LoadOptions();
        var objects = ExcelStaAsyncHost.Run(
            "Loading objects…",
            excel,
            (ct, report) =>
            {
                report("Loading Salesforce objects…");
                return client.ListObjectsAsync(ct);
            });

        IReadOnlyList<string>? selected = null;
        WpfUiThread.Run(() =>
        {
            var picker = new DescribeObjectPickerWindow(objects);
            if (picker.ShowDialog() == true)
            {
                selected = picker.GetSelectedApiNames();
            }
        });

        if (selected is null || selected.Count == 0)
        {
            return;
        }

        var failures = new List<string>();
        ExcelStaAsyncHost.Run(
            "Describe Sforce Object",
            excel,
            async (ct, report) =>
            {
                foreach (var objectName in selected)
                {
                    ct.ThrowIfCancellationRequested();
                    report(OperationProgress.Status($"Describing {objectName}..."));
                    try
                    {
                        var result = await DescribeObject.RunAsync(client, new[] { objectName }, ct).ConfigureAwait(false);
                        // Sheet writes must stay on Excel STA — not the await continuation thread.
                        ExcelComThread.Invoke(() =>
                        {
                            var sheet = (Worksheet)excel.Worksheets.Add();
                            sheet.Name = DescribeSheetName.Sanitize(objectName);
                            if (result.Projection is not null)
                            {
                                SheetProjectionWriter.Apply(
                                    sheet,
                                    result.Projection,
                                    options.ColumnSizingMode,
                                    options.RowSizingMode);
                            }
                        });
                    }
                    catch (Exception ex)
                    {
                        failures.Add($"{objectName}: {ex.Message}");
                    }
                }

                return 0;
            });

        if (failures.Count > 0)
        {
            ErrorDialogWindow.Show(
                AddInTitle,
                "Some objects could not be described:" + Environment.NewLine + string.Join(Environment.NewLine, failures));
        }
    }

    private static void RunBindingOperation(
        Func<SalesforceDataClient, ForceTableBinding, ForceTableSelection, ConnectorOptions, CancellationToken, Task<DataOperationResult>> operation,
        bool confirmDelete = false,
        bool confirmIncludeHidden = false,
        bool confirmInsert = false)
    {
        var excel = (ExcelApplication)ExcelDnaUtil.Application;
        if (!WorkbookContextGuard.TryRequireContext(excel, WorkbookContextRequirement.SalesforceConnectorTable))
        {
            ErrorDialogWindow.Show(
                AddInTitle,
                WorkbookContextGuard.GetMessage(WorkbookContextRequirement.SalesforceConnectorTable));
            return;
        }

        EnsureLoggedInAndRefreshUi(Services.Gate);
        var sheet = (Worksheet)excel.ActiveSheet;
        ForceTableSnapshot snapshot;
        try
        {
            snapshot = ForceTableReader.Capture(excel, sheet, (Range)excel.Selection);
        }
        catch (InvalidOperationException ex)
        {
            ErrorDialogWindow.Show(AddInTitle, ex.Message);
            return;
        }

        var client = Services.CreateDataClient();
        var describe = ExcelStaAsyncHost.Run(
            "Loading metadata…",
            excel,
            (ct, report) =>
            {
                report($"Describing {snapshot.ObjectApiName}…");
                return client.DescribeAsync(snapshot.ObjectApiName, ct);
            });
        var bindingResult = ForceTableBinder.Bind(snapshot, describe);
        if (!bindingResult.Succeeded)
        {
            ErrorDialogWindow.Show(
                AddInTitle,
                $"Could not read the ForceConnector table:{Environment.NewLine}{Environment.NewLine}{bindingResult.Errors[0].Message}");
            return;
        }

        var binding = bindingResult.Binding!;
        ForceTableSelection selection;
        try
        {
            selection = ReadSelection(excel, snapshot);
        }
        catch (InvalidOperationException ex)
        {
            ErrorDialogWindow.Show(AddInTitle, ex.Message);
            return;
        }

        var options = LoadOptions();

        if (confirmDelete && !ConfirmationDialogWindow.Show(
                AddInTitle,
                FormatRecordCountConfirm("Delete", selection.BodyRowIndices.Count, "from", binding.Describe.Label)
                    + " This cannot be undone."))
        {
            return;
        }

        // Update confirmation is controlled by NoWarning; include-hidden only changes its wording.
        if (confirmIncludeHidden && !options.NoWarning && !ConfirmationDialogWindow.Show(
                AddInTitle,
                options.IncludeHiddenCells
                    ? $"Update {selection.BodyRowIndices.Count} row(s), including any hidden rows/columns in the selection?"
                    : FormatRecordCountConfirm("Update", selection.BodyRowIndices.Count, "in", binding.Describe.Label)))
        {
            return;
        }

        // Insert confirm unless NoWarning (FR-ISR-2); N = selected rows, not "New" count.
        if (confirmInsert && !options.NoWarning && !ConfirmationDialogWindow.Show(
                AddInTitle,
                FormatRecordCountConfirm("Insert", selection.BodyRowIndices.Count, "into", binding.Describe.Label)))
        {
            return;
        }

        DataOperationResult result;
        try
        {
            result = ExcelStaAsyncHost.Run(
                GetOperationTitle(operation),
                excel,
                (ct, report) =>
                {
                    report(OperationProgress.Status("Processing..."));
                    return operation(client, binding, selection, options, ct);
                });
        }
        catch (OperationCanceledException)
        {
            SessionFlowTrace.Log($"{GetOperationTitle(operation)}: cancelled by user.");
            return;
        }

        DataOperationHost.ApplyResult(sheet, result, options, binding, excel);
    }

    private static string GetOperationTitle(
        Func<SalesforceDataClient, ForceTableBinding, ForceTableSelection, ConnectorOptions, CancellationToken, Task<DataOperationResult>> operation)
    {
        if (operation == UpdateCellsOperation)
        {
            return "Update Selected Cells";
        }

        if (operation == InsertRowsOperation)
        {
            return "Insert Selected Rows";
        }

        if (operation == DeleteRecordsOperation)
        {
            return "Delete Selected Records";
        }

        if (operation == RefreshTableDataOperation)
        {
            return "Refresh Table Data";
        }

        return "Query Selected Rows";
    }

    private static Task<DataOperationResult> QueryRowsOperation(
        SalesforceDataClient client,
        ForceTableBinding binding,
        ForceTableSelection selection,
        ConnectorOptions options,
        CancellationToken cancellationToken) =>
        QueryRows.RunAsync(client, new QueryRowsInput
        {
            Binding = binding,
            Selection = selection,
            Options = options,
        }, cancellationToken);

    private static Task<DataOperationResult> RefreshTableDataOperation(
        SalesforceDataClient client,
        ForceTableBinding binding,
        ForceTableSelection selection,
        ConnectorOptions options,
        CancellationToken cancellationToken) =>
        QueryRows.RunAsync(client, new QueryRowsInput
        {
            Binding = binding,
            Selection = selection,
            Options = options,
            RefreshAll = true,
        }, cancellationToken);

    private static Task<DataOperationResult> UpdateCellsOperation(
        SalesforceDataClient client,
        ForceTableBinding binding,
        ForceTableSelection selection,
        ConnectorOptions options,
        CancellationToken cancellationToken) =>
        UpdateCells.RunAsync(client, new UpdateCellsInput
        {
            Binding = binding,
            Selection = selection,
            Options = options,
        }, cancellationToken);

    private static Task<DataOperationResult> InsertRowsOperation(
        SalesforceDataClient client,
        ForceTableBinding binding,
        ForceTableSelection selection,
        ConnectorOptions options,
        CancellationToken cancellationToken) =>
        InsertRows.RunAsync(client, new InsertRowsInput
        {
            Binding = binding,
            Selection = selection,
            Options = options,
        }, cancellationToken);

    private static Task<DataOperationResult> DeleteRecordsOperation(
        SalesforceDataClient client,
        ForceTableBinding binding,
        ForceTableSelection selection,
        ConnectorOptions options,
        CancellationToken cancellationToken) =>
        DeleteRecords.RunAsync(client, new DeleteRecordsInput
        {
            Binding = binding,
            Selection = selection,
            Options = options,
        }, cancellationToken);

    internal static void DataOperationHostApply(
        Worksheet sheet,
        DataOperationResult result,
        ConnectorOptions options,
        ForceTableBinding? binding = null,
        ExcelApplication? excel = null) =>
        DataOperationHost.ApplyResult(sheet, result, options, binding, excel);

    private static ForceTableSelection ReadSelection(ExcelApplication excel, ForceTableSnapshot snapshot)
    {
        var selection = (Range)excel.Selection;
        var isMultiArea = selection.Areas.Count > 1;
        var bodyStartRow = snapshot.StartRow + 2;
        var bodyEndRow = bodyStartRow + snapshot.BodyRowCount - 1;
        var tableStartCol = snapshot.StartColumn;
        var tableEndCol = snapshot.StartColumn + snapshot.ColumnCount - 1;

        // Index-only walk of Areas — values come from the bulk-captured snapshot.Body.
        var columnsByBodyRow = new Dictionary<int, HashSet<int>>();
        var minCol = int.MaxValue;
        var maxCol = int.MinValue;

        // Index Areas — do not foreach COM enumerators (RCWs can pin EXCEL.EXE; bugs.md #6).
        var areas = selection.Areas;
        var areaCount = areas.Count;
        for (var areaIndex = 1; areaIndex <= areaCount; areaIndex++)
        {
            var area = (Range)areas[areaIndex];
            var areaStartRow = area.Row;
            var areaEndRow = area.Row + area.Rows.Count - 1;
            var areaStartCol = area.Column;
            var areaEndCol = area.Column + area.Columns.Count - 1;

            for (var sheetRow = areaStartRow; sheetRow <= areaEndRow; sheetRow++)
            {
                // Header / criteria rows are ignored; only body cells participate.
                if (sheetRow < bodyStartRow)
                {
                    continue;
                }

                if (sheetRow > bodyEndRow)
                {
                    throw new InvalidOperationException(UpdateCells.OutsideTableMessage);
                }

                var bodyRow = sheetRow - bodyStartRow;
                for (var sheetCol = areaStartCol; sheetCol <= areaEndCol; sheetCol++)
                {
                    if (sheetCol < tableStartCol || sheetCol > tableEndCol)
                    {
                        throw new InvalidOperationException(UpdateCells.OutsideTableMessage);
                    }

                    var tableCol = sheetCol - tableStartCol;
                    if (!columnsByBodyRow.TryGetValue(bodyRow, out var cols))
                    {
                        cols = new HashSet<int>();
                        columnsByBodyRow[bodyRow] = cols;
                    }

                    cols.Add(tableCol);
                    if (tableCol < minCol)
                    {
                        minCol = tableCol;
                    }

                    if (tableCol > maxCol)
                    {
                        maxCol = tableCol;
                    }
                }
            }
        }

        if (columnsByBodyRow.Count == 0)
        {
            // Selection was only headers/criteria — keep a degenerate span for limit checks.
            return new ForceTableSelection
            {
                BodyRowIndices = Array.Empty<int>(),
                StartColumnIndex = 0,
                EndColumnIndex = -1,
                IsMultiArea = isMultiArea,
                ColumnsByBodyRow = new Dictionary<int, IReadOnlyList<int>>(),
            };
        }

        var bodyRowIndices = columnsByBodyRow.Keys.OrderBy(i => i).ToList();
        var map = bodyRowIndices.ToDictionary(
            row => row,
            row => (IReadOnlyList<int>)columnsByBodyRow[row].OrderBy(c => c).ToList());

        return new ForceTableSelection
        {
            BodyRowIndices = bodyRowIndices,
            StartColumnIndex = minCol,
            EndColumnIndex = maxCol,
            IsMultiArea = isMultiArea,
            ColumnsByBodyRow = map,
        };
    }

    private static string FormatRecordCountConfirm(string verb, int count, string preposition, string objectLabel)
    {
        var noun = count == 1 ? "record" : "records";
        return $"{verb} {count} {noun} {preposition} {objectLabel}?";
    }

    private static int PromptAnchorCell(ExcelApplication excel, int defaultRow, int defaultColumn)
    {
        var defaultAddress = ((Range)excel.Cells[defaultRow, defaultColumn]).Address[false, false];
        object picked;
        try
        {
            picked = excel.InputBox(
                "Select the top-left cell for the new ForceConnector table:",
                AddInTitle,
                defaultAddress,
                Type: 8);
        }
        catch
        {
            // Legacy InputBox type 8 throws when the user cancels in some Excel builds.
            return -1;
        }

        // Type 8 returns a Range on OK and false on Cancel (legacy TableWizard.WizardStep1).
        if (picked is Range selected && selected.Cells[1, 1] is Range anchor)
        {
            return anchor.Row * 10000 + anchor.Column;
        }

        return -1;
    }
}
