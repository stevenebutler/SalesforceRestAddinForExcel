using System;
using System.Runtime.InteropServices;
using ExcelDna.ComInterop;
using ExcelDna.Integration;
using SalesforceRestAddin.Core.Net;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Windows.Ui;

namespace SalesforceRestAddin;

public sealed class ExcelAddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        try
        {
            TlsProtocolBootstrap.EnsureEnabled();
            SessionFlowTrace.Log("TLS: ServicePointManager.SecurityProtocol set to Tls12 (net48).");

            // Native WebView2Loader.dll is embedded in Windows.Ui and extracted to LocalAppData on demand.
            WebView2LoaderBootstrap.EnsureDeployed();
            // WPF Application is created lazily on first dialog (WpfUiThread) so a cold
            // Excel session without UI does not host an OnExplicitShutdown app.

            ComServer.DllRegisterServer();
            ExcelComAddInHelper.LoadComAddIn(new SalesforceRestAddinComAddIn());
            AddInHost.Initialize();
            SessionFlowTrace.Log($"AutoOpen completed. XllPath={ExcelDnaUtil.XllPath}");
        }
        catch (Exception ex)
        {
            try
            {
                SessionFlowTrace.LogException("AutoOpen failed", ex);
            }
            catch
            {
                // Tracing must not mask the original failure.
            }

            throw;
        }
    }

    public void AutoClose()
    {
        try
        {
            AddInShutdown.Run();
        }
        finally
        {
            ComServer.DllUnregisterServer();
        }
    }
}

/// <summary>
/// Shared teardown for Excel quit / add-in unload (bugs.md #6 — zombie EXCEL.EXE).
/// </summary>
internal static class AddInShutdown
{
    private static int _ran;

    public static void Run()
    {
        if (System.Threading.Interlocked.Exchange(ref _ran, 1) != 0)
        {
            return;
        }

        try
        {
            SessionFlowTrace.Log("AddInShutdown: beginning WPF + COM cleanup.");
            WpfApplicationHost.Shutdown();
            ReleaseOutstandingComObjects();
            SessionFlowTrace.Log("AddInShutdown: completed.");
        }
        catch (Exception ex)
        {
            try
            {
                SessionFlowTrace.LogException("AddInShutdown failed", ex);
            }
            catch
            {
                // Ignore tracing failures during teardown.
            }
        }
    }

    /// <summary>
    /// Excel-DNA guidance: outstanding RCWs (e.g. from <c>foreach</c> over Ranges) can keep
    /// Excel alive after the UI closes. Finalizers + AreComObjectsAvailableForCleanup drain them.
    /// </summary>
    private static void ReleaseOutstandingComObjects()
    {
        try
        {
            do
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            while (Marshal.AreComObjectsAvailableForCleanup());
        }
        catch
        {
            // Teardown must not throw into Excel.
        }
    }
}
