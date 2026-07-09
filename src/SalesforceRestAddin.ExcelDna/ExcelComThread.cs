using System;
using System.Threading;
using System.Windows;
using System.Windows.Threading;
using ExcelDna.Integration;
using SalesforceRestAddin.Windows.Ui;

namespace SalesforceRestAddin;

/// <summary>
/// Thin marshal onto Excel's main STA for COM. Not a wait/pump — use
/// <see cref="ExcelStaAsyncHost"/> for async work that needs a nested message loop.
/// After <c>await</c>, continuations run on the thread pool; touching COM there creates
/// RCWs that pin <c>EXCEL.EXE</c> (bugs.md #6 / Excel-DNA guidance).
/// </summary>
internal static class ExcelComThread
{
    /// <summary>True when this thread is safe for Excel COM (WPF dispatcher / Excel STA).</summary>
    public static bool IsMainThread()
    {
        WpfApplicationHost.EnsureInitialized();
        var dispatcher = Application.Current?.Dispatcher;
        return dispatcher is not null && dispatcher.CheckAccess();
    }

    /// <summary>Runs <paramref name="action"/> on the Excel STA, synchronously.</summary>
    public static void Invoke(Action action)
    {
        if (action is null)
        {
            throw new ArgumentNullException(nameof(action));
        }

        if (IsMainThread())
        {
            action();
            return;
        }

        // While ExcelStaAsyncHost holds modal ShowDialog, Excel's macro queue is blocked.
        // Prefer the WPF dispatcher (same STA) so COM runs without deadlocking on QueueAsMacro.
        WpfApplicationHost.EnsureInitialized();
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher is not null)
        {
            dispatcher.Invoke(action, DispatcherPriority.Normal);
            return;
        }

        var done = new ManualResetEventSlim(false);
        Exception? error = null;
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                error = ex;
            }
            finally
            {
                done.Set();
            }
        });
        done.Wait();
        if (error is not null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error).Throw();
        }
    }

    /// <summary>Queues <paramref name="action"/> on the Excel STA without blocking the caller.</summary>
    public static void BeginInvoke(Action action)
    {
        if (action is null)
        {
            throw new ArgumentNullException(nameof(action));
        }

        if (IsMainThread())
        {
            action();
            return;
        }

        WpfApplicationHost.EnsureInitialized();
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher is not null)
        {
            dispatcher.BeginInvoke(action, DispatcherPriority.Normal);
            return;
        }

        ExcelAsyncUtil.QueueAsMacro(() => action());
    }
}
