using System;
using System.Windows;

namespace SalesforceRestAddin.Windows.Ui;

/// <summary>
/// Excel-DNA does not host a WPF <see cref="Application"/>; create one before modal UI.
/// Must be shut down explicitly on Excel exit or the process can stay alive (bugs.md #6).
/// </summary>
public static class WpfApplicationHost
{
    private static readonly object Sync = new();
    private static bool _initialized;

    public static void EnsureInitialized()
    {
        lock (Sync)
        {
            if (_initialized && Application.Current is not null)
            {
                return;
            }

            if (Application.Current is null)
            {
                _ = new Application
                {
                    ShutdownMode = ShutdownMode.OnExplicitShutdown,
                };
            }

            _initialized = true;
        }
    }

    /// <summary>
    /// Tears down the hosted WPF application so Excel can exit after dialogs/progress UI.
    /// Safe to call when WPF was never started.
    /// </summary>
    public static void Shutdown()
    {
        lock (Sync)
        {
            var app = Application.Current;
            if (app is null)
            {
                _initialized = false;
                return;
            }

            try
            {
                var dispatcher = app.Dispatcher;
                if (dispatcher.CheckAccess())
                {
                    ShutdownCore(app);
                }
                else
                {
                    dispatcher.Invoke(new Action(() => ShutdownCore(app)));
                }
            }
            catch
            {
                // Excel may already be tearing down the STA thread.
            }
            finally
            {
                _initialized = false;
            }
        }
    }

    private static void ShutdownCore(Application app)
    {
        try
        {
            var windows = app.Windows;
            for (var i = windows.Count - 1; i >= 0; i--)
            {
                try
                {
                    var window = windows[i];
                    if (window.IsVisible)
                    {
                        window.Close();
                    }
                }
                catch
                {
                    // Ignore per-window close failures during Excel shutdown.
                }
            }
        }
        catch
        {
            // Windows collection may already be invalid.
        }

        try
        {
            app.Shutdown();
        }
        catch
        {
            // Already shutting down.
        }

        try
        {
            // Extra belt-and-braces: stop the dispatcher so it cannot keep the process alive.
            app.Dispatcher.InvokeShutdown();
        }
        catch
        {
            // Dispatcher may already be shut down.
        }
    }
}
