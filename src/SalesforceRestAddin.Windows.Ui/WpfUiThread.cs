using System;
using System.Threading.Tasks;
using System.Windows;

namespace SalesforceRestAddin.Windows.Ui;

/// <summary>
/// Runs modal WPF work on the application dispatcher without <see cref="Dispatcher.InvokeAsync"/>,
/// which can stall for minutes when Excel's macro thread is blocked on <c>GetAwaiter().GetResult()</c>.
/// </summary>
public static class WpfUiThread
{
    public static Task<T> RunAsync<T>(Func<T> func)
    {
        if (func is null)
        {
            throw new ArgumentNullException(nameof(func));
        }

        WpfApplicationHost.EnsureInitialized();
        var dispatcher = Application.Current.Dispatcher;
        if (dispatcher.CheckAccess())
        {
            return Task.FromResult(func());
        }

        return Task.FromResult(dispatcher.Invoke(func));
    }

    public static void Run(Action action)
    {
        if (action is null)
        {
            throw new ArgumentNullException(nameof(action));
        }

        RunAsync(() =>
        {
            action();
            return true;
        }).GetAwaiter().GetResult();
    }
}
