using System.Threading;
using SalesforceRestAddin.Windows.Ui;

namespace SalesforceRestAddin;

internal static class OperationGate
{
    private static int _busy;

    public static bool TryEnter()
    {
        if (Interlocked.CompareExchange(ref _busy, 1, 0) != 0)
        {
            ErrorDialogWindow.Show(
                AddInHost.AddInTitle,
                "An add-in operation is already in progress. Wait for it to finish.");
            return false;
        }

        return true;
    }

    public static void Exit() => Interlocked.Exchange(ref _busy, 0);
}
