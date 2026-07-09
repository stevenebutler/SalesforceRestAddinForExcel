using System.Threading;
using System.Windows;

namespace SalesforceRestAddin.Windows.Ui;

public static class ConfirmationDialogWindow
{
    public static bool Show(string title, string message)
    {
        var confirmed = false;
        WpfUiThread.Run(() =>
        {
            confirmed = MessageBox.Show(
                message,
                title,
                MessageBoxButton.YesNo,
                MessageBoxImage.Question) == MessageBoxResult.Yes;
        });
        return confirmed;
    }
}
