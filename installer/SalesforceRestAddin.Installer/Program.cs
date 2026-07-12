using System;
using System.Windows;

namespace SalesforceRestAddin.Installer;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        var app = new Application
        {
            ShutdownMode = ShutdownMode.OnMainWindowClose,
        };

        var window = new InstallerWindow(args);
        app.Run(window);
    }
}
