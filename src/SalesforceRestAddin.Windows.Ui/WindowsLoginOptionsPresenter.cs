using System.Threading;
using System.Threading.Tasks;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Windows.Ui;

public sealed class WindowsLoginOptionsPresenter : ILoginOptionsPresenter
{
    private readonly ISessionCredentialStore? _credentialStore;

    public WindowsLoginOptionsPresenter(ISessionCredentialStore? credentialStore = null)
    {
        _credentialStore = credentialStore;
    }

    public Task<LoginOptionsPresenterResult?> ShowAsync(
        UserLoginPreferences? currentPreferences,
        CancellationToken cancellationToken = default) =>
        WpfUiThread.RunAsync(() => ShowOnUiThread(currentPreferences, cancellationToken));

    private LoginOptionsPresenterResult? ShowOnUiThread(
        UserLoginPreferences? currentPreferences,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var window = new LoginOptionsWindow(currentPreferences, _credentialStore);
        if (window.ShowDialog() != true)
        {
            return null;
        }

        var result = window.GetResult();
        if (result is null)
        {
            return null;
        }

        return new LoginOptionsPresenterResult
        {
            LoginOptions = result.LoginOptions,
            PreferredApiVersion = result.PreferredApiVersion,
            ForceSignInAgain = result.ForceSignInAgain,
        };
    }
}
