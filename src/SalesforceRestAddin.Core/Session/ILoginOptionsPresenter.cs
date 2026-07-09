using System.Threading;
using System.Threading.Tasks;

namespace SalesforceRestAddin.Core.Session;

public sealed class LoginOptionsPresenterResult
{
    public required OAuth.SalesforceLoginOptions LoginOptions { get; init; }

    public string? PreferredApiVersion { get; init; }

    public bool ForceSignInAgain { get; init; }
}

/// <summary>
/// Shows tenant/environment login options before silent refresh or interactive OAuth.
/// </summary>
public interface ILoginOptionsPresenter
{
    Task<LoginOptionsPresenterResult?> ShowAsync(
        UserLoginPreferences? currentPreferences,
        CancellationToken cancellationToken = default);
}
