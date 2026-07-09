using System.Threading;
using System.Threading.Tasks;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Tests;

public sealed class FakeLoginOptionsPresenter : ILoginOptionsPresenter
{
    public LoginOptionsPresenterResult? NextResult { get; set; }

    public int ShowCount { get; private set; }

    public UserLoginPreferences? LastPreferences { get; private set; }

    public Task<LoginOptionsPresenterResult?> ShowAsync(
        UserLoginPreferences? currentPreferences,
        CancellationToken cancellationToken = default)
    {
        ShowCount++;
        LastPreferences = currentPreferences;
        return Task.FromResult(NextResult);
    }
}
