using System.Threading;
using System.Threading.Tasks;
using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Tests;

public sealed class FakeInteractiveLoginHandler : IInteractiveLoginHandler
{
    public InteractiveLoginResult NextResult { get; set; } = InteractiveLoginResult.CancelledResult();

    public int SignInCount { get; private set; }

    public SalesforceLoginOptions? LastLoginOptions { get; private set; }

    public Task<InteractiveLoginResult> TrySignInAsync(
        SalesforceLoginOptions loginOptions,
        string? preferredApiVersion,
        CancellationToken cancellationToken = default)
    {
        SignInCount++;
        LastLoginOptions = loginOptions;
        return Task.FromResult(NextResult);
    }
}
