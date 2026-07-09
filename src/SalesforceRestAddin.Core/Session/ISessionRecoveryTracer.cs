namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Optional trace sink for mid-call HTTP session recovery (spike / tests).
/// </summary>
public interface ISessionRecoveryTracer
{
    void AddStep(string step);
}
