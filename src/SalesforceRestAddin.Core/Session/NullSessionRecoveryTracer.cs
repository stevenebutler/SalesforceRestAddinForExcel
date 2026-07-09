namespace SalesforceRestAddin.Core.Session;

public sealed class NullSessionRecoveryTracer : ISessionRecoveryTracer
{
    public static NullSessionRecoveryTracer Instance { get; } = new();

    public void AddStep(string step)
    {
    }
}
