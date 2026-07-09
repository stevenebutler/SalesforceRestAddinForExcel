using System.Collections.Generic;

namespace SalesforceRestAddin.Core.Session;

public sealed class SessionRecoveryTrace : ISessionRecoveryTracer
{
    public IList<string> Steps { get; } = new List<string>();

    public void Clear() => Steps.Clear();

    public void AddStep(string step) => Steps.Add(step);
}
