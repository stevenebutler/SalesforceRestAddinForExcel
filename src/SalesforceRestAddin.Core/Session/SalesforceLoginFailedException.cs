using System;

namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Sign-in failed after options / silent refresh / interactive OAuth — not a user cancel.
/// </summary>
public sealed class SalesforceLoginFailedException : Exception
{
    public SalesforceLoginFailedException(string message)
        : base(message)
    {
    }

    public SalesforceLoginFailedException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    public SessionLoginDiagnostics? Diagnostics { get; init; }
}
