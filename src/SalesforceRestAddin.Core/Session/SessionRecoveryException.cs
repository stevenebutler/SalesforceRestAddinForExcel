using System;

namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Mid-call session recovery failed after refresh and interactive sign-in were exhausted.
/// </summary>
public sealed class SessionRecoveryException : Exception
{
    public SessionRecoveryException(string message)
        : base(message)
    {
    }

    public SessionRecoveryException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
