using System;

namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// User cancelled login options or interactive OAuth; the underlying operation must not continue.
/// </summary>
public sealed class SalesforceLoginCancelledException : Exception
{
    public SalesforceLoginCancelledException(string message)
        : base(message)
    {
    }

    public SalesforceLoginCancelledException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
