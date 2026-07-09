using System;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Tests;

public sealed class TestAuthRequestMutator : IAuthRequestMutator
{
    public bool CorruptBearer { get; set; }

    public bool CorruptRefresh { get; set; }

    public bool ForceFailureAfterRecovery { get; set; }

    public string MutateBearer(string accessToken, bool isRetryAfterRecovery = false)
    {
        if (ForceFailureAfterRecovery && isRetryAfterRecovery)
        {
            return $"CORRUPTED:{accessToken}";
        }

        if (CorruptBearer && !isRetryAfterRecovery
            && string.Equals(accessToken, "stale-access", StringComparison.Ordinal))
        {
            return $"CORRUPTED:{accessToken}";
        }

        return accessToken;
    }

    public string MutateRefreshToken(string refreshToken) =>
        CorruptRefresh ? $"CORRUPTED:{refreshToken}" : refreshToken;
}
