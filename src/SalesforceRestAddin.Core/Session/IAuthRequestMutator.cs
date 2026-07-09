namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Spike/tests hook to corrupt bearer or refresh tokens and force post-recovery failures.
/// </summary>
public interface IAuthRequestMutator
{
    /// <param name="isRetryAfterRecovery">
    /// True when re-sending after successful session recovery. CorruptBearer is ignored on retry unless
    /// <see cref="ForceFailureAfterRecovery"/> is set.
    /// </param>
    string MutateBearer(string accessToken, bool isRetryAfterRecovery = false);

    string MutateRefreshToken(string refreshToken);

    /// <summary>
    /// When true, the post-recovery REST retry still sends a corrupted bearer so recovery can be verified
    /// to give up if auth keeps failing.
    /// </summary>
    bool ForceFailureAfterRecovery { get; }
}
