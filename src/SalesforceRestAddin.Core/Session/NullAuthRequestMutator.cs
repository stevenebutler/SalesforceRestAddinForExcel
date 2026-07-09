namespace SalesforceRestAddin.Core.Session;

public sealed class NullAuthRequestMutator : IAuthRequestMutator
{
    public static NullAuthRequestMutator Instance { get; } = new();

    public string MutateBearer(string accessToken, bool isRetryAfterRecovery = false) => accessToken;

    public string MutateRefreshToken(string refreshToken) => refreshToken;

    public bool ForceFailureAfterRecovery => false;
}
