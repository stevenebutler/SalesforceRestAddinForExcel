using System.Security.Cryptography;
using System.Text;

namespace SalesforceRestAddin.Core.OAuth;

/// <summary>
/// PKCE code verifier/challenge pair for public OAuth clients.
/// </summary>
public sealed class PkcePair
{
    public required string Verifier { get; init; }

    public required string Challenge { get; init; }

    public static PkcePair Create()
    {
        var verifier = GenerateVerifier();
        return new PkcePair
        {
            Verifier = verifier,
            Challenge = ComputeS256Challenge(verifier),
        };
    }

    internal static string GenerateVerifier()
    {
        // RFC 7636: 43–128 characters from the unreserved set.
        var bytes = new byte[32];
        using (var rng = RandomNumberGenerator.Create())
        {
            rng.GetBytes(bytes);
        }

        return Base64UrlEncode(bytes);
    }

    internal static string ComputeS256Challenge(string verifier)
    {
        using var sha = SHA256.Create();
        var hash = sha.ComputeHash(Encoding.ASCII.GetBytes(verifier));
        return Base64UrlEncode(hash);
    }

    private static string Base64UrlEncode(byte[] data) =>
        Convert.ToBase64String(data)
            .TrimEnd('=')
            .Replace('+', '-')
            .Replace('/', '_');
}
