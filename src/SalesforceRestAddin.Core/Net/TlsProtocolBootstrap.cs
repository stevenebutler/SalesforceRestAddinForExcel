#if NET48
using System.Net;
#endif

namespace SalesforceRestAddin.Core.Net;

/// <summary>
/// Excel and other legacy hosts may not negotiate TLS 1.2 unless enabled explicitly.
/// The VSTO add-in set this in <c>ThisAddIn_Startup</c>; HttpClient uses <see cref="ServicePointManager"/> on net48.
/// </summary>
public static class TlsProtocolBootstrap
{
#if NET48
    private static bool _enabled;
#endif

    public static void EnsureEnabled()
    {
#if NET48
        if (_enabled)
        {
            return;
        }

        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        _enabled = true;
#endif
    }
}
