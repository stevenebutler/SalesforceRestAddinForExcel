using System;
using System.Net.Http;
using SalesforceRestAddin.Core;
using SalesforceRestAddin.Core.Net;
using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Rest;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Windows.Ui;

/// <summary>
/// Wires Core session gate and authenticated HTTP client for spike and future Excel-DNA host.
/// </summary>
public static class SessionGateComposer
{
    public sealed class SessionServices
    {
        public required SessionGate Gate { get; init; }

        public required SalesforceAuthenticatedClient AuthenticatedClient { get; init; }

        public required HttpClient HttpClient { get; init; }

        public required ISessionCredentialStore CredentialStore { get; init; }

        public required IUserLoginPreferencesStore PreferencesStore { get; init; }

        public required IMetadataCache MetadataCache { get; init; }

        public required ParsedMetadataCache ParsedMetadataCache { get; init; }

        public IAuthRequestMutator AuthMutator { get; init; } = NullAuthRequestMutator.Instance;

        public SessionRecoveryTrace? RecoveryTrace { get; init; }

        public SalesforceDataClient CreateDataClient(SessionContext? session = null) =>
            new(AuthenticatedClient, session ?? SessionContext.Current, MetadataCache, ParsedMetadataCache);

        /// <summary>Clears L1 + L2 metadata for the current session instance only.</summary>
        public void ClearCurrentInstanceMetadataCache()
        {
            var instanceUrl = SessionContext.Current.InstanceUrl;
            if (instanceUrl is null || string.IsNullOrWhiteSpace(instanceUrl))
            {
                return;
            }

            ParsedMetadataCache.ClearInstance(instanceUrl);
            MetadataCache.ClearInstance(instanceUrl);
        }
    }

    public static SessionServices Create(
        HttpClient? http = null,
        ISessionCredentialStore? credentialStore = null,
        IUserLoginPreferencesStore? preferencesStore = null,
        IAuthRequestMutator? authMutator = null,
        SessionContext? session = null,
        SessionRecoveryTrace? recoveryTrace = null,
        IMetadataCache? metadataCache = null,
        ParsedMetadataCache? parsedMetadataCache = null)
    {
        credentialStore ??= new WindowsCredentialSessionStore();
        preferencesStore ??= new JsonUserLoginPreferencesStore(SalesforceRestAddinDataPaths.PreferencesFile);
        authMutator ??= NullAuthRequestMutator.Instance;
        session ??= SessionContext.Current;
        recoveryTrace ??= new SessionRecoveryTrace();
        metadataCache ??= new FileMetadataCache(SalesforceRestAddinDataPaths.MetadataCacheDirectory);
        parsedMetadataCache ??= new ParsedMetadataCache();

        var httpClient = http ?? SalesforceHttpClientFactory.Create();
        var oauth = SalesforceOAuthOptions.Default;
        var authenticator = new SessionAuthenticator(
            credentialStore,
            new SalesforceOAuthClient(httpClient),
            oauth);
        var interactive = new WindowsInteractiveLoginHandler(authenticator, oauth);
        var orchestrator = new SessionLoginOrchestrator(session, authenticator, interactive, authMutator);
        var identityClient = new SalesforceIdentityClient(httpClient);
        var gate = new SessionGate(
            session,
            preferencesStore,
            new WindowsLoginOptionsPresenter(credentialStore),
            orchestrator,
            identityClient);
        var refresher = new SessionAccessTokenRefresher(authenticator, authMutator, recoveryTrace);
        var authenticatedClient = new SalesforceAuthenticatedClient(
            httpClient,
            session,
            preferencesStore,
            refresher,
            orchestrator,
            authMutator,
            recoveryTrace);

        return new SessionServices
        {
            Gate = gate,
            AuthenticatedClient = authenticatedClient,
            HttpClient = httpClient,
            CredentialStore = credentialStore,
            PreferencesStore = preferencesStore,
            MetadataCache = metadataCache,
            ParsedMetadataCache = parsedMetadataCache,
            AuthMutator = authMutator,
            RecoveryTrace = recoveryTrace,
        };
    }
}
