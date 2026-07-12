using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using SalesforceRestAddin.Core;
using SalesforceRestAddin.Core.Rest;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Tests.Fixtures;

namespace SalesforceRestAddin.Tests.Rest;

public sealed class SalesforceDataClientTests
{
    [Test]
    public async Task T_API_01_Describe_Parses_SObjectDescribe()
    {
        var (client, _) = CreateClient(handler =>
        {
            handler.Enqueue(HttpStatusCode.OK, LoadFixture("AccountDescribe.json"));
        });

        var describe = await client.DescribeAsync("Account");

        await Assert.That(describe.Name).IsEqualTo("Account");
        await Assert.That(describe.Fields.Any(f => f.Name == "Custom__c")).IsTrue();
    }

    [Test]
    public async Task T_API_02_Query_Encodes_Soql_In_Url()
    {
        var (client, h) = CreateClient(handler =>
        {
            handler.Enqueue(HttpStatusCode.OK, LoadFixture("AccountQueryPage1.json"));
        });

        await client.QueryAsync("SELECT Id FROM Account");

        await Assert.That(h.Requests[0].RequestUri!.Query).Contains("SELECT");
    }

    [Test]
    public async Task T_API_02b_Retrieve_Skips_Null_Composite_Entries()
    {
        var (client, _) = CreateClient(handler =>
        {
            handler.Enqueue(
                HttpStatusCode.OK,
                "[null,{\"attributes\":{\"type\":\"Account\"},\"Id\":\"001xx0000000001\",\"Name\":\"Acme\"}]");
        });

        var records = await client.RetrieveAsync(
            "Account",
            new[] { "001xx0000000000", "001xx0000000001" },
            new[] { "Id", "Name" });

        await Assert.That(records.Count).IsEqualTo(1);
        await Assert.That(records[0]["Id"]).IsEqualTo("001xx0000000001");
        await Assert.That(records[0]["Name"]).IsEqualTo("Acme");
    }

    [Test]
    public async Task ResolveAggregateCount_Uses_Expr0_Not_TotalSize()
    {
        var page = new QueryResultPage
        {
            TotalSize = 1,
            Done = true,
            Records = new[]
            {
                new Dictionary<string, object?> { ["expr0"] = 42L },
            },
        };

        await Assert.That(SalesforceDataClient.ResolveAggregateCount(page)).IsEqualTo(42);
    }

    [Test]
    public async Task ResolveAggregateCount_Falls_Back_To_TotalSize_For_Row_Queries()
    {
        var page = new QueryResultPage
        {
            TotalSize = 2,
            Done = true,
            Records = new[]
            {
                new Dictionary<string, object?> { ["Id"] = "001A" },
                new Dictionary<string, object?> { ["Id"] = "001B" },
            },
        };

        await Assert.That(SalesforceDataClient.ResolveAggregateCount(page)).IsEqualTo(2);
    }

    [Test]
    public async Task T_API_03_ListObjects_Parses_Summaries()
    {
        var (client, _) = CreateClient(handler =>
        {
            handler.Enqueue(HttpStatusCode.OK, LoadFixture("SObjectList.json"));
        });

        var objects = await client.ListObjectsAsync();
        await Assert.That(objects.Count).IsEqualTo(3);
        await Assert.That(objects.Any(o => o.Name == "Account")).IsTrue();
    }

    [Test]
    public async Task T_API_05_Create_Gzips_Large_Composite_Body()
    {
        string? contentEncoding = null;
        byte[]? bodyBytes = null;

        var (client, _) = CreateClient(handler =>
        {
            handler.Enqueue(request =>
            {
                contentEncoding = request.Content?.Headers.ContentEncoding.ToString();
                bodyBytes = request.Content?.ReadAsByteArrayAsync().GetAwaiter().GetResult();
                return new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(
                        "[{\"success\":true,\"id\":\"001xx0000000001\",\"errors\":[]}]",
                        Encoding.UTF8,
                        "application/json"),
                };
            });
        });

        var records = new List<Dictionary<string, object?>>();
        for (var i = 0; i < 40; i++)
        {
            records.Add(new Dictionary<string, object?>
            {
                ["attributes"] = new Dictionary<string, object?> { ["type"] = "Account" },
                ["Name"] = $"Account-{i}-with-enough-payload-bytes-to-cross-gzip-threshold",
            });
        }

        var results = await client.CreateAsync(records, sendPreventAutoAssignHeader: false);

        await Assert.That(results.Count).IsEqualTo(1);
        await Assert.That(contentEncoding).IsEqualTo("gzip");
        await Assert.That(bodyBytes).IsNotNull();
        await Assert.That(bodyBytes!.Length).IsGreaterThan(0);
    }

    [Test]
    public async Task T_API_06_Create_Leaves_Small_Body_Uncompressed()
    {
        string? contentEncoding = null;

        var (client, _) = CreateClient(handler =>
        {
            handler.Enqueue(request =>
            {
                contentEncoding = request.Content?.Headers.ContentEncoding.ToString();
                return new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(
                        "[{\"success\":true,\"id\":\"001xx0000000001\",\"errors\":[]}]",
                        Encoding.UTF8,
                        "application/json"),
                };
            });
        });

        var records = new[]
        {
            new Dictionary<string, object?>
            {
                ["attributes"] = new Dictionary<string, object?> { ["type"] = "Account" },
                ["Name"] = "A",
            },
        };

        await client.CreateAsync(records, sendPreventAutoAssignHeader: false);

        await Assert.That(string.IsNullOrEmpty(contentEncoding)).IsTrue();
    }

    private static (SalesforceDataClient client, SequentialMockHttpHandler handler) CreateClient(
        Action<SequentialMockHttpHandler>? configure = null)
    {
        var handler = new SequentialMockHttpHandler();
        configure?.Invoke(handler);
        var http = new HttpClient(handler) { BaseAddress = new Uri("https://example.my.salesforce.com") };
        var session = new SessionContext
        {
            AccessToken = "token",
            InstanceUrl = "https://example.my.salesforce.com",
            Id = "https://login.salesforce.com/id/00D/005",
            IssuedAtUnixMs = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(),
        };
        session.SetApiVersion("66.0");
        var creds = new InMemorySessionCredentialStore();
        var prefs = new InMemoryUserLoginPreferencesStore();
        var oauth = SalesforceRestAddin.Core.OAuth.SalesforceOAuthOptions.Default;
        var authenticator = new SessionAuthenticator(creds, new SalesforceRestAddin.Core.OAuth.SalesforceOAuthClient(http), oauth);
        var orchestrator = new SessionLoginOrchestrator(session, authenticator, new FakeInteractiveLoginHandler());
        var authClient = new SalesforceAuthenticatedClient(
            http,
            session,
            prefs,
            new SessionAccessTokenRefresher(authenticator),
            orchestrator);
        return (new SalesforceDataClient(authClient, session), handler);
    }

    private static string LoadFixture(string name)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Fixtures", name);
        return File.ReadAllText(path);
    }
}
