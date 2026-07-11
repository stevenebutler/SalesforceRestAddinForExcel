using System.Net;
using System.Net.Http;
using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Tests.Fixtures;

namespace SalesforceRestAddin.Tests.DataPlane;

public sealed class QueryTablePagedTests
{
    [Test]
    public async Task QueryMore_Starts_While_Previous_Page_Is_Awaiting_Its_Write()
    {
        var firstPageQueued = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseFirstWrite = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var nextRequestStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var nextRequestStartedWhileWriteBlocked = false;
        var queuedPages = new List<QueryTablePage>();

        var (client, handler) = SalesforceTestClients.Create();
        handler.Enqueue(HttpStatusCode.OK, ReadFixture("AccountDescribe.json"));
        handler.Enqueue(HttpStatusCode.OK, """
            {
              "totalSize": 2,
              "done": false,
              "nextRecordsUrl": "/services/data/v66.0/query/01g-next",
              "records": [
                { "attributes": { "type": "Account" }, "Id": "001A", "Name": "Acme", "Industry": "Technology" }
              ]
            }
            """);
        handler.Enqueue(_ =>
        {
            nextRequestStartedWhileWriteBlocked = !releaseFirstWrite.Task.IsCompleted;
            nextRequestStarted.SetResult();
            return JsonResponse("""
                {
                  "totalSize": 2,
                  "done": true,
                  "records": [
                    { "attributes": { "type": "Account" }, "Id": "001B", "Name": "Beta", "Industry": "Finance" }
                  ]
                }
                """);
        });

        var firstWrite = Task.CompletedTask;
        var result = await QueryTable.RunPagedAsync(
            client,
            new QueryTableInput
            {
                Snapshot = ForceTableSnapshots.ValidAccountTable(),
                Options = ConnectorOptions.Default,
            },
            (page, _) =>
            {
                queuedPages.Add(page);
                if (queuedPages.Count == 1)
                {
                    firstPageQueued.SetResult();
                    firstWrite = SimulateFirstPageWriteAsync(firstPageQueued.Task, releaseFirstWrite.Task);
                }

                return Task.CompletedTask;
            });

        await nextRequestStarted.Task;
        await Assert.That(nextRequestStartedWhileWriteBlocked).IsTrue();
        await Assert.That(queuedPages.Count).IsEqualTo(2);
        await Assert.That(queuedPages[0].Projection.Values[0, 0]).IsEqualTo("001A");
        await Assert.That(queuedPages[1].Projection.Values[0, 0]).IsEqualTo("001B");
        await Assert.That(queuedPages[0].TotalRowCount).IsEqualTo(2);
        await Assert.That(queuedPages[1].TotalRowCount).IsNull();
        await Assert.That(result.Result.RecordsProcessed).IsEqualTo(2);

        releaseFirstWrite.SetResult();
        await firstWrite;
    }

    [Test]
    public async Task Paged_Query_Offsets_Second_Page_And_Formats_Only_The_First_Full_Result_Body()
    {
        var pages = new List<QueryTablePage>();
        var (client, handler) = SalesforceTestClients.Create();
        handler.Enqueue(HttpStatusCode.OK, ReadFixture("AccountDescribe.json"));
        handler.Enqueue(HttpStatusCode.OK, """
            {
              "totalSize": 2,
              "done": false,
              "nextRecordsUrl": "/services/data/v66.0/query/01g-next",
              "records": [
                { "attributes": { "type": "Account" }, "Id": "001A", "Name": "Acme", "Industry": "Technology" }
              ]
            }
            """);
        handler.Enqueue(HttpStatusCode.OK, """
            {
              "totalSize": 2,
              "done": true,
              "records": [
                { "attributes": { "type": "Account" }, "Id": "001B", "Name": "Beta", "Industry": "Finance" }
              ]
            }
            """);

        await QueryTable.RunPagedAsync(
            client,
            new QueryTableInput { Snapshot = ForceTableSnapshots.ValidAccountTableAt(8, 5, bodyRows: 5) },
            (page, _) =>
            {
                pages.Add(page);
                return Task.CompletedTask;
            });

        await Assert.That(pages.Count).IsEqualTo(2);
        await Assert.That(pages[0].ClearBody).IsNotNull();
        await Assert.That(pages[0].ClearBody!.EndRow).IsEqualTo(14);
        await Assert.That(pages[0].Projection.StartRow).IsEqualTo(10);
        await Assert.That(pages[0].Projection.StartColumn).IsEqualTo(5);
        await Assert.That(pages[0].Projection.Values.GetLength(1)).IsEqualTo(3);
        await Assert.That(pages[0].TotalRowCount).IsEqualTo(2);
        var formatEndRow = pages[0].Projection.StartRow + pages[0].TotalRowCount.GetValueOrDefault() - 1;
        var clearEndRow = pages[0].ClearBody?.EndRow ?? throw new InvalidOperationException("First page must clear the existing body.");
        await Assert.That(formatEndRow).IsEqualTo(11);
        await Assert.That(formatEndRow).IsLessThan(clearEndRow);
        await Assert.That(pages[1].ClearBody).IsNull();
        await Assert.That(pages[1].Projection.StartRow).IsEqualTo(11);
        await Assert.That(pages[1].TotalRowCount).IsNull();
    }

    [Test]
    public async Task First_Page_Total_Row_Count_Is_Never_Smaller_Than_Its_Record_Count()
    {
        var pages = new List<QueryTablePage>();
        var (client, handler) = SalesforceTestClients.Create();
        handler.Enqueue(HttpStatusCode.OK, ReadFixture("AccountDescribe.json"));
        handler.Enqueue(HttpStatusCode.OK, """
            {
              "totalSize": 1,
              "done": true,
              "records": [
                { "attributes": { "type": "Account" }, "Id": "001A", "Name": "Acme", "Industry": "Technology" },
                { "attributes": { "type": "Account" }, "Id": "001B", "Name": "Beta", "Industry": "Finance" }
              ]
            }
            """);

        await QueryTable.RunPagedAsync(
            client,
            new QueryTableInput { Snapshot = ForceTableSnapshots.ValidAccountTableAt(8, 5) },
            (page, _) =>
            {
                pages.Add(page);
                return Task.CompletedTask;
            });

        await Assert.That(pages.Count).IsEqualTo(1);
        await Assert.That(pages[0].TotalRowCount).IsEqualTo(2);
    }

    [Test]
    public async Task Cancellation_After_Queued_Page_Does_Not_Start_QueryMore()
    {
        using var cancellation = new CancellationTokenSource();
        var queuedPages = new List<QueryTablePage>();
        var (client, handler) = SalesforceTestClients.Create();
        handler.Enqueue(HttpStatusCode.OK, ReadFixture("AccountDescribe.json"));
        handler.Enqueue(HttpStatusCode.OK, """
            {
              "totalSize": 2,
              "done": false,
              "nextRecordsUrl": "/services/data/v66.0/query/01g-next",
              "records": [
                { "attributes": { "type": "Account" }, "Id": "001A", "Name": "Acme", "Industry": "Technology" }
              ]
            }
            """);

        var result = await QueryTable.RunPagedAsync(
            client,
            new QueryTableInput { Snapshot = ForceTableSnapshots.ValidAccountTable() },
            (page, _) =>
            {
                queuedPages.Add(page);
                cancellation.Cancel();
                return Task.CompletedTask;
            },
            cancellation.Token);

        await Assert.That(result.Result.WasCancelled).IsTrue();
        await Assert.That(result.Result.RecordsProcessed).IsEqualTo(1);
        await Assert.That(queuedPages.Count).IsEqualTo(1);
        await Assert.That(handler.Requests.Count).IsEqualTo(2);
    }

    [Test]
    public async Task QueryMore_Http_Failure_Marks_The_First_Unwritten_Row()
    {
        var pages = new List<QueryTablePage>();
        var (client, handler) = SalesforceTestClients.Create();
        handler.Enqueue(HttpStatusCode.OK, ReadFixture("AccountDescribe.json"));
        handler.Enqueue(HttpStatusCode.OK, """
            {
              "totalSize": 2,
              "done": false,
              "nextRecordsUrl": "/services/data/v66.0/query/01g-next",
              "records": [
                { "attributes": { "type": "Account" }, "Id": "001A", "Name": "Acme", "Industry": "Technology" }
              ]
            }
            """);
        handler.Enqueue(HttpStatusCode.ServiceUnavailable, "{ \"message\": \"temporarily unavailable\" }");

        var result = await QueryTable.RunPagedAsync(
            client,
            new QueryTableInput { Snapshot = ForceTableSnapshots.ValidAccountTableAt(8, 5) },
            (page, _) =>
            {
                pages.Add(page);
                return Task.CompletedTask;
            });

        await Assert.That(pages.Count).IsEqualTo(1);
        await Assert.That(result.Result.ErrorSummary).IsNotNull();
        await Assert.That(result.Result.Projection!.Values[0, 0]).IsEqualTo("#Err");
        await Assert.That(result.Result.Projection.StartRow).IsEqualTo(11);
        await Assert.That(result.Result.Projection.StartColumn).IsEqualTo(5);
    }

    private static async Task SimulateFirstPageWriteAsync(Task queued, Task release)
    {
        await queued;
        await release;
    }

    private static HttpResponseMessage JsonResponse(string body) => new(HttpStatusCode.OK)
    {
        Content = new StringContent(body),
    };

    private static string ReadFixture(string name)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Fixtures", name);
        return File.ReadAllText(path);
    }
}
