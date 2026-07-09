using System;
using System.IO;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Tests;

public sealed class SessionFlowTraceTests
{
    private static readonly object TestLock = new();

    [Test]
    public async Task Session_flow_trace_writes_to_data_directory()
    {
        lock (TestLock)
        {
            RunSessionFlowTraceFileTest().GetAwaiter().GetResult();
        }
    }

    private static async Task RunSessionFlowTraceFileTest()
    {
        var logPath = SalesforceRestAddinDataPaths.SessionTraceLogFile;
        if (File.Exists(logPath))
        {
            File.Delete(logPath);
        }

        SessionFlowTrace.ResetForOperation("test-scope");
        var marker = $"hello-from-test-{Guid.NewGuid():N}";
        SessionFlowTrace.Log(marker);

        await Assert.That(File.Exists(logPath)).IsTrue();
        var text = await File.ReadAllTextAsync(logPath);
        await Assert.That(text).Contains("test-scope");
        await Assert.That(text).Contains(marker);
    }

    [Test]
    public async Task Nested_scopes_preserve_earlier_trace_lines()
    {
        lock (TestLock)
        {
            RunNestedScopesTest().GetAwaiter().GetResult();
        }
    }

    private static async Task RunNestedScopesTest()
    {
        SessionFlowTrace.ResetForOperation("Table Query Wizard");
        SessionFlowTrace.Log("REST describe Account: GET https://example.test/sobjects/Account/describe");
        SessionFlowTrace.BeginScope("Query Table Data");
        SessionFlowTrace.Log("QueryTable: object=");

        var recent = SessionFlowTrace.FormatRecent();
        await Assert.That(recent).Contains("Table Query Wizard");
        await Assert.That(recent).Contains("REST describe Account");
        await Assert.That(recent).Contains("Query Table Data");
        await Assert.That(recent).Contains("QueryTable: object=");
        await Assert.That(recent).Contains("→ Query Table Data");
    }
}
