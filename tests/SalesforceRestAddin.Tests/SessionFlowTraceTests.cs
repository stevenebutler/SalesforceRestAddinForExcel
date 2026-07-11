using System;
using System.IO;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Tests;

[NotInParallel]
public sealed class SessionFlowTraceTests
{
    private static readonly object TestLock = new();

    [Test]
    public async Task Session_flow_trace_writes_to_data_directory()
    {
        lock (TestLock)
        {
            ResetTraceFiles();
            RunSessionFlowTraceFileTest().GetAwaiter().GetResult();
        }
    }

    private static async Task RunSessionFlowTraceFileTest()
    {
        var logPath = SalesforceRestAddinDataPaths.SessionTraceLogFile;
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
            ResetTraceFiles();
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

    [Test]
    public async Task Session_flow_trace_rolls_existing_log_on_first_append()
    {
        lock (TestLock)
        {
            ResetTraceFiles();
            RunRollsExistingLogTest().GetAwaiter().GetResult();
        }
    }

    private static async Task RunRollsExistingLogTest()
    {
        var logPath = SalesforceRestAddinDataPaths.SessionTraceLogFile;
        var rolledPath = Path.Combine(Path.GetDirectoryName(logPath)!, "session-trace-1.log");
        await File.WriteAllTextAsync(logPath, "legacy-current" + Environment.NewLine);

        SessionFlowTrace.ResetForTests();
        SessionFlowTrace.ResetForOperation("rollover");
        SessionFlowTrace.Log("first-write");

        await Assert.That(File.Exists(logPath)).IsTrue();
        await Assert.That(File.Exists(rolledPath)).IsTrue();
        await Assert.That(await File.ReadAllTextAsync(rolledPath)).Contains("legacy-current");
        await Assert.That(await File.ReadAllTextAsync(logPath)).Contains("first-write");
        await Assert.That(await File.ReadAllTextAsync(logPath)).DoesNotContain("legacy-current");
    }

    [Test]
    public async Task Session_flow_trace_rolls_only_once_per_session()
    {
        lock (TestLock)
        {
            ResetTraceFiles();
            RunRollsOnlyOnceTest().GetAwaiter().GetResult();
        }
    }

    private static async Task RunRollsOnlyOnceTest()
    {
        var logPath = SalesforceRestAddinDataPaths.SessionTraceLogFile;
        var rolledPath = Path.Combine(Path.GetDirectoryName(logPath)!, "session-trace-1.log");
        await File.WriteAllTextAsync(logPath, "legacy-current" + Environment.NewLine);

        SessionFlowTrace.ResetForTests();
        SessionFlowTrace.ResetForOperation("first-operation");
        SessionFlowTrace.Log("first-write");
        SessionFlowTrace.Log("second-write");

        await Assert.That(await File.ReadAllTextAsync(rolledPath)).Contains("legacy-current");
        var current = await File.ReadAllTextAsync(logPath);
        await Assert.That(current).Contains("first-write");
        await Assert.That(current).Contains("second-write");
        await Assert.That(current).DoesNotContain("legacy-current");
    }

    [Test]
    public async Task Session_flow_trace_replaces_existing_rolled_file()
    {
        lock (TestLock)
        {
            ResetTraceFiles();
            RunReplacesExistingRolledFileTest().GetAwaiter().GetResult();
        }
    }

    private static async Task RunReplacesExistingRolledFileTest()
    {
        var logPath = SalesforceRestAddinDataPaths.SessionTraceLogFile;
        var rolledPath = Path.Combine(Path.GetDirectoryName(logPath)!, "session-trace-1.log");
        await File.WriteAllTextAsync(logPath, "legacy-current" + Environment.NewLine);
        await File.WriteAllTextAsync(rolledPath, "stale-backup" + Environment.NewLine);

        SessionFlowTrace.ResetForTests();
        SessionFlowTrace.ResetForOperation("replace-backup");
        SessionFlowTrace.Log("first-write");

        var rolled = await File.ReadAllTextAsync(rolledPath);
        await Assert.That(rolled).Contains("legacy-current");
        await Assert.That(rolled).DoesNotContain("stale-backup");
        await Assert.That(await File.ReadAllTextAsync(logPath)).Contains("first-write");
    }

    private static void ResetTraceFiles()
    {
        SessionFlowTrace.ResetForTests();
        var logPath = SalesforceRestAddinDataPaths.SessionTraceLogFile;
        var rolledPath = Path.Combine(Path.GetDirectoryName(logPath)!, "session-trace-1.log");
        if (File.Exists(logPath))
        {
            File.Delete(logPath);
        }

        if (File.Exists(rolledPath))
        {
            File.Delete(rolledPath);
        }
    }
}
