using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Append-only diagnostic trace (sign-in, REST, SOQL) for spike UI and Excel host.
/// </summary>
public static class SessionFlowTrace
{
    private static readonly object Sync = new();
    private static readonly List<string> Buffer = new();
    private static readonly List<string> ScopeStack = new();
    private const int MaxBufferLines = 300;

    /// <summary>
    /// Starts a new top-level ribbon/COM operation trace (clears in-memory buffer).
    /// </summary>
    public static void ResetForOperation(string operationName)
    {
        lock (Sync)
        {
            Buffer.Clear();
            ScopeStack.Clear();
            BeginScopeUnlocked(operationName);
        }
    }

    /// <summary>
    /// Marks a nested phase within the current operation without discarding earlier trace lines.
    /// </summary>
    public static void BeginScope(string scope)
    {
        lock (Sync)
        {
            BeginScopeUnlocked(scope);
        }
    }

    public static void Log(string message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return;
        }

        lock (Sync)
        {
            WriteUnlocked(message);
        }
    }

    public static void LogException(string context, Exception exception)
    {
        if (exception is null)
        {
            return;
        }

        Log($"{context}: {ExceptionChainText.Format(exception)}");
    }

    public static string FormatRecent()
    {
        lock (Sync)
        {
            if (Buffer.Count == 0)
            {
                return "(no session trace recorded)";
            }

            var builder = new StringBuilder();
            builder.AppendLine($"Session trace log: {SalesforceRestAddinDataPaths.SessionTraceLogFile}");
            if (ScopeStack.Count > 0)
            {
                builder.AppendLine($"Scope: {string.Join(" → ", ScopeStack)}");
            }

            builder.AppendLine();
            foreach (var line in Buffer)
            {
                builder.AppendLine(line);
            }

            return builder.ToString().TrimEnd();
        }
    }

    private static void BeginScopeUnlocked(string scope)
    {
        ScopeStack.Add(scope);
        WriteUnlocked($"── {scope} — begin");
    }

    private static void WriteUnlocked(string message)
    {
        var line = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss.fff} {message}";
        Buffer.Add(line);
        TrimBufferIfNeeded();

        try
        {
            SalesforceRestAddinDataPaths.EnsureDataDirectory();
            File.AppendAllText(SalesforceRestAddinDataPaths.SessionTraceLogFile, line + Environment.NewLine);
        }
        catch
        {
            // Trace must not break sign-in; in-memory buffer still available for error UI.
        }
    }

    private static void TrimBufferIfNeeded()
    {
        while (Buffer.Count > MaxBufferLines)
        {
            Buffer.RemoveAt(0);
        }
    }
}
