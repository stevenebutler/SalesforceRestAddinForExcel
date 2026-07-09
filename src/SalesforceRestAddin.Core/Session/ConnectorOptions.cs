namespace SalesforceRestAddin.Core.Session;

/// <summary>
/// Excel/data-plane behavior toggles (legacy Options dialog / RegDB).
/// Most flags default false. Update skips AutoFilter/manually hidden rows and columns
/// unless <see cref="IncludeHiddenCells"/> is enabled.
/// </summary>
public sealed class ConnectorOptions
{
    /// <summary>Default options — hidden rows/columns are excluded from update.</summary>
    public static ConnectorOptions Default { get; } = new();

    /// <summary>Resolve Salesforce names ↔ ids (<c>UseReference</c>).</summary>
    public bool UseReference { get; init; }

    /// <summary>Skip bulk-operation confirmation dialogs (<c>NoWarn</c>).</summary>
    public bool NoWarning { get; init; }

    /// <summary>Confirm large query-table downloads with a COUNT check.</summary>
    public bool ConfirmLargeQuery { get; init; }

    /// <summary>Disable row/column selection limits on query/delete (<c>NoLimit</c>).</summary>
    public bool NoQueryLimit { get; init; }

    /// <summary>
    /// Legacy <c>AutoAssignRule</c>. When false (default), create/update sends <c>Sforce-Auto-Assign: False</c>.
    /// When true, that header is omitted and Salesforce assignment rules run normally.
    /// </summary>
    public bool AutoAssignRule { get; init; }

    /// <summary>
    /// When false (default), update omits AutoFilter/manually hidden Excel rows and columns.
    /// When true, hidden cells in the selection are included in the PATCH payload.
    /// </summary>
    public bool IncludeHiddenCells { get; init; }

    /// <summary>REST composite batch size (create/update/delete/retrieve). Default 200.</summary>
    public int CompositeBatchSize { get; init; } = 200;

    /// <summary>Whether to send the legacy assignment-rule suppression header on create/update.</summary>
    public bool SendPreventAutoAssignHeader => !AutoAssignRule;
}
