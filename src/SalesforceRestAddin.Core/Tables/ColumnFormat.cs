namespace SalesforceRestAddin.Core.Tables;

public enum ColumnFormatKind
{
    General,
    Text,
    Date,
    DateTime,
    Number,
}

public sealed class ColumnFormat
{
    public ColumnFormatKind Kind { get; init; } = ColumnFormatKind.General;

    public string? ExcelFormat { get; init; }
}
