using SalesforceRestAddin.Core.Rest;

namespace SalesforceRestAddin.Windows.Ui;

internal sealed class SObjectListItem
{
    public SObjectListItem(SObjectSummary summary) => Summary = summary;

    public SObjectSummary Summary { get; }

    public override string ToString() => $"{Summary.Label} ({Summary.Name})";
}
