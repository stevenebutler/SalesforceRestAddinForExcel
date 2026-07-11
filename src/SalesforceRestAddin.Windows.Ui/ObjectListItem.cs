using SalesforceRestAddin.Core.Rest;

namespace SalesforceRestAddin.Windows.Ui;

internal sealed class ObjectListItem
{
    public ObjectListItem(SObjectSummary summary) => Summary = summary;

    public SObjectSummary Summary { get; }

    public string Label => Summary.Label;

    public string ApiName => Summary.Name;

    public override string ToString() => $"{Label} ({ApiName})";
}
