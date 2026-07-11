namespace SalesforceRestAddin.Core.Session;

/// <summary>Controls result-row height and wrapping after Salesforce data is written.</summary>
public enum RowSizingMode
{
    /// <summary>Disable wrapping and use the worksheet standard height.</summary>
    ForceSingleLine,

    /// <summary>AutoFit each written page and cap long rows at three standard heights.</summary>
    FitEachPage,

    /// <summary>Leave existing row height and wrapping unchanged.</summary>
    None,
}
