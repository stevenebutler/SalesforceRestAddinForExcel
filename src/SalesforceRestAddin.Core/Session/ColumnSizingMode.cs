namespace SalesforceRestAddin.Core.Session;

/// <summary>Controls how result columns are AutoFit after Salesforce data is written.</summary>
public enum ColumnSizingMode
{
    /// <summary>Fit the header and the first downloaded page only.</summary>
    FirstDownloadedPage,

    /// <summary>Fit each page, retaining any wider width found on an earlier page.</summary>
    AllDownloadedData,

    /// <summary>Fit field headers without considering downloaded body values.</summary>
    HeadersOnly,
}
