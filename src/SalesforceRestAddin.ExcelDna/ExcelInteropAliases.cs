// Implicit System usings expose System.Range on current SDKs. Keep unqualified Range
// references in the Excel host bound to the COM interop type.
global using Range = Microsoft.Office.Interop.Excel.Range;
