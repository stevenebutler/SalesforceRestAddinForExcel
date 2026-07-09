using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Core.Tables;

namespace SalesforceRestAddin.Core.Values;

public static class FieldValueConverter
{
    public static object? ToSalesforceValue(
        FieldDescriptor field,
        object? cellValue,
        ConnectorOptions options,
        bool forCreate)
    {
        // Legacy Util.toSalesforceType: empty cell is always null (all field types).
        // Create omits nulls; update includes them so Salesforce clears the field.
        if (cellValue is null || IsEmptyCell(cellValue))
        {
            return null;
        }

        if (!IsWritable(field, forCreate))
        {
            return null;
        }

        return field.Type.ToLowerInvariant() switch
        {
            "date" or "datetime" => ToIsoDate(cellValue),
            "multipicklist" => cellValue.ToString(),
            "boolean" => ToBoolean(cellValue),
            "double" or "currency" or "percent" => ToDecimal(cellValue),
            _ => cellValue.ToString(),
        };
    }

    public static object? ToDisplayValue(
        FieldDescriptor field,
        object? salesforceValue,
        ConnectorOptions options,
        IReferenceResolver? referenceResolver = null,
        string? resolvedReferenceName = null)
    {
        if (salesforceValue is null)
        {
            return null;
        }

        if (field.IsReference && options.UseReference && resolvedReferenceName is not null)
        {
            return resolvedReferenceName;
        }

        return field.Type.ToLowerInvariant() switch
        {
            "address" or "location" => FlattenCompound(salesforceValue),
            "date" or "datetime" => ParseSalesforceDate(salesforceValue),
            "multipicklist" => salesforceValue.ToString(),
            _ => salesforceValue,
        };
    }

    public static ColumnFormatKind GetColumnFormatKind(FieldDescriptor field, object? sampleValue)
    {
        if (field.Type.Equals("id", StringComparison.OrdinalIgnoreCase)
            || (sampleValue is string text && text.Length > 0 && char.IsDigit(text[0]) == false && field.Type.Equals("string", StringComparison.OrdinalIgnoreCase)))
        {
            return ColumnFormatKind.Text;
        }

        return field.Type.ToLowerInvariant() switch
        {
            "date" => ColumnFormatKind.Date,
            "datetime" => ColumnFormatKind.DateTime,
            "double" or "currency" or "percent" or "int" => ColumnFormatKind.Number,
            "string" or "picklist" or "multipicklist" or "phone" or "url" or "email" or "textarea" => ColumnFormatKind.Text,
            _ => ColumnFormatKind.General,
        };
    }

    /// <summary>
    /// Column format for sheet projection — Excel number-format string matches legacy <c>Util.typeToFormat</c>.
    /// </summary>
    public static ColumnFormat CreateColumnFormat(FieldDescriptor field, object? sampleValue = null)
    {
        var kind = GetColumnFormatKind(field, sampleValue);
        return new ColumnFormat
        {
            Kind = kind,
            ExcelFormat = ResolveExcelNumberFormat(field.Type, kind),
        };
    }

    /// <summary>Legacy <c>Util.typeToFormat</c> mapping for column-scoped <c>NumberFormat</c>.</summary>
    public static string? ResolveExcelNumberFormat(string fieldType, ColumnFormatKind kind)
    {
        var type = fieldType?.ToLowerInvariant() ?? string.Empty;
        return type switch
        {
            "date" => "yyyy-MM-dd",
            "datetime" => "yyyy-MM-dd HH:mm:ss",
            "currency" => "$#,##0_);($#,##0)",
            "string" or "picklist" or "multipicklist" or "phone" or "url" or "email" or "textarea" or "id" => "@",
            _ => kind switch
            {
                ColumnFormatKind.Date => "yyyy-MM-dd",
                ColumnFormatKind.DateTime => "yyyy-MM-dd HH:mm:ss",
                ColumnFormatKind.Text => "@",
                _ => null,
            },
        };
    }

    public static bool IsWritable(FieldDescriptor field, bool forCreate) =>
        forCreate ? field.Createable && !field.IsId : field.Updateable && !field.IsId;

    public static bool IsEmptyCell(object? cellValue) =>
        cellValue is null
        || cellValue is DBNull
        || (cellValue is string s && string.IsNullOrEmpty(s));

    private static string? ToIsoDate(object? cellValue)
    {
        if (cellValue is DateTime dt)
        {
            return dt.ToString("yyyy-MM-dd");
        }

        if (cellValue is double oaDate)
        {
            return DateTime.FromOADate(oaDate).ToString("yyyy-MM-dd");
        }

        return cellValue?.ToString();
    }

    private static bool ToBoolean(object? cellValue) =>
        cellValue switch
        {
            bool b => b,
            string s => s.Equals("true", StringComparison.OrdinalIgnoreCase),
            _ => Convert.ToBoolean(cellValue),
        };

    private static decimal? ToDecimal(object? cellValue) =>
        cellValue switch
        {
            decimal d => d,
            double dbl => (decimal)dbl,
            _ => decimal.TryParse(cellValue?.ToString(), out var parsed) ? parsed : null,
        };

    private static object? ParseSalesforceDate(object? salesforceValue)
    {
        if (salesforceValue is DateTime dt)
        {
            return dt;
        }

        if (DateTime.TryParse(salesforceValue?.ToString(), out var parsed))
        {
            return parsed;
        }

        return salesforceValue;
    }

    private static string FlattenCompound(object? salesforceValue)
    {
        if (salesforceValue is not Dictionary<string, object?> dict)
        {
            return salesforceValue?.ToString() ?? string.Empty;
        }

        var parts = new[] { "street", "city", "state", "postalCode", "country" }
            .Select(key => dict.TryGetValue(key, out var v) ? v?.ToString() : null)
            .Where(v => !string.IsNullOrWhiteSpace(v));
        return string.Join(", ", parts);
    }
}
