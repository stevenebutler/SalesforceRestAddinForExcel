using System.Linq;
using SalesforceRestAddin.Core.Session;
using SalesforceRestAddin.Core.Tables;
using SalesforceRestAddin.Core.Values;

namespace SalesforceRestAddin.Core.Soql;

public static class SoqlEscape
{
    public static string EscapeLiteral(string value)
    {
        return value
            .Replace("\\", "\\\\")
            .Replace("'", "\\'");
    }
}

public static class SoqlCriteriaParser
{
    public static async Task<SoqlCriteriaParseResult> ParseAsync(
        object?[] criteriaRow,
        FieldCatalog catalog,
        ConnectorOptions options,
        IReferenceResolver? referenceResolver,
        CancellationToken cancellationToken = default,
        IReadOnlyDictionary<int, IReadOnlyList<string>>? criteriaReferenceIds = null)
    {
        if (criteriaRow is null)
        {
            throw new ArgumentNullException(nameof(criteriaRow));
        }

        if (catalog is null)
        {
            throw new ArgumentNullException(nameof(catalog));
        }

        var errors = new List<SoqlCriteriaError>();
        var clauses = new List<string>();
        string? joinField = null;
        List<string>? joinIds = null;
        var joinMode = false;

        for (var i = 0; i + 2 < criteriaRow.Length; i += 3)
        {
            var fieldCell = criteriaRow[i];
            var operatorCell = criteriaRow[i + 1];
            var valueCell = criteriaRow[i + 2];

            var fieldLabel = fieldCell?.ToString();
            if (string.IsNullOrWhiteSpace(fieldLabel))
            {
                break;
            }

            var field = catalog.ResolveHeaderField(fieldLabel, null);
            if (field is null)
            {
                errors.Add(new SoqlCriteriaError
                {
                    Cell = new CellRef(1, i + 2),
                    Message = $"Unknown criteria field '{fieldLabel}'.",
                });
                break;
            }

            var rawOperator = operatorCell?.ToString() ?? string.Empty;
            var normalizedOperator = NormalizeOperator(rawOperator);
            var value = valueCell?.ToString() ?? string.Empty;

            if (string.Equals(normalizedOperator, "like", StringComparison.OrdinalIgnoreCase)
                && string.Equals(field.Type, "picklist", StringComparison.OrdinalIgnoreCase))
            {
                errors.Add(new SoqlCriteriaError
                {
                    Cell = new CellRef(1, i + 3),
                    Message = "like (or contains) operator is not valid on picklist fields; use equals or not equals.",
                });
                return new SoqlCriteriaParseResult { Errors = errors };
            }

            if (normalizedOperator == "on")
            {
                errors.Add(new SoqlCriteriaError
                {
                    Cell = new CellRef(1, i + 3),
                    Message = "ON is not supported - use IN with a range reference to select multiple items.",
                });
                return new SoqlCriteriaParseResult { Errors = errors };
            }

            if (normalizedOperator == "in")
            {
                if (!field.IsReference && !field.IsId)
                {
                    errors.Add(new SoqlCriteriaError
                    {
                        Cell = new CellRef(1, i + 2),
                        Message = $"{normalizedOperator} is only valid on Id or reference fields.",
                    });
                    return new SoqlCriteriaParseResult { Errors = errors };
                }

                if (criteriaReferenceIds is null
                    || !criteriaReferenceIds.TryGetValue(i + 2, out var ids)
                    || ids.Count == 0)
                {
                    errors.Add(new SoqlCriteriaError
                    {
                        Cell = new CellRef(1, i + 3),
                        Message = "Reference criteria must name an Excel range or named range containing Salesforce record Ids.",
                    });
                    return new SoqlCriteriaParseResult { Errors = errors };
                }

                joinField = field.Name;
                joinIds = ids.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                joinMode = true;
                continue;
            }

            var (clause, dateError) = await BuildClauseAsync(
                field,
                normalizedOperator,
                rawOperator,
                value,
                valueCell,
                options,
                referenceResolver,
                cancellationToken).ConfigureAwait(false);

            if (dateError is not null)
            {
                errors.Add(new SoqlCriteriaError
                {
                    Cell = new CellRef(1, i + 3),
                    Message = dateError,
                });
                return new SoqlCriteriaParseResult { Errors = errors };
            }

            clauses.Add(clause!);
        }

        if (errors.Count > 0)
        {
            return new SoqlCriteriaParseResult { Errors = errors };
        }

        var where = clauses.Count == 0
            ? string.Empty
            : string.Join(" and ", clauses);

        return new SoqlCriteriaParseResult
        {
            WhereClause = where,
            ReferenceJoinField = joinField,
            ReferenceJoinIds = joinIds,
            IsReferenceJoinMode = joinMode,
        };
    }

    private static string NormalizeOperator(string raw)
    {
        return raw.Trim().ToLowerInvariant() switch
        {
            "equals" => "=",
            "contains" => "like",
            "not equals" => "!=",
            "less than" => "<",
            "greater than" => ">",
            "begins with" or "starts with" => "begins with",
            "ends with" => "ends with",
            "regexp" => "like",
            "in" => "in",
            _ => raw.Trim().ToLowerInvariant(),
        };
    }

    private static async Task<(string? Clause, string? Error)> BuildClauseAsync(
        FieldDescriptor field,
        string normalizedOperator,
        string rawOperator,
        string value,
        object? rawValue,
        ConnectorOptions options,
        IReferenceResolver? referenceResolver,
        CancellationToken cancellationToken)
    {
        var type = field.Type.ToLowerInvariant();
        var isDateType = type is "date" or "datetime";

        // Excel Value2 for date cells is often a DateTime or OADate double — convert before
        // stringifying so we never emit serials like 44368 into SOQL.
        if (isDateType && TryConvertExcelDateCell(rawValue, out var excelDate))
        {
            return ($"{field.Name} {normalizedOperator} {FormatDateTimeLiteral(type, excelDate)}", null);
        }

        if (string.IsNullOrEmpty(value))
        {
            if (isDateType)
            {
                return ($"{field.Name} {normalizedOperator} null", null);
            }

            return ($"{field.Name} {normalizedOperator} ''", null);
        }

        if (isDateType)
        {
            // Comma-separated OR groups for dates: parse each part independently.
            var parts = value.Split(',')
                .Select(p => p.Trim())
                .Where(p => p.Length > 0)
                .ToArray();
            var subClauses = new List<string>();
            foreach (var part in parts)
            {
                if (IsRelativeSoqlDateLiteral(part))
                {
                    subClauses.Add($"{field.Name} {normalizedOperator} {part}");
                    continue;
                }

                if (!TryParseCriteriaDateString(part, type, out var dt, out var error))
                {
                    return (null, error);
                }

                subClauses.Add($"{field.Name} {normalizedOperator} {FormatDateTimeLiteral(type, dt)}");
            }

            return (subClauses.Count == 1 ? subClauses[0] : $"({string.Join(" or ", subClauses)})", null);
        }

        var textParts = value.Split(',')
            .Select(p => p.Trim())
            .Where(p => p.Length > 0)
            .ToArray();
        var textSubClauses = new List<string>();
        foreach (var part in textParts)
        {
            var resolved = await ResolveValueAsync(field, part, options, referenceResolver, cancellationToken)
                .ConfigureAwait(false);
            var formatted = FormatValue(field, normalizedOperator, rawOperator, resolved);
            textSubClauses.Add($"{field.Name} {GetEffectiveOperator(field, normalizedOperator, rawOperator)} {formatted}");
        }

        return (textSubClauses.Count == 1 ? textSubClauses[0] : $"({string.Join(" or ", textSubClauses)})", null);
    }

    private static string GetEffectiveOperator(FieldDescriptor field, string normalizedOperator, string rawOperator)
    {
        if (string.Equals(field.Type, "multipicklist", StringComparison.OrdinalIgnoreCase))
        {
            var negated = normalizedOperator is "!=" or "not equals" or "excludes"
                || rawOperator.Equals("not equals", StringComparison.OrdinalIgnoreCase)
                || rawOperator.Equals("excludes", StringComparison.OrdinalIgnoreCase);
            return negated ? "excludes" : "includes";
        }

        if (normalizedOperator is "begins with" or "ends with")
        {
            return "like";
        }

        return normalizedOperator;
    }

    private static string FormatValue(
        FieldDescriptor field,
        string normalizedOperator,
        string rawOperator,
        string resolvedValue)
    {
        var type = field.Type.ToLowerInvariant();

        // Numeric / boolean / date literals must not be quoted (Salesforce INVALID_FIELD otherwise).
        if (type is "double" or "currency" or "percent")
        {
            return FormatNumericLiteral(resolvedValue, forceDecimal: true);
        }

        if (type is "int")
        {
            return FormatNumericLiteral(resolvedValue, forceDecimal: false);
        }

        if (type is "boolean")
        {
            return FormatBooleanLiteral(resolvedValue);
        }

        var escaped = SoqlEscape.EscapeLiteral(resolvedValue);
        var quoted = $"'{escaped}'";

        if (type == "multipicklist")
        {
            return $"({quoted})";
        }

        if (normalizedOperator == "like"
            || rawOperator.Equals("contains", StringComparison.OrdinalIgnoreCase))
        {
            return $"'%{escaped}%'";
        }

        if (normalizedOperator == "begins with")
        {
            return $"'{escaped}%'";
        }

        if (normalizedOperator == "ends with")
        {
            return $"'%{escaped}'";
        }

        return quoted;
    }

    /// <summary>
    /// Legacy <c>Util.QueryValueFormat</c> for double/currency/percent/int — unquoted SOQL numbers.
    /// </summary>
    private static string FormatNumericLiteral(string value, bool forceDecimal)
    {
        var trimmed = value.Trim();
        if (!double.TryParse(
                trimmed,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out var number)
            && !double.TryParse(
                trimmed,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.CurrentCulture,
                out number))
        {
            // Fall back to raw text so Salesforce returns a clear type error rather than quoting.
            return trimmed;
        }

        if (!forceDecimal)
        {
            return ((long)Math.Truncate(number)).ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        var invariant = number.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (invariant.IndexOf('.') < 0)
        {
            return invariant + ".0";
        }

        return invariant;
    }

    private static string FormatBooleanLiteral(string value)
    {
        var trimmed = value.Trim();
        if (bool.TryParse(trimmed, out var b))
        {
            return b ? "TRUE" : "FALSE";
        }

        if (trimmed == "1" || trimmed.Equals("true", StringComparison.OrdinalIgnoreCase))
        {
            return "TRUE";
        }

        if (trimmed == "0" || trimmed.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            return "FALSE";
        }

        // Non-zero numeric → TRUE (legacy Val(vlu) OR "true").
        if (double.TryParse(trimmed, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var n))
        {
            return n != 0 ? "TRUE" : "FALSE";
        }

        return "FALSE";
    }

    /// <summary>
    /// Parses a date/datetime criteria string: ISO (T or space) or Excel OADate serial text.
    /// Locale-formatted text (e.g. 26/07/2021) is rejected — use a real Excel date cell,
    /// ISO, or a relative SOQL literal (TODAY, LAST_N_DAYS, …).
    /// </summary>
    internal static bool TryParseCriteriaDateString(
        string value,
        string fieldType,
        out DateTime dateTime,
        out string? error)
    {
        dateTime = default;
        error = null;
        var trimmed = value.Trim();
        var type = fieldType.ToLowerInvariant();
        var includeTime = type == "datetime";

        if (IsRelativeSoqlDateLiteral(trimmed))
        {
            error = null;
            return false;
        }

        if (TryParseIsoDateTime(trimmed, includeTime, out dateTime))
        {
            return true;
        }

        if (TryParseExcelOaDateSerial(trimmed, out dateTime))
        {
            return true;
        }

        error =
            $"Invalid {(includeTime ? "datetime" : "date")} value '{trimmed}'. " +
            "Use an Excel date cell, an ISO value (yyyy-MM-dd" +
            (includeTime ? "[T| ]HH:mm:ss" : string.Empty) +
            "), or a relative SOQL literal (TODAY, LAST_N_DAYS, …).";
        return false;
    }

    internal static bool IsRelativeSoqlDateLiteral(string trimmed) =>
        trimmed.StartsWith("today", StringComparison.OrdinalIgnoreCase)
        || trimmed.StartsWith("yesterday", StringComparison.OrdinalIgnoreCase)
        || trimmed.StartsWith("tomorrow", StringComparison.OrdinalIgnoreCase)
        || trimmed.StartsWith("last_", StringComparison.OrdinalIgnoreCase)
        || trimmed.StartsWith("this_", StringComparison.OrdinalIgnoreCase)
        || trimmed.StartsWith("next_", StringComparison.OrdinalIgnoreCase);

    private static bool TryParseIsoDateTime(string text, bool includeTime, out DateTime dateTime)
    {
        dateTime = default;

        // Date-only ISO.
        if (DateTime.TryParseExact(
                text,
                "yyyy-MM-dd",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None,
                out dateTime))
        {
            return true;
        }

        if (!includeTime && text.Length == 10)
        {
            return false;
        }

        // ISO datetime with T or space separator.
        string[] formats =
        [
            "yyyy-MM-dd'T'HH:mm:ss.fff'Z'",
            "yyyy-MM-dd'T'HH:mm:ss'Z'",
            "yyyy-MM-dd'T'HH:mm:ss.fff",
            "yyyy-MM-dd'T'HH:mm:ss",
            "yyyy-MM-dd'T'HH:mm",
            "yyyy-MM-dd HH:mm:ss.fff",
            "yyyy-MM-dd HH:mm:ss",
            "yyyy-MM-dd HH:mm",
            "yyyy-MM-dd'T'HH:mm:ss.fffzzz",
            "yyyy-MM-dd'T'HH:mm:sszzz",
        ];

        if (DateTime.TryParseExact(
                text,
                formats,
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.AllowWhiteSpaces
                | System.Globalization.DateTimeStyles.AssumeUniversal
                | System.Globalization.DateTimeStyles.AdjustToUniversal,
                out dateTime))
        {
            return true;
        }

        // Round-trip ISO without exact format list (still requires recognizable ISO shape).
        if (LooksLikeIsoDateTime(text)
            && DateTime.TryParse(
                text,
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.RoundtripKind
                | System.Globalization.DateTimeStyles.AllowWhiteSpaces,
                out dateTime))
        {
            return true;
        }

        return false;
    }

    private static bool LooksLikeIsoDateTime(string text) =>
        text.Length >= 10
        && char.IsDigit(text[0])
        && text[4] == '-'
        && text[7] == '-'
        && (text.Length == 10 || text[10] is 'T' or ' ');

    /// <summary>
    /// Excel date cells arrive as <see cref="DateTime"/> or OADate <see cref="double"/> via Value2.
    /// </summary>
    internal static bool TryConvertExcelDateCell(object? rawValue, out DateTime dateTime)
    {
        dateTime = default;
        if (rawValue is null || rawValue is string)
        {
            // Strings go through TryParseCriteriaDateString (ISO / OADate serial only).
            return false;
        }

        if (rawValue is DateTime dt)
        {
            dateTime = dt;
            return true;
        }

        if (rawValue is double oaDouble)
        {
            return TryFromOaDate(oaDouble, out dateTime);
        }

        if (rawValue is float oaFloat)
        {
            return TryFromOaDate(oaFloat, out dateTime);
        }

        if (rawValue is decimal oaDecimal)
        {
            return TryFromOaDate((double)oaDecimal, out dateTime);
        }

        // Boxed Excel numbers sometimes arrive as int for whole-day dates.
        if (rawValue is int oaInt)
        {
            return TryFromOaDate(oaInt, out dateTime);
        }

        if (rawValue is long oaLong)
        {
            return TryFromOaDate(oaLong, out dateTime);
        }

        return false;
    }

    private static bool TryParseExcelOaDateSerial(string text, out DateTime dateTime)
    {
        dateTime = default;
        if (!double.TryParse(
                text,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out var serial)
            && !double.TryParse(
                text,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.CurrentCulture,
                out serial))
        {
            return false;
        }

        // Plain integers/decimals that look like years (e.g. 2024) must not become OADates.
        if (text.IndexOf('-') >= 0 || text.IndexOf('/') >= 0 || text.IndexOf('T') >= 0 || text.IndexOf(' ') >= 0)
        {
            return false;
        }

        // Heuristic: treat as OADate when it looks like an Excel serial (not a 4-digit year alone).
        if (serial is >= 1 and <= 2958465 && !(serial is >= 1900 and <= 2100 && Math.Abs(serial % 1) < double.Epsilon))
        {
            return TryFromOaDate(serial, out dateTime);
        }

        return false;
    }

    private static bool TryFromOaDate(double serial, out DateTime dateTime)
    {
        dateTime = default;
        try
        {
            dateTime = DateTime.FromOADate(serial);
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }

    private static string FormatDateTimeLiteral(string type, DateTime dt) =>
        type == "datetime"
            ? dt.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", System.Globalization.CultureInfo.InvariantCulture)
            : dt.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);

    private static async Task<string> ResolveValueAsync(
        FieldDescriptor field,
        string value,
        ConnectorOptions options,
        IReferenceResolver? referenceResolver,
        CancellationToken cancellationToken)
    {
        if (options.UseReference
            && field.IsReference
            && referenceResolver is not null
            && field.ReferenceTo.Count > 0)
        {
            var id = await referenceResolver
                .ResolveNameToIdAsync(field.ReferenceTo[0], value, cancellationToken)
                .ConfigureAwait(false);
            if (id is not null && !string.IsNullOrEmpty(id))
            {
                return id;
            }
        }

        return value;
    }

}
