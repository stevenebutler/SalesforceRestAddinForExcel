using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Tables;

/// <summary>
/// Keeps worksheet and table AutoFilter criteria alive while values are written,
/// then asks Excel to evaluate those filters against the post-write data.
/// </summary>
internal sealed class AutoFilterWriteScope : IDisposable
{
    private readonly IReadOnlyList<FilterState> _filters;
    private bool _disposed;

    private AutoFilterWriteScope(IReadOnlyList<FilterState> filters)
    {
        _filters = filters;
    }

    public static AutoFilterWriteScope Preserve(Worksheet worksheet)
    {
        if (worksheet is null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }

        var states = new List<FilterState>();
        var ranges = new HashSet<string>(StringComparer.Ordinal);

        Capture(worksheet.AutoFilter, states, ranges);

        var tables = worksheet.ListObjects;
        var tableCount = tables.Count;
        for (var index = 1; index <= tableCount; index++)
        {
            var table = tables.Item[index];
            Capture(table.AutoFilter, states, ranges);
        }

        SessionFlowTrace.Log($"Excel bulk write: preserved active AutoFilters={states.Count}");
        return new AutoFilterWriteScope(states);
    }

    internal static bool HasActiveFilters(Worksheet worksheet)
    {
        if (worksheet.AutoFilter is AutoFilter worksheetFilter
            && HasActiveFilter(worksheetFilter))
        {
            return true;
        }

        var tables = worksheet.ListObjects;
        var tableCount = tables.Count;
        for (var index = 1; index <= tableCount; index++)
        {
            if (HasActiveFilter(tables.Item[index].AutoFilter))
            {
                return true;
            }
        }

        return false;
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        foreach (var filter in _filters)
        {
            filter.Restore();
        }

        if (_filters.Count > 0)
        {
            SessionFlowTrace.Log($"Excel bulk write: re-evaluated AutoFilters={_filters.Count}");
        }
    }

    private static void Capture(
        AutoFilter? autoFilter,
        ICollection<FilterState> states,
        ISet<string> ranges)
    {
        if (autoFilter is null)
        {
            return;
        }

        var range = autoFilter.Range;
        var rangeKey = $"{range.Row}:{range.Column}:{range.Rows.Count}:{range.Columns.Count}";
        if (!ranges.Add(rangeKey))
        {
            return;
        }

        if (HasActiveFilter(autoFilter))
        {
            states.Add(new FilterState(autoFilter));
        }
    }

    private static bool HasActiveFilter(AutoFilter autoFilter)
    {
        var filters = autoFilter.Filters;
        var filterCount = filters.Count;
        for (var field = 1; field <= filterCount; field++)
        {
            if (filters.Item[field].On)
            {
                return true;
            }
        }

        return false;
    }

    private sealed class FilterState
    {
        private readonly AutoFilter _autoFilter;

        public FilterState(AutoFilter autoFilter)
        {
            _autoFilter = autoFilter;
        }

        public void Restore() => _autoFilter.ApplyFilter();
    }
}
