using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Tables;

/// <summary>
/// Temporarily exposes rows hidden by worksheet or table AutoFilters, then restores
/// the captured criteria against the post-write data when disposed.
/// </summary>
internal sealed class AutoFilterWriteScope : IDisposable
{
    private readonly IReadOnlyList<FilterState> _filters;
    private bool _disposed;

    private AutoFilterWriteScope(IReadOnlyList<FilterState> filters)
    {
        _filters = filters;
    }

    public static AutoFilterWriteScope Suspend(Worksheet worksheet)
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

        foreach (var state in states)
        {
            state.Suspend();
        }

        SessionFlowTrace.Log($"Excel bulk write: suspended AutoFilters={states.Count}");
        return new AutoFilterWriteScope(states);
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
            SessionFlowTrace.Log($"Excel bulk write: reapplied AutoFilters={_filters.Count}");
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

        var criteria = new List<FilterCriterion>();
        var filters = autoFilter.Filters;
        var filterCount = filters.Count;
        for (var field = 1; field <= filterCount; field++)
        {
            var filter = filters.Item[field];
            if (!filter.On)
            {
                continue;
            }

            object criteria1;
            try
            {
                criteria1 = filter.Criteria1;
            }
            catch
            {
                criteria1 = Type.Missing;
            }

            var autoFilterOperator = filter.Operator;
            object criteria2;
            try
            {
                criteria2 = filter.Criteria2;
            }
            catch
            {
                criteria2 = Type.Missing;
            }

            if (ReferenceEquals(criteria1, Type.Missing) && ReferenceEquals(criteria2, Type.Missing))
            {
                throw new InvalidOperationException(
                    $"Could not preserve AutoFilter criteria for field {field} in {range.Address[false, false]}.");
            }

            criteria.Add(new FilterCriterion(field, criteria1, autoFilterOperator, criteria2));
        }

        if (criteria.Count > 0)
        {
            states.Add(new FilterState(autoFilter, range, criteria));
        }
    }

    private sealed class FilterState
    {
        private readonly AutoFilter _autoFilter;
        private readonly Range _range;
        private readonly IReadOnlyList<FilterCriterion> _criteria;

        public FilterState(
            AutoFilter autoFilter,
            Range range,
            IReadOnlyList<FilterCriterion> criteria)
        {
            _autoFilter = autoFilter;
            _range = range;
            _criteria = criteria;
        }

        public void Suspend() => _autoFilter.ShowAllData();

        public void Restore()
        {
            foreach (var criterion in _criteria)
            {
                if (criterion.Operator == 0 && ReferenceEquals(criterion.Criteria2, Type.Missing))
                {
                    _range.AutoFilter(criterion.Field, criterion.Criteria1);
                    continue;
                }

                _range.AutoFilter(
                    criterion.Field,
                    criterion.Criteria1,
                    criterion.Operator,
                    criterion.Criteria2,
                    Type.Missing);
            }
        }
    }

    private sealed class FilterCriterion
    {
        public FilterCriterion(
            int field,
            object criteria1,
            XlAutoFilterOperator autoFilterOperator,
            object criteria2)
        {
            Field = field;
            Criteria1 = criteria1;
            Operator = autoFilterOperator;
            Criteria2 = criteria2;
        }

        public int Field { get; }

        public object Criteria1 { get; }

        public XlAutoFilterOperator Operator { get; }

        public object Criteria2 { get; }
    }
}
