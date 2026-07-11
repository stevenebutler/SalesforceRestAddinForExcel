using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Rest;

namespace SalesforceRestAddin.Windows.Ui;

internal sealed class ObjectPickerControl : UserControl
{
    private readonly IReadOnlyList<SObjectSummary> _objects;
    private readonly ObjectPickerSortState _sortState = new();
    private readonly ListView _list;
    private readonly TextBox _filter;
    private readonly CheckBox _standardObjectsCheck;
    private readonly CheckBox _customObjectsCheck;
    private readonly CheckBox _systemObjectsCheck;
    private readonly SelectionMode _selectionMode;
    private bool _suppressSelectionEvents;
    private readonly HashSet<string> _selectedApiNames = new(StringComparer.OrdinalIgnoreCase);

    public ObjectPickerControl(
        IReadOnlyList<SObjectSummary> objects,
        SelectionMode selectionMode = SelectionMode.Single,
        bool showStandardObjects = true,
        bool showCustomObjects = true,
        bool showSystemObjects = false)
    {
        _objects = objects;
        _selectionMode = selectionMode;

        _filter = new TextBox { Margin = new Thickness(0, 0, 0, 8), MinWidth = 280 };
        _filter.TextChanged += (_, _) => RefreshList();

        _standardObjectsCheck = new CheckBox { Content = "Standard Objects", IsChecked = showStandardObjects, Margin = new Thickness(0, 0, 16, 0) };
        _customObjectsCheck = new CheckBox { Content = "Custom Objects", IsChecked = showCustomObjects, Margin = new Thickness(0, 0, 16, 0) };
        _systemObjectsCheck = new CheckBox { Content = "System Objects", IsChecked = showSystemObjects };
        _standardObjectsCheck.Checked += (_, _) => RefreshList();
        _standardObjectsCheck.Unchecked += (_, _) => RefreshList();
        _customObjectsCheck.Checked += (_, _) => RefreshList();
        _customObjectsCheck.Unchecked += (_, _) => RefreshList();
        _systemObjectsCheck.Checked += (_, _) => RefreshList();
        _systemObjectsCheck.Unchecked += (_, _) => RefreshList();

        _list = new ListView
        {
            SelectionMode = selectionMode,
            BorderBrush = Brushes.LightGray,
            BorderThickness = new Thickness(1),
            ItemContainerStyle = UiListBox.CreatePaddedItemStyle(),
            View = CreateGridView(),
        };
        _list.SelectionChanged += OnSelectionChanged;
        _list.AddHandler(GridViewColumnHeader.ClickEvent, new RoutedEventHandler(OnColumnHeaderClick));

        var root = new DockPanel();
        var filterPanel = CreateFilterPanel();
        DockPanel.SetDock(_filter, Dock.Top);
        DockPanel.SetDock(filterPanel, Dock.Top);
        root.Children.Add(_filter);
        root.Children.Add(filterPanel);
        root.Children.Add(_list);
        Content = root;

        RefreshList();
    }

    public IReadOnlyList<string> SelectedApiNames =>
        _objects
            .Where(o => _selectedApiNames.Contains(o.Name))
            .Select(o => o.Name)
            .ToList();

    public IReadOnlyList<SObjectSummary> SelectedObjects =>
        _objects
            .Where(o => _selectedApiNames.Contains(o.Name))
            .ToList();

    public string? SelectedApiName => SelectedApiNames.FirstOrDefault();

    public void SelectApiNames(IEnumerable<string>? apiNames)
    {
        _selectedApiNames.Clear();
        if (apiNames is not null)
        {
            foreach (var apiName in apiNames)
            {
                if (!string.IsNullOrWhiteSpace(apiName))
                {
                    _selectedApiNames.Add(apiName);
                }
            }
        }

        RestoreSelection();
    }

    public void RefreshList()
    {
        var filterText = _filter.Text.Trim();
        var showStandard = _standardObjectsCheck.IsChecked == true;
        var showCustom = _customObjectsCheck.IsChecked == true;
        var showSystem = _systemObjectsCheck.IsChecked == true;

        var rows = WizardObjectListFilter.BuildVisibleRows(_objects, showStandard, showCustom, showSystem);
        if (!string.IsNullOrWhiteSpace(filterText))
        {
            rows = rows
                .Where(row =>
                    row.Label.Contains(filterText, StringComparison.OrdinalIgnoreCase)
                    || row.ApiName.Contains(filterText, StringComparison.OrdinalIgnoreCase)
                    || row.Group.Contains(filterText, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        rows = _sortState.Sort(rows);

        _suppressSelectionEvents = true;
        try
        {
            _list.ItemsSource = rows;
            RestoreSelection();
        }
        finally
        {
            _suppressSelectionEvents = false;
        }
    }

    private static GridView CreateGridView()
    {
        return new GridView
        {
            AllowsColumnReorder = false,
            Columns =
            {
                new GridViewColumn
                {
                    Header = "Group",
                    DisplayMemberBinding = new Binding(nameof(ObjectPickerRow.Group)),
                    Width = 120,
                },
                new GridViewColumn
                {
                    Header = "Label",
                    DisplayMemberBinding = new Binding(nameof(ObjectPickerRow.Label)),
                    Width = 280,
                },
                new GridViewColumn
                {
                    Header = "API Name",
                    DisplayMemberBinding = new Binding(nameof(ObjectPickerRow.ApiName)),
                    Width = 260,
                },
            },
        };
    }

    private StackPanel CreateFilterPanel()
    {
        var filters = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Margin = new Thickness(0, 0, 0, 8),
        };

        filters.Children.Add(_standardObjectsCheck);
        filters.Children.Add(_customObjectsCheck);
        filters.Children.Add(_systemObjectsCheck);
        return filters;
    }

    private void OnSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (_suppressSelectionEvents)
        {
            return;
        }

        if (_selectionMode == SelectionMode.Single)
        {
            _selectedApiNames.Clear();
            if (_list.SelectedItem is ObjectPickerRow selected)
            {
                _selectedApiNames.Add(selected.ApiName);
            }

            return;
        }

        _selectedApiNames.Clear();
        foreach (var item in _list.SelectedItems.OfType<ObjectPickerRow>())
        {
            _selectedApiNames.Add(item.ApiName);
        }
    }

    private void RestoreSelection()
    {
        if (_list.ItemsSource is not IEnumerable<ObjectPickerRow> rows)
        {
            return;
        }

        if (_selectedApiNames.Count == 0)
        {
            return;
        }

        _suppressSelectionEvents = true;
        try
        {
            if (_selectionMode == SelectionMode.Single)
            {
                _list.SelectedItem = rows.FirstOrDefault(row => _selectedApiNames.Contains(row.ApiName));
                return;
            }

            _list.SelectedItems.Clear();

            foreach (var row in rows)
            {
                if (!_selectedApiNames.Contains(row.ApiName))
                {
                    continue;
                }

                _list.SelectedItems.Add(row);
            }
        }
        finally
        {
            _suppressSelectionEvents = false;
        }
    }

    private void OnColumnHeaderClick(object sender, RoutedEventArgs e)
    {
        if (e.OriginalSource is not GridViewColumnHeader header)
        {
            return;
        }

        var column = header.Content?.ToString();
        if (string.IsNullOrWhiteSpace(column))
        {
            return;
        }

        switch (column)
        {
            case "Group":
                _sortState.Promote(ObjectPickerSortColumn.Group);
                break;
            case "Label":
                _sortState.Promote(ObjectPickerSortColumn.Label);
                break;
            case "API Name":
                _sortState.Promote(ObjectPickerSortColumn.ApiName);
                break;
            default:
                return;
        }

        RefreshList();
    }
}
