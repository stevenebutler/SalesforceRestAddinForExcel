using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Media;
using SalesforceRestAddin.Core.DataPlane;
using SalesforceRestAddin.Core.Rest;
using SalesforceRestAddin.Core.Tables;

namespace SalesforceRestAddin.Windows.Ui;

public sealed class TableQueryWizardResult
{
    public required SObjectDescribe Describe { get; init; }

    public required IReadOnlyList<FieldDescriptor> Fields { get; init; }

    public required IReadOnlyList<WizardCriteriaClause> Criteria { get; init; }

    public int AnchorRow { get; init; }

    public int AnchorColumn { get; init; }
}

/// <summary>
/// Three-step WPF wizard after Excel's anchor InputBox: object → fields → criteria.
/// Destination cell is shown at the top of every step (bugs.md #1).
/// </summary>
public sealed class TableQueryWizardWindow : Window
{
    private const int StepObject = 1;
    private const int StepFields = 2;
    private const int StepCriteria = 3;
    private const int LastStep = StepCriteria;

    private static readonly string[] Operators =
    [
        "equals", "not equals", "contains", "begins with", "ends with", "less than", "greater than", "in",
    ];

    private readonly IReadOnlyList<SObjectSummary> _objects;
    private readonly Func<string, Task<SObjectDescribe>> _describeLoader;
    private readonly int _anchorRow;
    private readonly int _anchorColumn;
    private readonly string _destinationLabel;
    private readonly Grid _content;
    private readonly Button _back;
    private readonly Button _next;
    private readonly Button _cancel;
    private int _step;
    private string? _selectedObjectApiName;
    private string? _committedObjectApiName;
    private ObjectPickerControl? _objectPicker;
    private SObjectDescribe? _describe;
    private FieldListItem? _idFieldItem;
    private readonly HashSet<string> _selectedFieldNames = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<WizardCriteriaDraft> _criteriaDrafts = new();
    private readonly FieldPickerSortState _fieldSortState = new();
    private bool _suppressSelectionEvents;
    private ListView? _fieldGrid;
    private readonly List<FieldListItem> _fieldRows = new();
    private StackPanel? _criteriaPanel;
    private readonly List<(ComboBox Field, ComboBox Operator, TextBox Value)> _criteriaRows = new();

    private sealed record WizardCriteriaDraft(string FieldApiName, string? Operator, string? Value);

    public TableQueryWizardWindow(
        IReadOnlyList<SObjectSummary> objects,
        Func<string, Task<SObjectDescribe>> describeLoader,
        int anchorRow,
        int anchorColumn)
    {
        _objects = objects;
        _describeLoader = describeLoader;
        _anchorRow = anchorRow;
        _anchorColumn = anchorColumn;
        _destinationLabel = FormatA1(anchorRow, anchorColumn);

        Title = "Table Query Wizard";
        Width = 720;
        Height = 480;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;

        _content = new Grid { Margin = new Thickness(16) };
        _back = new Button { Content = "Back", Width = 80, Margin = new Thickness(0, 0, 8, 0) };
        _next = new Button { Content = "Next", Width = 100 };
        _cancel = new Button { Content = "Cancel", Width = 80, Margin = new Thickness(8, 0, 0, 0), IsCancel = true };
        _back.Click += (_, _) => MoveStep(-1);
        _next.Click += async (_, _) => await OnNextAsync();
        _cancel.Click += (_, _) => { DialogResult = false; Close(); };

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(16),
        };
        buttons.Children.Add(_back);
        buttons.Children.Add(_next);
        buttons.Children.Add(_cancel);

        var root = new Grid();
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        Grid.SetRow(_content, 0);
        Grid.SetRow(buttons, 1);
        root.Children.Add(_content);
        root.Children.Add(buttons);
        Content = root;

        ShowStep(StepObject);
    }

    public TableQueryWizardResult GetResult() =>
        new()
        {
            Describe = _describe!,
            Fields = GetSelectedFields(),
            Criteria = BuildCriteria(),
            AnchorRow = _anchorRow,
            AnchorColumn = _anchorColumn,
        };

    private async Task OnNextAsync()
    {
        if (_step == StepObject && string.IsNullOrWhiteSpace(_objectPicker?.SelectedApiName))
        {
            MessageBox.Show("Select a Salesforce object.", Title, MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        _next.IsEnabled = false;
        try
        {
            if (_step == StepObject)
            {
                var selectedObjectName = _objectPicker!.SelectedApiName!;
                _selectedObjectApiName = selectedObjectName;
                var objectChanged = !string.Equals(
                    _committedObjectApiName,
                    selectedObjectName,
                    StringComparison.OrdinalIgnoreCase);

                _describe = await _describeLoader(selectedObjectName).ConfigureAwait(true);
                if (objectChanged)
                {
                    _selectedFieldNames.Clear();
                    _criteriaDrafts.Clear();
                }

                _committedObjectApiName = selectedObjectName;
            }

            if (_step == StepFields && GetSelectedFields().Count == 0)
            {
                MessageBox.Show("Select at least one field.", Title, MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_step == StepCriteria)
            {
                CaptureCriteriaDrafts();
            }

            if (_step < LastStep)
            {
                ShowStep(_step + 1);
                return;
            }

            DialogResult = true;
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Could not continue the wizard:{Environment.NewLine}{Environment.NewLine}{ex.Message}",
                Title,
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            _next.IsEnabled = true;
        }
    }

    private void MoveStep(int delta)
    {
        if (_step + delta < StepObject)
        {
            return;
        }

        if (_step == StepCriteria)
        {
            CaptureCriteriaDrafts();
        }

        ShowStep(_step + delta);
    }

    private void ShowStep(int step)
    {
        _step = step;
        _content.Children.Clear();
        _content.RowDefinitions.Clear();
        _back.IsEnabled = step > StepObject;
        _next.IsDefault = true;
        _next.Content = step == LastStep ? "Run Query" : "Next";

        AddContentRow(GridLength.Auto);
        AddToGrid(CreateDestinationBanner(), 0);

        switch (step)
        {
            case StepObject:
                BuildObjectStep(contentStartRow: 1);
                break;
            case StepFields:
                BuildFieldStep(contentStartRow: 1);
                break;
            case StepCriteria:
                BuildCriteriaStep(contentStartRow: 1);
                break;
        }
    }

    private UIElement CreateDestinationBanner() =>
        new TextBlock
        {
            Text = $"Destination cell: {_destinationLabel}",
            FontWeight = FontWeights.SemiBold,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 12),
        };

    private void AddContentRow(GridLength height) =>
        _content.RowDefinitions.Add(new RowDefinition { Height = height });

    private void AddToGrid(UIElement element, int row)
    {
        Grid.SetRow(element, row);
        _content.Children.Add(element);
    }

    private void BuildObjectStep(int contentStartRow)
    {
        AddContentRow(GridLength.Auto);
        AddContentRow(new GridLength(1, GridUnitType.Star));

        AddToGrid(new TextBlock { Text = "Choose a Salesforce object:", Margin = new Thickness(0, 0, 0, 8) }, contentStartRow);

        var picker = new ObjectPickerControl(_objects, SelectionMode.Single);
        picker.SelectApiNames(string.IsNullOrWhiteSpace(_selectedObjectApiName)
            ? Array.Empty<string>()
            : new[] { _selectedObjectApiName! });
        _objectPicker = picker;
        AddToGrid(picker, contentStartRow + 1);
    }

    private void BuildFieldStep(int contentStartRow)
    {
        AddContentRow(GridLength.Auto);
        AddContentRow(new GridLength(1, GridUnitType.Star));

        AddToGrid(new TextBlock { Text = "Select fields for the table:", Margin = new Thickness(0, 0, 0, 8) }, contentStartRow);

        var fieldList = CreateFieldPickerGrid();
        UiListBox.MakeScrollable(fieldList);
        _fieldGrid = fieldList;
        _fieldRows.Clear();
        _idFieldItem = null;
        foreach (var (field, index) in WizardTableLayoutBuilder.OrderFieldsForWizard(_describe!).Select((field, index) => (field, index)))
        {
            var item = new FieldListItem(field, index);
            _fieldRows.Add(item);
            if (field.IsId)
            {
                _idFieldItem = item;
                _selectedFieldNames.Add(item.Field.Name);
            }
            else if (_selectedFieldNames.Contains(field.Name))
            {
                _selectedFieldNames.Add(item.Field.Name);
            }
        }

        fieldList.SelectionChanged += OnFieldListSelectionChanged;
        RefreshFieldGrid();
        AddToGrid(fieldList, contentStartRow + 1);
    }

    private ListView CreateFieldPickerGrid()
    {
        var list = new ListView
        {
            SelectionMode = SelectionMode.Extended,
            BorderBrush = Brushes.LightGray,
            BorderThickness = new Thickness(1),
            ItemContainerStyle = UiListBox.CreateStretchingPaddedItemStyle(),
        };
        list.View = new GridView
        {
            AllowsColumnReorder = false,
            Columns =
            {
                CreateFieldColumn("Category", nameof(FieldListItem.CategoryText), nameof(FieldListItem.CategoryBackground), nameof(FieldListItem.CategoryForeground), 150, isCategory: true),
                CreateFieldColumn("Label", nameof(FieldListItem.LabelText), null, null, 240, isCategory: false),
                CreateFieldColumn("API Name", nameof(FieldListItem.ApiName), null, null, 220, isCategory: false),
            },
        };
        list.AddHandler(GridViewColumnHeader.ClickEvent, new RoutedEventHandler(OnFieldColumnHeaderClick));
        return list;
    }

    private static GridViewColumn CreateFieldColumn(
        string header,
        string textProperty,
        string? backgroundProperty,
        string? foregroundProperty,
        double width,
        bool isCategory)
    {
        var column = new GridViewColumn { Header = header, Width = width };
        column.CellTemplate = CreateFieldTemplate(textProperty, backgroundProperty, foregroundProperty, isCategory);
        return column;
    }

    private static DataTemplate CreateFieldTemplate(
        string textProperty,
        string? backgroundProperty,
        string? foregroundProperty,
        bool isCategory)
    {
        var border = new FrameworkElementFactory(typeof(Border));
        border.SetValue(Border.PaddingProperty, new Thickness(0));
        border.SetValue(Border.HorizontalAlignmentProperty, HorizontalAlignment.Stretch);
        border.SetValue(Border.VerticalAlignmentProperty, VerticalAlignment.Stretch);
        border.SetValue(Border.SnapsToDevicePixelsProperty, true);

        if (backgroundProperty is not null)
        {
            border.SetBinding(Border.BackgroundProperty, new Binding(backgroundProperty));
        }

        var text = new FrameworkElementFactory(typeof(TextBlock));
        text.SetBinding(TextBlock.TextProperty, new Binding(textProperty));
        text.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
        text.SetValue(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Stretch);
        text.SetValue(TextBlock.TextAlignmentProperty, TextAlignment.Left);
        text.SetValue(TextBlock.PaddingProperty, new Thickness(8, 4, 8, 4));
        text.SetValue(TextBlock.TextWrappingProperty, TextWrapping.NoWrap);
        text.SetValue(TextBlock.ForegroundProperty, Brushes.Black);
        if (foregroundProperty is not null)
        {
            text.SetBinding(TextBlock.ForegroundProperty, new Binding(foregroundProperty));
        }

        if (isCategory)
        {
            text.SetValue(TextBlock.FontWeightProperty, FontWeights.SemiBold);
        }

        border.AppendChild(text);
        return new DataTemplate { VisualTree = border };
    }

    private void OnFieldListSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressSelectionEvents)
        {
            return;
        }

        SyncSelectedFieldNames();
    }

    private void OnFieldColumnHeaderClick(object sender, RoutedEventArgs e)
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
            case "Category":
                _fieldSortState.Promote(FieldPickerSortColumn.Category);
                break;
            case "Label":
                _fieldSortState.Promote(FieldPickerSortColumn.Label);
                break;
            case "API Name":
                _fieldSortState.Promote(FieldPickerSortColumn.ApiName);
                break;
            default:
                return;
        }

        RefreshFieldGrid();
    }

    private void SyncSelectedFieldNames()
    {
        var fieldList = _fieldGrid;
        if (fieldList is null)
        {
            return;
        }

        _selectedFieldNames.Clear();
        foreach (var item in fieldList.SelectedItems.Cast<FieldListItem>())
        {
            _selectedFieldNames.Add(item.Field.Name);
        }

        if (_idFieldItem is null)
        {
            return;
        }

        _selectedFieldNames.Add(_idFieldItem.Field.Name);
        if (!fieldList.SelectedItems.Contains(_idFieldItem))
        {
            fieldList.SelectedItems.Add(_idFieldItem);
        }
    }

    private void RefreshFieldGrid()
    {
        var fieldGrid = _fieldGrid;
        if (fieldGrid is null)
        {
            return;
        }

        var rows = _fieldSortState.Sort(_fieldRows);
        _suppressSelectionEvents = true;
        try
        {
            fieldGrid.ItemsSource = rows;
            RestoreFieldSelection();
        }
        finally
        {
            _suppressSelectionEvents = false;
        }
    }

    private void RestoreFieldSelection()
    {
        var fieldGrid = _fieldGrid;
        if (fieldGrid is null)
        {
            return;
        }

        if (_selectedFieldNames.Count == 0)
        {
            return;
        }

        fieldGrid.SelectedItems.Clear();
        foreach (var row in fieldGrid.Items.OfType<FieldListItem>())
        {
            if (_selectedFieldNames.Contains(row.Field.Name))
            {
                fieldGrid.SelectedItems.Add(row);
            }
        }
    }

    private void BuildCriteriaStep(int contentStartRow)
    {
        AddContentRow(GridLength.Auto);
        AddContentRow(new GridLength(1, GridUnitType.Star));
        AddContentRow(GridLength.Auto);

        AddToGrid(new TextBlock
        {
            Text = "Optional WHERE clauses (leave empty to query all records):",
            Margin = new Thickness(0, 0, 0, 8),
        }, contentStartRow);

        var criteriaPanel = new StackPanel();
        _criteriaPanel = criteriaPanel;
        _criteriaRows.Clear();
        var criteriaScroll = new ScrollViewer
        {
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            Content = criteriaPanel,
        };
        AddToGrid(criteriaScroll, contentStartRow + 1);
        if (_criteriaDrafts.Count == 0)
        {
            AddCriteriaRow();
        }
        else
        {
            foreach (var draft in _criteriaDrafts)
            {
                AddCriteriaRow(draft);
            }
        }

        var add = new Button { Content = "Add clause", Width = 100, HorizontalAlignment = HorizontalAlignment.Left, Margin = new Thickness(0, 8, 0, 0) };
        add.Click += (_, _) => AddCriteriaRow();
        AddToGrid(add, contentStartRow + 2);
    }

    private void AddCriteriaRow(WizardCriteriaDraft? draft = null)
    {
        var criteriaPanel = _criteriaPanel;
        if (criteriaPanel is null)
        {
            return;
        }

        var row = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 8) };
        var field = new ComboBox { Width = 180, Margin = new Thickness(0, 0, 8, 0) };
        foreach (var f in _describe!.Fields.OrderBy(f => f.Label))
        {
            field.Items.Add(f);
        }

        field.DisplayMemberPath = "Label";
        var op = new ComboBox { Width = 120, Margin = new Thickness(0, 0, 8, 0) };
        foreach (var o in Operators)
        {
            op.Items.Add(o);
        }

        op.SelectedIndex = 0;
        var value = new TextBox { Width = 160 };
        if (draft is not null)
        {
            field.SelectedItem = _describe!.Fields.FirstOrDefault(f =>
                string.Equals(f.Name, draft.FieldApiName, StringComparison.OrdinalIgnoreCase));
            op.SelectedItem = draft.Operator;
            value.Text = draft.Value ?? string.Empty;
        }

        row.Children.Add(field);
        row.Children.Add(op);
        row.Children.Add(value);
        criteriaPanel.Children.Add(row);
        _criteriaRows.Add((field, op, value));
    }

    private void CaptureCriteriaDrafts()
    {
        _criteriaDrafts.Clear();
        foreach (var (field, op, value) in _criteriaRows)
        {
            if (field.SelectedItem is not FieldDescriptor fieldDescriptor)
            {
                continue;
            }

            var operatorText = op.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(operatorText))
            {
                continue;
            }

            _criteriaDrafts.Add(new WizardCriteriaDraft(fieldDescriptor.Name, operatorText, value.Text));
        }
    }

    private IReadOnlyList<FieldDescriptor> GetSelectedFields()
    {
        var selected = _selectedFieldNames.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var id = _describe!.Fields.First(f => f.IsId);
        selected.Add(id.Name);

        return WizardTableLayoutBuilder.OrderFieldsForWizard(_describe)
            .Where(f => selected.Contains(f.Name))
            .ToList();
    }

    private IReadOnlyList<WizardCriteriaClause> BuildCriteria()
    {
        var clauses = new List<WizardCriteriaClause>();
        foreach (var (field, op, value) in _criteriaRows)
        {
            if (field.SelectedItem is not FieldDescriptor fieldDescriptor)
            {
                continue;
            }

            var operatorText = op.SelectedItem?.ToString();
            if (operatorText is null || string.IsNullOrWhiteSpace(operatorText))
            {
                continue;
            }

            clauses.Add(new WizardCriteriaClause
            {
                Field = fieldDescriptor,
                Operator = operatorText,
                Value = value.Text,
            });
        }

        if (clauses.Count == 0)
        {
            var idField = _describe!.Fields.First(f => f.IsId);
            clauses.Add(new WizardCriteriaClause
            {
                Field = idField,
                Operator = "not equals",
                Value = string.Empty,
            });
        }

        return clauses;
    }

    /// <summary>Formats 1-based row/column as an A1-style address (e.g. R3 C5 → E3).</summary>
    internal static string FormatA1(int row, int column)
    {
        if (row < 1 || column < 1)
        {
            return $"R{row}C{column}";
        }

        var letters = new StringBuilder();
        var n = column;
        while (n > 0)
        {
            n--;
            letters.Insert(0, (char)('A' + (n % 26)));
            n /= 26;
        }

        return letters.ToString() + row;
    }
}
