using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
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
    private int _step;
    private SObjectSummary? _selectedObject;
    private SObjectDescribe? _describe;
    private FieldListItem? _idFieldItem;
    private readonly ListBox _fieldList = CreateFieldPickerList();
    private readonly StackPanel _criteriaPanel = new();
    private readonly List<(ComboBox Field, ComboBox Operator, TextBox Value)> _criteriaRows = new();

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
        Width = 560;
        Height = 460;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;

        _content = new Grid { Margin = new Thickness(16) };
        _back = new Button { Content = "Back", Width = 80, Margin = new Thickness(0, 0, 8, 0) };
        _next = new Button { Content = "Next", Width = 100 };
        var cancel = new Button { Content = "Cancel", Width = 80, Margin = new Thickness(8, 0, 0, 0) };
        _back.Click += (_, _) => MoveStep(-1);
        _next.Click += async (_, _) => await OnNextAsync();
        cancel.Click += (_, _) => { DialogResult = false; Close(); };
        // Record Id is required and non-deselectable (FR-TQW-4).
        _fieldList.SelectionChanged += OnFieldListSelectionChanged;

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(16),
        };
        buttons.Children.Add(_back);
        buttons.Children.Add(_next);
        buttons.Children.Add(cancel);

        var root = new DockPanel();
        DockPanel.SetDock(buttons, Dock.Bottom);
        root.Children.Add(buttons);
        root.Children.Add(_content);
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
        if (_step == StepObject && _selectedObject is null)
        {
            MessageBox.Show("Select a Salesforce object.", Title, MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        _next.IsEnabled = false;
        try
        {
            if (_step == StepObject)
            {
                _describe = await _describeLoader(_selectedObject!.Name).ConfigureAwait(true);
            }

            if (_step == StepFields && GetSelectedFields().Count == 0)
            {
                MessageBox.Show("Select at least one field.", Title, MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
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

        ShowStep(_step + delta);
    }

    private void ShowStep(int step)
    {
        _step = step;
        _content.Children.Clear();
        _content.RowDefinitions.Clear();
        _back.IsEnabled = step > StepObject;
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
        AddContentRow(GridLength.Auto);
        AddContentRow(new GridLength(1, GridUnitType.Star));

        AddToGrid(new TextBlock { Text = "Choose a Salesforce object:", Margin = new Thickness(0, 0, 0, 8) }, contentStartRow);

        var filter = new TextBox { Margin = new Thickness(0, 0, 0, 8) };
        AddToGrid(filter, contentStartRow + 1);

        var list = UiListBox.Create();
        UiListBox.MakeScrollable(list);
        var items = _objects.Where(o => o.Queryable).OrderBy(o => o.Label, StringComparer.OrdinalIgnoreCase).ToList();
        foreach (var item in items)
        {
            list.Items.Add(new SObjectListItem(item));
        }

        list.SelectionChanged += (_, _) =>
        {
            _selectedObject = (list.SelectedItem as SObjectListItem)?.Summary;
        };

        filter.TextChanged += (_, _) =>
        {
            var text = filter.Text.Trim();
            list.Items.Clear();
            foreach (var item in items.Where(i =>
                         string.IsNullOrEmpty(text)
                         || i.Label.Contains(text, StringComparison.OrdinalIgnoreCase)
                         || i.Name.Contains(text, StringComparison.OrdinalIgnoreCase)))
            {
                list.Items.Add(new SObjectListItem(item));
            }
        };

        AddToGrid(list, contentStartRow + 2);
    }

    private void BuildFieldStep(int contentStartRow)
    {
        AddContentRow(GridLength.Auto);
        AddContentRow(new GridLength(1, GridUnitType.Star));

        AddToGrid(new TextBlock { Text = "Select fields for the table:", Margin = new Thickness(0, 0, 0, 8) }, contentStartRow);

        _fieldList.Items.Clear();
        UiListBox.MakeScrollable(_fieldList);
        _idFieldItem = null;
        foreach (var field in WizardTableLayoutBuilder.OrderFieldsForWizard(_describe!))
        {
            var item = new FieldListItem(field);
            _fieldList.Items.Add(item);
            if (field.IsId)
            {
                _idFieldItem = item;
                _fieldList.SelectedItems.Add(item);
            }
        }

        AddToGrid(_fieldList, contentStartRow + 1);
    }

    private void OnFieldListSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_idFieldItem is null)
        {
            return;
        }

        if (!_fieldList.SelectedItems.Contains(_idFieldItem))
        {
            _fieldList.SelectedItems.Add(_idFieldItem);
        }
    }

    private static ListBox CreateFieldPickerList()
    {
        var list = UiListBox.Create(SelectionMode.Extended);
        list.ItemContainerStyle = UiListBox.CreateFieldPickerItemStyle();
        return list;
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

        _criteriaPanel.Children.Clear();
        _criteriaRows.Clear();
        var criteriaScroll = new ScrollViewer
        {
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            Content = _criteriaPanel,
        };
        AddToGrid(criteriaScroll, contentStartRow + 1);
        AddCriteriaRow();

        var add = new Button { Content = "Add clause", Width = 100, HorizontalAlignment = HorizontalAlignment.Left, Margin = new Thickness(0, 8, 0, 0) };
        add.Click += (_, _) => AddCriteriaRow();
        AddToGrid(add, contentStartRow + 2);
    }

    private void AddCriteriaRow()
    {
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
        row.Children.Add(field);
        row.Children.Add(op);
        row.Children.Add(value);
        _criteriaPanel.Children.Add(row);
        _criteriaRows.Add((field, op, value));
    }

    private IReadOnlyList<FieldDescriptor> GetSelectedFields()
    {
        var selected = _fieldList.SelectedItems.Cast<FieldListItem>().Select(i => i.Field).ToList();
        var id = _describe!.Fields.First(f => f.IsId);
        if (!selected.Any(f => f.IsId))
        {
            selected.Insert(0, id);
        }

        return WizardTableLayoutBuilder.OrderFieldsForWizard(_describe)
            .Where(f => selected.Any(s => string.Equals(s.Name, f.Name, StringComparison.OrdinalIgnoreCase)))
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
