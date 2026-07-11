using SalesforceRestAddin.Core.Rest;

namespace SalesforceRestAddin.Core.DataPlane;

public enum WizardObjectKind
{
    Standard = 0,
    Custom = 1,
    System = 2,
    HiddenSystem = 3,
}

public static class WizardObjectListFilter
{
    private static readonly HashSet<string> LegacyStandardObjects = new(StringComparer.OrdinalIgnoreCase)
    {
        "Account",
        "Campaign",
        "Case",
        "Contact",
        "Contract",
        "Event",
        "Lead",
        "Opportunity",
        "Pricebook2",
        "Product2",
        "Profile",
        "Quote",
        "Task",
        "User",
        "UserRole",
    };

    public static bool IsVisibleInPicker(SObjectSummary summary) =>
        GetKind(summary) != WizardObjectKind.HiddenSystem;

    public static WizardObjectKind GetKind(SObjectSummary summary)
    {
        if (summary.DeprecatedAndHidden)
        {
            return WizardObjectKind.HiddenSystem;
        }

        if (summary.Custom)
        {
            return WizardObjectKind.Custom;
        }

        return LegacyStandardObjects.Contains(summary.Name)
            ? WizardObjectKind.Standard
            : WizardObjectKind.System;
    }

    public static string? GetNamespace(SObjectSummary summary)
    {
        if (summary.Name.StartsWith("SFL5_", StringComparison.OrdinalIgnoreCase))
        {
            return "SFL5";
        }

        if (summary.Name.StartsWith("SFDC_", StringComparison.OrdinalIgnoreCase))
        {
            return "SFDC";
        }

        var firstSeparator = summary.Name.IndexOf("__", StringComparison.Ordinal);
        if (firstSeparator <= 0)
        {
            return null;
        }

        var secondSeparator = summary.Name.IndexOf("__", firstSeparator + 2, StringComparison.Ordinal);
        if (secondSeparator <= firstSeparator + 2)
        {
            return null;
        }

        return summary.Name[..firstSeparator];
    }

    public static string GetGroup(SObjectSummary summary)
    {
        var namespaceName = GetNamespace(summary);
        if (!string.IsNullOrEmpty(namespaceName))
        {
            return namespaceName!;
        }

        return summary.Custom ? " " : string.Empty;
    }

    public static ObjectPickerRow CreateRow(SObjectSummary summary)
    {
        return new ObjectPickerRow(
            summary,
            GetGroup(summary),
            summary.Label,
            summary.Name,
            GetKind(summary));
    }

    public static bool MatchesFilters(
        SObjectSummary summary,
        bool showStandardObjects,
        bool showCustomObjects,
        bool showSystemObjects)
    {
        if (!IsVisibleInPicker(summary))
        {
            return false;
        }

        return GetKind(summary) switch
        {
            WizardObjectKind.Standard => showStandardObjects,
            WizardObjectKind.Custom => showCustomObjects,
            WizardObjectKind.System => showSystemObjects,
            WizardObjectKind.HiddenSystem => false,
            _ => false,
        };
    }

    public static IReadOnlyList<ObjectPickerRow> BuildVisibleRows(
        IEnumerable<SObjectSummary> objects,
        bool showStandardObjects,
        bool showCustomObjects,
        bool showSystemObjects)
    {
        return objects
            .Where(o => o.Queryable)
            .Where(o => MatchesFilters(o, showStandardObjects, showCustomObjects, showSystemObjects))
            .Select(CreateRow)
            .ToList();
    }

    public static IReadOnlyList<ObjectPickerRow> FilterAndSort(
        IEnumerable<SObjectSummary> objects,
        bool showStandardObjects,
        bool showCustomObjects,
        bool showSystemObjects)
    {
        return FilterAndSort(
            objects,
            showStandardObjects,
            showCustomObjects,
            showSystemObjects,
            new ObjectPickerSortState());
    }

    public static IReadOnlyList<ObjectPickerRow> FilterAndSort(
        IEnumerable<SObjectSummary> objects,
        bool showStandardObjects,
        bool showCustomObjects,
        bool showSystemObjects,
        ObjectPickerSortState sortState)
    {
        return sortState.Sort(BuildVisibleRows(objects, showStandardObjects, showCustomObjects, showSystemObjects));
    }
}
