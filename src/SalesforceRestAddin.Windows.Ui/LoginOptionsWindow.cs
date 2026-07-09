using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Rest;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Windows.Ui;

public sealed class LoginDialogResult
{
    public required SalesforceLoginOptions LoginOptions { get; init; }

    /// <summary>
    /// Preferred REST API version, or <c>null</c> for latest supported by the org.
    /// </summary>
    public string? PreferredApiVersion { get; init; }

    /// <summary>
    /// When true, skip the saved refresh token and open Salesforce sign-in.
    /// </summary>
    public bool ForceSignInAgain { get; init; }
}

/// <summary>
/// Login target picker: tenant, production/sandbox, sandbox id, optional forced sign-in.
/// </summary>
public sealed class LoginOptionsWindow : Window
{
    private readonly TextBox _tenantBox;
    private readonly TextBox _sandboxIdBox;
    private readonly RadioButton _productionRadio;
    private readonly RadioButton _sandboxRadio;
    private readonly CheckBox _signInAgainCheck;
    private readonly TextBlock _savedSessionHint;
    private readonly StackPanel _sandboxPanel;
    private readonly ComboBox _apiVersionCombo;
    private readonly StackPanel _signInAgainPanel;
    private readonly ISessionCredentialStore? _credentialStore;

    public LoginOptionsWindow(UserLoginPreferences? preferences = null, ISessionCredentialStore? credentialStore = null)
    {
        _credentialStore = credentialStore;
        preferences ??= new UserLoginPreferences();

        Title = "Salesforce REST Add-in Sign In";
        Width = 440;
        SizeToContent = SizeToContent.Height;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        ResizeMode = ResizeMode.NoResize;

        var root = new StackPanel { Margin = new Thickness(16) };

        var organizationGroup = new GroupBox
        {
            Header = "Organization",
            Padding = new Thickness(12, 8, 12, 12),
            Margin = new Thickness(0, 0, 0, 12),
        };
        var organizationPanel = new StackPanel();

        organizationPanel.Children.Add(MakeFieldLabel("Tenant or URL"));
        _tenantBox = new TextBox { Text = preferences.Tenant ?? string.Empty };
        _tenantBox.TextChanged += (_, _) => UpdateSessionState();
        organizationPanel.Children.Add(_tenantBox);

        organizationPanel.Children.Add(MakeFieldLabel("Environment", topMargin: 12));
        var envRow = new StackPanel { Orientation = Orientation.Horizontal };
        _productionRadio = new RadioButton
        {
            Content = "Production",
            GroupName = "Environment",
            IsChecked = preferences.Environment == SalesforceEnvironment.Production,
            Margin = new Thickness(0, 0, 16, 0),
        };
        _sandboxRadio = new RadioButton
        {
            Content = "Sandbox",
            GroupName = "Environment",
            IsChecked = preferences.Environment == SalesforceEnvironment.Sandbox,
        };
        _productionRadio.Checked += (_, _) => UpdateSessionState();
        _sandboxRadio.Checked += (_, _) => UpdateSessionState();
        envRow.Children.Add(_productionRadio);
        envRow.Children.Add(_sandboxRadio);
        organizationPanel.Children.Add(envRow);

        _sandboxPanel = new StackPanel { Margin = new Thickness(0, 12, 0, 0) };
        _sandboxPanel.Children.Add(MakeFieldLabel("Sandbox ID"));
        _sandboxIdBox = new TextBox { Text = preferences.SandboxId ?? string.Empty };
        _sandboxIdBox.TextChanged += (_, _) => UpdateSessionState();
        _sandboxPanel.Children.Add(_sandboxIdBox);
        organizationPanel.Children.Add(_sandboxPanel);

        organizationPanel.Children.Add(MakeFieldLabel("API version", topMargin: 12));
        _apiVersionCombo = new ComboBox { IsEditable = false };
        PopulateApiVersionChoices(preferences);
        organizationPanel.Children.Add(_apiVersionCombo);

        organizationGroup.Content = organizationPanel;
        root.Children.Add(organizationGroup);

        var signInGroup = new GroupBox
        {
            Header = "Sign-in",
            Padding = new Thickness(12, 8, 12, 12),
            Margin = new Thickness(0, 0, 0, 12),
        };
        var signInPanel = new StackPanel();

        _savedSessionHint = new TextBlock
        {
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 4),
        };
        signInPanel.Children.Add(_savedSessionHint);

        _signInAgainPanel = new StackPanel { Margin = new Thickness(0, 8, 0, 0) };
        _signInAgainCheck = new CheckBox();
        _signInAgainCheck.Content = new TextBlock
        {
            Text = "Sign in again in Salesforce (ignore saved sign-in)",
            TextWrapping = TextWrapping.Wrap,
        };
        _signInAgainCheck.Checked += (_, _) => UpdateSessionState();
        _signInAgainCheck.Unchecked += (_, _) => UpdateSessionState();
        _signInAgainPanel.Children.Add(_signInAgainCheck);
        signInPanel.Children.Add(_signInAgainPanel);

        signInGroup.Content = signInPanel;
        root.Children.Add(signInGroup);

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
        };
        var cancel = new Button { Content = "Cancel", Width = 88, Margin = new Thickness(0, 0, 8, 0), IsCancel = true };
        cancel.Click += (_, _) =>
        {
            DialogResult = false;
            Close();
        };
        var signIn = new Button { Content = "Sign in", Width = 88, IsDefault = true };
        signIn.Click += (_, _) =>
        {
            DialogResult = true;
            Close();
        };
        buttons.Children.Add(cancel);
        buttons.Children.Add(signIn);
        root.Children.Add(buttons);

        Content = root;
        UpdateSessionState();
    }

    public LoginDialogResult? GetResult()
    {
        if (DialogResult != true)
        {
            return null;
        }

        var isSandbox = _sandboxRadio.IsChecked == true;
        var forceSignInAgain = _signInAgainPanel.Visibility == Visibility.Visible
                               && _signInAgainCheck.IsChecked == true;

        return new LoginDialogResult
        {
            LoginOptions = SalesforceLoginOptionsNormalizer.Normalize(new SalesforceLoginOptions
            {
                Tenant = string.IsNullOrWhiteSpace(_tenantBox.Text) ? null : _tenantBox.Text.Trim(),
                Environment = isSandbox ? SalesforceEnvironment.Sandbox : SalesforceEnvironment.Production,
                SandboxId = isSandbox && !string.IsNullOrWhiteSpace(_sandboxIdBox.Text)
                    ? _sandboxIdBox.Text.Trim()
                    : null,
            }),
            PreferredApiVersion = GetSelectedApiVersion(),
            ForceSignInAgain = forceSignInAgain,
        };
    }

    private void UpdateSessionState()
    {
        var isSandbox = _sandboxRadio.IsChecked == true;
        _sandboxPanel.Visibility = isSandbox ? Visibility.Visible : Visibility.Collapsed;

        if (_credentialStore is null)
        {
            _savedSessionHint.Text = string.Empty;
            _signInAgainPanel.Visibility = Visibility.Collapsed;
            _signInAgainCheck.IsChecked = false;
            return;
        }

        try
        {
            var loginOptions = BuildCurrentLoginOptions();
            var hostKey = SessionHostKey.FromLoginOptions(loginOptions);
            var hasSaved = _credentialStore.Load(hostKey) is not null;

            if (_signInAgainCheck.IsChecked == true)
            {
                _savedSessionHint.Text = string.Empty;
                _signInAgainPanel.Visibility = Visibility.Visible;
                return;
            }

            _signInAgainPanel.Visibility = hasSaved ? Visibility.Visible : Visibility.Collapsed;
            _signInAgainCheck.IsChecked = false;

            _savedSessionHint.Text = hasSaved
                ? $"A saved sign-in is available for {hostKey}."
                : $"You will sign in to {hostKey} in Salesforce.";
        }
        catch (Exception ex)
        {
            _savedSessionHint.Text = $"Could not check saved sign-in: {ex.Message}";
            _signInAgainPanel.Visibility = Visibility.Collapsed;
            _signInAgainCheck.IsChecked = false;
        }
    }

    private SalesforceLoginOptions BuildCurrentLoginOptions()
    {
        var isSandbox = _sandboxRadio.IsChecked == true;
        return SalesforceLoginOptionsNormalizer.Normalize(new SalesforceLoginOptions
        {
            Tenant = string.IsNullOrWhiteSpace(_tenantBox.Text) ? null : _tenantBox.Text.Trim(),
            Environment = isSandbox ? SalesforceEnvironment.Sandbox : SalesforceEnvironment.Production,
            SandboxId = isSandbox && !string.IsNullOrWhiteSpace(_sandboxIdBox.Text)
                ? _sandboxIdBox.Text.Trim()
                : null,
        });
    }

    private void PopulateApiVersionChoices(UserLoginPreferences preferences)
    {
        _apiVersionCombo.Items.Clear();
        _apiVersionCombo.Items.Add(new ApiVersionComboItem("Latest (from org)", null));

        var versions = new SortedSet<string>(Comparer<string>.Create(SalesforceApiVersions.CompareDescending));
        foreach (var version in preferences.CachedSupportedApiVersions)
        {
            versions.Add(version);
        }

        var preferredApiVersion = preferences.ApiVersion;
        if (preferredApiVersion is not null && !string.IsNullOrWhiteSpace(preferredApiVersion))
        {
            versions.Add(preferredApiVersion);
        }

        foreach (var version in versions)
        {
            _apiVersionCombo.Items.Add(new ApiVersionComboItem(version, version));
        }

        var selected = preferredApiVersion is null || string.IsNullOrWhiteSpace(preferredApiVersion)
            ? 0
            : _apiVersionCombo.Items
                .Cast<ApiVersionComboItem>()
                .ToList()
                .FindIndex(item => item.Version is not null
                                   && SalesforceApiVersions.AreEqual(item.Version, preferredApiVersion));

        _apiVersionCombo.SelectedIndex = selected >= 0 ? selected : 0;
    }

    private string? GetSelectedApiVersion() =>
        _apiVersionCombo.SelectedItem is ApiVersionComboItem item ? item.Version : null;

    private sealed class ApiVersionComboItem
    {
        public ApiVersionComboItem(string label, string? version)
        {
            Label = label;
            Version = version;
        }

        public string Label { get; }

        public string? Version { get; }

        public override string ToString() => Label;
    }

    private static TextBlock MakeFieldLabel(string text, double topMargin = 0) =>
        new()
        {
            Text = text,
            Margin = new Thickness(0, topMargin, 0, 4),
            FontWeight = FontWeights.SemiBold,
        };
}
