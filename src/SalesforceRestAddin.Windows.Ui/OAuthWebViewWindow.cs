using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using SalesforceRestAddin.Core.OAuth;
using SalesforceRestAddin.Core.Session;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;

namespace SalesforceRestAddin.Windows.Ui;

/// <summary>
/// Embedded WebView2 (Edge Chromium) OAuth authorize flow. Intercepts the loopback redirect URI.
/// </summary>
public sealed class OAuthWebViewWindow : Window
{
    private readonly WebView2 _webView;
    private readonly TextBox _addressBox;
    private readonly SalesforceOAuthOptions _oauth;
    private bool _completed;

    public OAuthWebViewWindow(string authorizeUrl, SalesforceOAuthOptions? oauth = null)
    {
        _oauth = oauth ?? SalesforceOAuthOptions.Default;
        AuthorizeUrl = authorizeUrl;

        Title = "Salesforce Sign In";
        Width = 960;
        Height = 720;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;

        var root = new Grid { Margin = new Thickness(8) };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        _addressBox = new TextBox
        {
            IsReadOnly = true,
            IsTabStop = false,
            Margin = new Thickness(4, 0, 4, 8),
            Text = authorizeUrl,
            FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
        };
        Grid.SetRow(_addressBox, 0);
        root.Children.Add(_addressBox);

        _webView = new WebView2 { Margin = new Thickness(4, 0, 4, 0) };
        Grid.SetRow(_webView, 1);
        root.Children.Add(_webView);

        Content = root;
        Loaded += OnLoadedAsync;
    }

    public string AuthorizeUrl { get; }

    public string? AuthorizationCode { get; private set; }

    public string? ErrorMessage { get; private set; }

    private async void OnLoadedAsync(object sender, RoutedEventArgs e)
    {
        Loaded -= OnLoadedAsync;
        var stopwatch = Stopwatch.StartNew();

        try
        {
            SessionFlowTrace.Log("WebView2: configuring loader.");
            WebView2LoaderBootstrap.EnsureConfigured();

            SalesforceRestAddinDataPaths.EnsureDataDirectory();
            var webViewFolder = SalesforceRestAddinDataPaths.WebView2UserDataFolder;
            Directory.CreateDirectory(webViewFolder);

            SessionFlowTrace.Log($"WebView2: creating environment (userData={webViewFolder}).");
            var environment = await CoreWebView2Environment
                .CreateAsync(browserExecutableFolder: null, userDataFolder: webViewFolder)
                .ConfigureAwait(true);
            SessionFlowTrace.Log($"WebView2: environment ready ({stopwatch.ElapsedMilliseconds} ms).");

            await _webView.EnsureCoreWebView2Async(environment).ConfigureAwait(true);
            SessionFlowTrace.Log($"WebView2: control ready ({stopwatch.ElapsedMilliseconds} ms).");
            _webView.CoreWebView2.Settings.AreDevToolsEnabled = false;
            _webView.CoreWebView2.NavigationStarting += OnNavigationStarting;
            _webView.CoreWebView2.SourceChanged += OnSourceChanged;
            _webView.CoreWebView2.Navigate(AuthorizeUrl);
            UpdateAddress(AuthorizeUrl);
            SessionFlowTrace.Log($"WebView2: navigating to authorize URL ({stopwatch.ElapsedMilliseconds} ms).");
        }
        catch (Exception ex)
        {
            SessionFlowTrace.LogException($"WebView2: failed after {stopwatch.ElapsedMilliseconds} ms", ex);
            ErrorMessage = $"WebView2 failed to start: {ex.Message}";
            DialogResult = false;
            Close();
        }
    }

    private void OnSourceChanged(object? sender, CoreWebView2SourceChangedEventArgs e) =>
        UpdateAddress(_webView.Source?.ToString());

    private void OnNavigationStarting(object? sender, CoreWebView2NavigationStartingEventArgs e)
    {
        UpdateAddress(e.Uri);

        if (_completed)
        {
            return;
        }

        if (!OAuthRedirectParser.TryGetAuthorizationCode(e.Uri, _oauth.RedirectUri, out var code, out var error))
        {
            return;
        }

        e.Cancel = true;
        _completed = true;

        if (!string.IsNullOrWhiteSpace(error))
        {
            ErrorMessage = error;
            DialogResult = false;
            Close();
            return;
        }

        AuthorizationCode = code;
        SessionFlowTrace.Log("WebView2: authorization code captured from redirect.");
        DialogResult = true;
        Close();
    }

    private void UpdateAddress(string? url)
    {
        if (string.IsNullOrWhiteSpace(url))
        {
            return;
        }

        _addressBox.Text = url;
    }
}
