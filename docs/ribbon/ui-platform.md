# UI platform

User interface for the Excel-DNA add-in. **No WinForms** in `src/` — all dialogs are WPF (code-only) or WebView2 for OAuth.

## WPF: code-only (no XAML)

All WPF UI is built in **C# only** — construct `Window`, `Grid`, `Button`, etc. in code and assign `Content` / `Children` in constructors or helper methods.

**No XAML** (no `.xaml` files, no `InitializeComponent()` from generated partials, no resource dictionaries). This is required by the build setup:

- `SalesforceRestAddin.Windows.Ui` sets `<UseWPF>false</UseWPF>` and references `PresentationCore`, `PresentationFramework`, `WindowsBase`, and `System.Xaml` directly.
- **`SalesforceRestAddin.ExcelDna`** must reference the same WPF assemblies when it uses `Windows.Ui` windows or `System.Windows.MessageBox` — see [AGENTS.md](../../AGENTS.md) (WPF compile on Linux).
- The Linux/container build path does not run the WPF markup compiler; XAML would not pack reliably into the Excel-DNA XLL workflow.
- Existing pattern: `LoginOptionsWindow`, `OAuthWebViewWindow` — subclass `Window`, build the tree in the constructor.

When adding a new dialog, copy that style. Do not add `<UseWPF>true</UseWPF>` or `.xaml` files without an explicit build/packaging change.

## Stack

| Layer | Technology | Use for |
|-------|------------|---------|
| `SalesforceRestAddin.Windows.Ui` | **WPF in C#** (`System.Windows`) | All dialogs and windows: About, Options, login options, confirmations, progress, Table Query Wizard, Describe object picker |
| `SalesforceRestAddin.Windows.Ui` | **WebView2** (`OAuthWebViewWindow`) | OAuth authorization redirect only — embedded browser for Salesforce sign-in |
| `SalesforceRestAddin.ExcelDna` | Excel-DNA + COM | Ribbon, `QueueAsMacro`, bulk range read/write — **not** general UI |
| `SalesforceRestAddin.Core` | None | No UI references |

## Rules

1. **No WinForms** in `src/` — do not add `System.Windows.Forms` references or forms.
2. **No XAML** — code-built WPF only (see above).
3. **OAuth** uses WebView2, not an external browser and not plain WPF controls in place of a browser (except the WebView2 host chrome).
4. **Login options** (`LoginOptionsWindow`) and **OAuth** (`OAuthWebViewWindow`) are the reference implementations — extend that project for new UI.
5. **Modal dialogs** — WPF `Window` with `ShowDialog()` from the Excel UI thread, same pattern as existing login flow.
6. **Long operations** — sole wait/pump is `ExcelStaAsyncHost` (modal WPF `ShowDialog` + StatusBar); async HTTP in Core; Excel sheet updates via `QueueAsMacro` / `ExcelComThread`. See NFR-STA-1 in [common-performance-requirements.md](./common-performance-requirements.md).
7. **Excel-DNA** hosts WPF via `SalesforceRestAddin.Windows.Ui`; use code-built WPF windows from that project (not `Interaction.MsgBox`).

## What goes where

| Feature | UI home |
|---------|---------|
| About | WPF `AboutWindow` |
| Options | WPF `OptionsWindow` |
| Login options | WPF `LoginOptionsWindow` |
| OAuth sign-in | WPF `OAuthWebViewWindow` + WebView2 |
| Table Query Wizard | WPF `TableQueryWizardWindow` |
| Describe object picker | WPF `DescribeObjectPickerWindow` |
| Insert/delete confirm | WPF message dialog or lightweight window |
| Query/update/delete progress | WPF `DataOperationProgressWindow` |
| Sheet read/write | ExcelDna only |

## Testing

- Core and data-plane logic: no UI; test with fakes (`FakeLoginOptionsPresenter`, `FakeInteractiveLoginHandler`).
- WPF: manual Windows smoke; optional future UI tests out of scope for Linux CI.
- WebView2: manual OAuth smoke on Windows VM.
