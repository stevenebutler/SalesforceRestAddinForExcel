# About

**Ribbon:** `btnAbout` — large button, Add-Ins tab  
**Legacy:** `ForceRibbon.btnAbout_Click` → `ForceConnector.OpenAbout()`  
**Login required:** No  
**COM / VBA:** None

See also: [common-performance-requirements.md](./common-performance-requirements.md) (minimal Excel surface).

---

## Summary

Shows a modal **About** dialog with product name, version, attribution, and an external link. No Salesforce or worksheet interaction.

---

## Functional requirements

### FR-AB-1 Trigger

- User clicks **About** on the ForceConnector ribbon group.
- Handler runs immediately; no session check.

### FR-AB-2 Dialog content

Display at minimum:

| Element | Legacy source | Status |
|---------|---------------|--------|
| Product title | “Salesforce REST Add-in for Excel” | Implemented |
| Version | Build/add-in version string | Implemented |
| Author / copyright | Attribution labels | **Pending** |
| Link | GitHub or project URL (optional, user-opened in browser) | **Pending** |

### FR-AB-3 Dismissal

- User closes dialog with **Close** or standard window chrome.
- No side effects on session, workbook, or options.

---

## Non-functional requirements

### NFR-AB-1 Performance

- No Excel range access; no COM hot path.
- Dialog is modal and blocks only the add-in UI thread (acceptable).

### NFR-AB-2 Platform

- WPF window in `SalesforceRestAddin.Windows.Ui` — see [ui-platform.md](./ui-platform.md).
- Not required on Linux CI (no Excel/WPF).

### NFR-AB-3 Accessibility

- Standard WPF dialog patterns; keyboard dismiss.

---

## Legacy reference

- `ForceConnector/frmAbout.cs`
- `ThisAddIn.Ver` for version string

## Out of scope

- Update checks, license validation, telemetry.
