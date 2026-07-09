# About

**Ribbon:** `btnAbout` — large button, Add-Ins tab  
**Login required:** No  
**COM / VBA:** None

See also: [common-performance-requirements.md](./common-performance-requirements.md) (minimal Excel surface).

---

## Summary

Shows a modal **About** dialog with product name, version, attribution, and an external link. No Salesforce or worksheet interaction.

---

## Functional requirements

### FR-AB-1 Trigger

- User clicks **About** on the ribbon group (custom large icon: teal cloud + worksheet motif — original artwork, not Salesforce branding).
- Handler runs immediately; no session check.

### FR-AB-2 Dialog content

Display at minimum:

| Element | Content | Status |
|---------|---------|--------|
| Product title | `ProductBranding.ProductName` | Implemented |
| Tagline | GitHub / product description | Implemented |
| Version | Build/add-in commit id | Implemented |
| Author / copyright | Steven Butler / © 2026 | Implemented |
| Link | https://github.com/stevenebutler/SalesforceRestAddinForExcel (opens in browser) | Implemented |
| Credits | Inspired by Force.com Connector Next Generation (https://github.com/good-ghost/ForceConnector) | Implemented |

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

## Out of scope

- Update checks, license validation, telemetry.
