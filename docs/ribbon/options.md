# Options

**Ribbon:** `btnOptions` — “Options”  
**Login required:** No — ribbon button always available; clear-cache inside the dialog requires a live instance URL  
**COM / VBA:** None

See also: [common-performance-requirements.md](./common-performance-requirements.md)

---

## Summary

Modal dialog to edit connector behavior flags persisted to disk. Changes affect subsequent data-plane operations; no immediate Salesforce calls.

---

## Functional requirements

### FR-OPT-1 Availability

- Ribbon button **always enabled** (connector options are global, not per org).
- Click opens modal options dialog whether or not the user is signed in.
- Clear metadata cache remains session-dependent — see FR-OPT-6.

### FR-OPT-2 Settings (in scope)

Map to [`ConnectorOptions`](../../src/SalesforceRestAddin.Core/Session/ConnectorOptions.cs) / `JsonConnectorOptionsStore`:

| UI label | Property | Default | Affects |
|----------|----------|---------|---------|
| Use Reference Name/Id | `UseReference` | false | Reference field read/write in query/update/insert |
| Do not show warning dialogs before commencing operations. | `NoWarning` | false | Skip insert confirm; skip include-hidden update confirm |
| Confirm Large Query Download | `ConfirmLargeQuery` | false | Pre-query COUNT + confirm for Query Table Data |
| No Query Limit | `NoQueryLimit` | false | Skip 3,500 row / 20 col caps on query-rows, delete, update limits |
| Enable Auto Assign Rule | `AutoAssignRule` | false | When false (default), send `Sforce-Auto-Assign: FALSE` on create/update |
| Include Hidden Columns/Rows | `IncludeHiddenCells` | false | When unchecked (default), update omits AutoFilter/manually hidden rows/columns — see [update-selected-cells.md](./update-selected-cells.md) |

### FR-OPT-3 Settings (dropped)

| Topic | Status |
|-------|--------|
| Managed / translation data toggle | **Dropped** — Translation Helper out of scope; omit from dialog |

### FR-OPT-4 Persistence

- Load current values on open.
- **OK** saves to JSON store (`SalesforceRestAddinDataPaths`); **Cancel** discards.
- JSON store only (no Windows registry).

### FR-OPT-5 No side effects on save

- Do not logout, refresh sheets, or call Salesforce on OK.

### FR-OPT-6 Clear metadata cache

- Options dialog provides a **Clear metadata cache for this org** button.
- On click: confirm with the user, then clear L1 (parsed) + L2 (disk) metadata for the **current session’s Salesforce instance URL only** — see [NFR-META-1](./common-performance-requirements.md). Other orgs/sandboxes are left intact.
- Immediate action: does **not** require OK; survives Cancel on other option edits.
- No Salesforce call; show a short success confirmation after clear.
- When not signed in (no instance URL), the button is disabled with guidance to sign in.

---

## Non-functional requirements

### NFR-OPT-1 Thread-safe read

- Data-plane code reads immutable snapshot or thread-safe options per operation to avoid mid-flight toggles.

### NFR-OPT-2 Testability

- `JsonConnectorOptionsStore` round-trip tests (existing in `SalesforceRestAddin.Tests`).

### NFR-OPT-3 Batch size

- Composite **batch size** (default 200) is exposed in the Options dialog — aligns with [common-performance-requirements.md](./common-performance-requirements.md).

---

## Cross-button impact matrix

| Option | Buttons affected |
|--------|------------------|
| `UseReference` | Query table, query rows, update, insert |
| `NoWarning` | Insert confirm; Update (include-hidden path) |
| `ConfirmLargeQuery` | Query table |
| `NoQueryLimit` | Query rows, delete, update |
| `AutoAssignRule` | Insert, update |
| `IncludeHiddenCells` | Update |

---

## Implementation notes

- Core defines options; **WPF** `OptionsWindow` in `SalesforceRestAddin.Windows.Ui` binds to `JsonConnectorOptionsStore`; Excel-DNA reads store when composing `SessionGate` / operation context.
- See [ui-platform.md](./ui-platform.md).
