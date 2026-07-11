# Describe Sforce Object

**Ribbon:** `btnDescribeSobject` — “Describe Sforce Object”  
**Login required:** Yes  
**COM / VBA:** None

See also: [common-performance-requirements.md](./common-performance-requirements.md)

---

## Summary

User picks one or more Salesforce objects; add-in creates a **new worksheet per object** with a field metadata grid (API names, types, lengths, flags, picklists, etc.). Used to design valid ForceConnector tables and understand object shape.

---

## Functional requirements

### FR-DSO-1 Session

- Require authenticated session before object picker.

### FR-DSO-2 Object selection UI

- WPF modal window listing queryable objects (search/filter) — `DescribeObjectPickerWindow` in `SalesforceRestAddin.Windows.Ui`.
- The picker uses a WPF grid with the same shared sort rule as the object picker elsewhere in the add-in.
- The dialog keeps a visible footer with `OK` and `Cancel` buttons even when the list is short.
- User may select **multiple** objects.
- No translation language picker — Translation Helper / METAAPI is out of scope ([AGENTS.md](../../AGENTS.md)).

### FR-DSO-3 Metadata source

| Source | Implementation |
|--------|----------------|
| Object / field metadata | **REST** `GET .../sobjects/{name}/describe` |
| Field labels | Standard labels from describe (no METAAPI translations) |

### FR-DSO-4 Worksheet output

Per selected object:

1. Create new worksheet (name derived from object label/API — avoid invalid sheet name characters).
2. Write title/header rows identifying object.
3. Write field grid with columns covering at least:

| Column concept | Content |
|----------------|---------|
| Label | `label` |
| API name | `name` |
| Type | `type` |
| Length / precision | as applicable |
| Required | `nillable`, `createable`, `updateable` |
| Custom | `custom` |
| Reference targets | `referenceTo` |
| Picklist values | joined list for picklists |

4. **Field ordering:** well-known standard fields first → remaining standard → custom (`__c`).

### FR-DSO-5 Progress and cancel

- Long describe for many objects: progress UI; cancel between objects.

### FR-DSO-6 Errors

- Per-object failure should not block other objects; summarize failures at end.

---

## Non-functional requirements

### NFR-DSO-1 Bulk value write

- Build `object[numFields, numCols]` then assign with one `rng.Value = data`.

### NFR-DSO-2 Core ownership

- Parse `DescribeSObjectResult` JSON into row DTOs in Core.
- ExcelDna only maps DTOs → `object[,]`.

### NFR-DSO-3 Caching

- Optional: cache describe per object for session to speed wizard + describe (shared metadata cache — NFR-META-1).

### NFR-DSO-4 Testability

- Golden-file tests: describe JSON → ordered field rows (no Excel).

---

## Salesforce API

| Call | Purpose |
|------|---------|
| `GET .../sobjects/` | Object list for picker |
| `GET .../sobjects/{name}/describe` | Field metadata |

---

## Out of scope

- Translation Helper metadata columns
- Managed-data / translation-only options
- Downloading translations for describe grid
- Per-cell comments on the describe grid — bulk values only

## Relation to Table Wizard

Wizard uses live describe for field pickers; this feature exports a **human-readable catalog** on a dedicated sheet.
