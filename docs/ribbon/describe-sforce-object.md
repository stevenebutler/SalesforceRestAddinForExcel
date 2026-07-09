# Describe Sforce Object

**Ribbon:** `btnDescribeSobject` — “Describe Sforce Object”  
**Legacy:** `DescribeCustomObject.DescribeSalesforceObjectsBySOAP()` → `processDescribeCustomObject`  
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
- User may select **multiple** objects.
- **Dropped from legacy:** translation language picker (`METAAPI.getTranslations`) — not ported.

### FR-DSO-3 Metadata source

| Legacy | New implementation |
|--------|-------------------|
| SOAP `DescribeSObject` | **REST** `GET .../sobjects/{name}/describe` |
| METAAPI field translations | **Dropped** — use standard labels from describe |

### FR-DSO-4 Worksheet output

Per selected object:

1. Create new worksheet (name derived from object label/API — avoid invalid sheet name characters).
2. Write title/header rows identifying object.
3. Write field grid with columns comparable to legacy (~13 columns), minimally:

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

4. **Field ordering:** named standard fields (legacy fixed set) → remaining standard → custom (`__c`).

### FR-DSO-5 Progress and cancel

- Long describe for many objects: progress UI; cancel between objects.

### FR-DSO-6 Errors

- Per-object failure should not block other objects; summarize failures at end.

---

## Non-functional requirements

### NFR-DSO-1 Bulk value write

**Legacy good pattern:** build `object[numFields, numCols]` then `rng.Value = data` once.

**Target:** preserve single bulk assign for the main grid.

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

## Legacy reference

- `ForceConnector/DescribeCustomObject.cs`
- `ForceConnector/processDescribeCustomObject.cs`
- `SOAPAPI.DescribeSObject` — **replace with REST**

## Out of scope

- Translation Helper metadata columns
- `GET_MANAGED` option (translation-only in legacy Options UI)
- Downloading translations for describe grid
- Per-cell comments on the describe grid (legacy `renderComments`) — bulk values only

## Relation to Table Wizard

Wizard uses live describe for field pickers; this feature exports a **human-readable catalog** on a dedicated sheet.
