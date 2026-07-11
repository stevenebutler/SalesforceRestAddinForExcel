# Legacy Criteria Compatibility

This document captures the legacy `ForceConnector/ForceConnector` Table Query Wizard
criteria/operator behavior so the new add-in can preserve workbook compatibility.

## Legacy wizard operator list

The legacy wizard dropdown exposed:

- `equals`
- `not equals`
- `like`
- `starts with`
- `ends with`
- `less than`
- `greater than`
- `includes`
- `excludes`
- `regexp`

It did **not** show `in` or `on` as visible wizard options.

## Operator mapping

| Wizard operator | Legacy SOQL behavior | New add-in behavior |
|---|---|---|
| `equals` | Emits `=` | Emits `=` |
| `not equals` | Emits `!=` | Emits `!=` |
| `like` | Wraps the value as `%value%` and emits `LIKE` | Same |
| `starts with` | Appends `%` and emits `Like` | Normalized to `begins with`, emits `like 'value%'` |
| `begins with` | Not a visible legacy wizard label, but parser accepted it as the same as `starts with` | Accepted as an alias for `starts with` |
| `ends with` | Prepends `%` and emits `Like` | Emits `like '%value'` |
| `less than` | Emits `<` | Emits `<` |
| `greater than` | Emits `>` | Emits `>` |
| `includes` | Passed through as `includes` | On multipicklist fields, treated as `includes`; otherwise passed through |
| `excludes` | Passed through as `excludes` | On multipicklist fields, treated as `excludes`; otherwise passed through |
| `regexp` | Rewritten to `like`, with the user supplying wildcard-style input | Rewritten to `like` |
| `contains` | Not a legacy wizard label | Accepted by the new parser as an alias for `like` |
| `in` | Not a visible legacy wizard option; on reference fields, legacy treated saved `in` clauses as a reference join | On reference fields, parsed as `IN (...)` from a comma-separated Id list |
| `on` | Not a visible legacy wizard option; on reference fields, legacy treated saved `on` clauses as one-row-per-reference join mode | On reference fields, parsed as join mode and batched by Id list |

## Value formatting

Legacy and new code both format values by Salesforce field type, but the source of truth differs:

- `date` / `datetime`
  - Legacy: `QueryValueFormat(...)` formats from the Excel cell value, or from strings like `today`, `today + 30`, `today - 1`
  - New: accepts Excel `DateTime` / OA date serials, ISO date strings, and relative SOQL literals like `TODAY`, `LAST_N_DAYS`, `THIS_MONTH`
- `double` / `currency` / `percent`
  - Legacy: emits numeric text, adding `.0` when needed
  - New: emits numeric text via `FormatNumericLiteral(...)`
- `boolean`
  - Legacy: emits `TRUE` / `FALSE`
  - New: emits `TRUE` / `FALSE`
- `int`
  - Legacy: emits raw numeric text
  - New: emits raw numeric text
- string-like fields, including picklist / id / reference
  - Legacy: quotes the value
  - New: quotes the value, with multipicklist special-casing

## Reference resolution

- Legacy only converts names to Ids when the `USE_REFERENCE` option is enabled.
- New code only resolves names to Ids when `ConnectorOptions.UseReference` is enabled and an `IReferenceResolver` is provided.

## Sources

- Legacy wizard operator list: `ForceConnector/ForceConnector/WizardStep4.Designer.cs`
- Legacy clause assembly: `ForceConnector/ForceConnector/Operation.cs`
- Legacy value formatting: `ForceConnector/ForceConnector/Util.cs`
- New parser: `src/SalesforceRestAddin.Core/Soql/SoqlCriteriaParser.cs`
- New query builder: `src/SalesforceRestAddin.Core/Soql/SoqlQueryBuilder.cs`
