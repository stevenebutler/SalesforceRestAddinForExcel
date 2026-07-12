# Query Criteria Operators

Criteria live in row 1 of a ForceConnector table. Starting in column B, each clause uses three adjacent cells:

```text
[field label or API name] | [operator] | [value]
```

For example, `Account Name | like | Acme` produces `Name like '%Acme%'`. Field labels resolve through Salesforce describe metadata; API names such as `ParentId` and `CreatedDate` are also accepted. Operators are case-insensitive.

## Wizard Operators

| Operator | Result | Sample use-case | Example condition |
|---|---|---|---|
| `equals` | `=` | Accounts in one industry | `Industry = 'Technology'` |
| `not equals` | `!=` | Exclude closed cases | `Status != 'Closed'` |
| `like` | `LIKE`, adds `%` on both sides | Contacts with `Smith` anywhere in the name | `Name like '%Smith%'` |
| `starts with` | `LIKE`, adds `%` at the end | Accounts beginning `Acme` | `Name like 'Acme%'` |
| `ends with` | `LIKE`, adds `%` at the start | Cases ending in `renewal` | `Subject like '%renewal'` |
| `less than` | `<` | Opportunities below a threshold | `Amount < 1000.0` |
| `greater than` | `>` | Opportunities above a threshold | `Amount > 10000.0` |
| `includes` | `includes` for multi-select picklists | Records tagged `APAC` | `Regions__c includes ('APAC')` |
| `excludes` | `excludes` for multi-select picklists | Exclude records tagged `APAC` | `Regions__c excludes ('APAC')` |
| `regexp` | Same wildcard `LIKE` behavior as `like` | Historical ForceConnector-compatible text search | `Name like '%Acme%'` |

`regexp` is a historical label, not regular-expression support. `begins with` and `contains` are accepted in existing sheets as aliases for `starts with` and `like`; they are not wizard choices.

## Values And Field Types

### Multiple Values

For ordinary criteria, comma-separated values become an OR group. For example:

```text
B1: Industry   C1: equals   D1: Technology,Finance
```

becomes:

```sql
(Industry = 'Technology' or Industry = 'Finance')
```

### Empty Values

An empty text value becomes `''`; an empty date/datetime becomes `null`. The wizard default, `Record Id | not equals | [empty]`, means all records with an Id.

```text
Parent Account | equals | [empty]  -> ParentId = ''
Close Date     | equals | [empty]  -> CloseDate = null
```

### Numbers, Booleans, And Dates

Numbers and booleans are unquoted. Currency, percent, and double values retain a decimal place where needed; booleans become `TRUE` or `FALSE`.

```text
Amount    | greater than | 10000  -> Amount > 10000.0
Is Active | equals       | true   -> IsActive = TRUE
```

For date/datetime criteria, use a native Excel date cell, an ISO string, or a Salesforce relative literal. Locale-formatted date text is rejected as ambiguous.

```text
Close Date | greater than | 2026-07-01  -> CloseDate > 2026-07-01
Created On | equals       | TODAY       -> CreatedDate = TODAY
```

`like` is invalid for a single-select picklist; use `equals` or `not equals`. Use `includes` and `excludes` for multi-select picklists. For ordinary reference criteria, use Salesforce Ids: displayed reference names are not currently resolved by the shipped Excel host.

## Hidden ForceConnector Reference-List Syntax

`in` is deliberately absent from the wizard but remains available in existing ForceConnector-style sheets for reference fields. Its value must be an Excel range address or named range containing Salesforce Ids, not a comma-separated list.

Example with a named range `CustomerIds`:

```text
B1: Parent Account   C1: in   D1: CustomerIds

CustomerIds:
001000000000001
001000000000002
001000000000003
```

Empty values, duplicates, and values that are not 15- or 18-character alphanumeric Salesforce Ids are ignored. If no Id remains, Query Table Data reports an error on the criteria value cell. The IDs are read on the Excel thread and translated into one or more batched queries such as:

```sql
WHERE ParentId IN ('001000000000001', '001000000000002', '001000000000003')
```

Large lists are split by the configured composite batch size; there is never one Salesforce request per Id.

### Why `on` Is Rejected

`on` is not supported. Query Table Data reports:

```text
ON is not supported - use IN with a range reference to select multiple items.
```

`on` is neither a SOQL `JOIN ... ON` operation nor a one-row-per-Id mode. The old ForceConnector code set an unused mode flag, then built `ReferenceField = Id1 AND ReferenceField = Id2` for a multi-Id range, which normally returns no records. Use `in` with the same range/name when the intended meaning is "match any of these reference Ids."

## Combining Clauses

Successive triplets are combined with `and`.

```text
B1: Industry        C1: equals       D1: Technology
E1: Annual Revenue  F1: greater than G1: 1000000
```

produces:

```sql
Industry = 'Technology' and AnnualRevenue > 1000000.0
```

See [legacy-criteria-compatibility.md](./legacy-criteria-compatibility.md) for the compatibility matrix and [query-table-data.md](./ribbon/query-table-data.md) for the complete query specification.
