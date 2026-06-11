# Baquiano — Search Data

Search and select Rhino objects by their User Keys/Values using include/exclude conditions.

## Search Scope

- **Search all objects in model** — searches every object in the document
- **Search only pre-selected objects** — available when objects are selected before running; shows the count

## Building Conditions

**Include conditions (AND logic)** — objects must match ALL include conditions.

**Exclude conditions (OR logic)** — objects matching ANY exclude condition are removed from results.

For each condition: type the Key name, the Value to match, and select a Match Type. At least one include condition with both Key and Value is required.

## Match Types

| Match Type | Description |
| --- | --- |
| Contains | Value appears anywhere in the key's value |
| Equals | Exact match (case-insensitive) |
| Starts with | Key's value begins with the search value |
| Ends with | Key's value ends with the search value |
| Does not contain | Negation of Contains |
| Does not equal | Negation of Equals |
| Does not start with | Negation of Starts with |
| Does not end with | Negation of Ends with |

## Examples

**Cross-search** — objects where `Name` contains `House` but `User` is not `university`:

1. Include: Key = `Name`, Value = `House`, Match = `Contains`
2. Exclude: Key = `User`, Value = `university`, Match = `Equals`

**Finding outliers** — objects where `Status` is not `Approved`:

1. Include: Key = `Status`, Value = `Approved`, Match = `Does not equal`

## Results

Matching objects are selected in the viewport. A summary shows match count, conditions used, and search scope. Cancelling restores the original selection.

## Troubleshooting

**No results** — Key names are case-sensitive; use `GetUserText` in Rhino to verify existing keys; try `Contains` instead of `Equals` for a broader match.
