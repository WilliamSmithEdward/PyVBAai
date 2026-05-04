# Dynamic array formulas (FILTER, XLOOKUP, UNIQUE, SORT, SEQUENCE, …)

How PyVBAai writes Excel-365-compatible dynamic array formulas, and the
specific OOXML pieces that openpyxl does not produce.

## The symptom

When openpyxl saves a workbook containing a formula such as
`=UNIQUE(B2:B8)` and you open it in Excel 365, you see one or more of:

- `#NAME?` errors in the formula cell.
- An `@` (implicit-intersection operator) prepended to the formula in the
  formula bar — e.g. `=@UNIQUE(B2:B8)` — and the result fails to spill.
- A "Repaired Records: Cell information from /xl/worksheets/sheet1.xml part"
  dialog when opening the file.
- A "Removed Records: Formula from /xl/worksheets/sheet1.xml part" dialog.

None of these are openpyxl bugs — openpyxl simply writes the formula
verbatim. Excel 365 needs additional OOXML metadata that signals
"this workbook was authored by a dynamic-array-aware application and the
formula was written with spill intent."

## Five things Excel 365 requires

### 1. `_xlfn.` prefix on post-Excel-2010 functions

Excel ≥ 2013 functions must be stored in OOXML as `_xlfn.FUNCNAME(...)`.
Without the prefix Excel cannot resolve the name and shows `#NAME?`.

`_add_xlfn_prefix()` in [core/excel_writer.py](../core/excel_writer.py)
handles this — it runs inside `_set_cell` before the formula reaches
openpyxl. Examples:

| Written by user / LLM       | Stored in XML                       |
|-----------------------------|-------------------------------------|
| `=XLOOKUP(A1,B:B,C:C,"")`   | `=_xlfn.XLOOKUP(A1,B:B,C:C,"")`     |
| `=FILTER(A2:C8,B2:B8>0)`    | `=_xlfn.FILTER(A2:C8,B2:B8>0)`      |
| `=IFERROR(XLOOKUP(...),0)`  | `=IFERROR(_xlfn.XLOOKUP(...),0)`    |
| `=SUM(A1:A10)`              | `=SUM(A1:A10)` (unchanged)          |

#### 1a. Spilled-range operator `A1#` → `_xlfn.ANCHORARRAY(A1)`

The `#` (spill) operator is formula-bar syntax only. Storing it verbatim in
OOXML triggers "Removed Records: Formula" on the next open. It must be
rewritten to its OOXML form:

| Written by user / LLM                              | Stored in XML                                            |
|----------------------------------------------------|----------------------------------------------------------|
| `=XLOOKUP("Apple",O2#,D2:D7,"Not found")`          | `=_xlfn.XLOOKUP("Apple",_xlfn.ANCHORARRAY(O2),D2:D7,"Not found")` |
| `=SUM(Sheet1!A1#)`                                  | `=SUM(_xlfn.ANCHORARRAY(Sheet1!A1))`                     |
| `=SUM($A$1#)`                                       | `=SUM(_xlfn.ANCHORARRAY($A$1))`                          |

`_add_xlfn_prefix()` handles both transformations in a single pass.

### 2. `xcalcf:calcFeatures` extension in `xl/workbook.xml`

This is the workbook-level capability declaration. Without it Excel 365
assumes the file came from a pre-365 application and prepends `@` to every
spill-capable formula to preserve old implicit-intersection behaviour.

Two details that broke this for us initially:

- The XML namespace URL is `…/2018/calcfeatures` (lowercase), not
  `calcFeatures`.
- Feature names use a colon separator: `microsoft.com:RD`, not
  `microsoft.com/RD`.

The exact extension we now write:

```xml
<extLst>
  <ext uri="{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}"
       xmlns:xcalcf="http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures">
    <xcalcf:calcFeatures>
      <xcalcf:feature name="microsoft.com:RD"/>
      <xcalcf:feature name="microsoft.com:Single"/>
      <xcalcf:feature name="microsoft.com:FV"/>
      <xcalcf:feature name="microsoft.com:CNMTM"/>
      <xcalcf:feature name="microsoft.com:LET_WF"/>
      <xcalcf:feature name="microsoft.com:LAMBDA_WF"/>
      <xcalcf:feature name="microsoft.com:ARRAYTEXT_WF"/>
    </xcalcf:calcFeatures>
  </ext>
</extLst>
```

### 3. `xl/metadata.xml` with the XLDAPR cell-metadata type

Each spill-capable cell carries a per-cell metadata index (`cm="1"`). That
index references `cellMetadata > bk > rc` in `xl/metadata.xml`, which
ultimately resolves to `<xda:dynamicArrayProperties fDynamic="1"/>`.

The file is mandatory and openpyxl never produces it. We write it in full,
plus its `[Content_Types].xml` override and `xl/_rels/workbook.xml.rels`
relationship of type `…/relationships/sheetMetadata`.

### 4. `cm="1"` on every dynamic-array anchor `<c>` element

```xml
<c r="F2" cm="1"><f t="array" ref="F2">_xlfn.FILTER(...)</f><v/></c>
```

`cm="1"` is the index (1-based) into the `cellMetadata` block of
`xl/metadata.xml`. Without it Excel does not know the cell was authored
as a dynamic array, even if the calcFeatures extension is present.

### 5. `t="array" ref="<anchor>"` on the `<f>` element

This was the last missing piece and the easiest to get wrong:

- `t="array"` alone (without `ref=`) — Excel reports
  *"Repaired Records: Cell information"* and strips the `cm` attribute.
- No `t="array"` at all — Excel still adds `@` to the formula on open.

`ref` must be a valid range. Excel itself writes the *full* spill range
(e.g. `ref="B1:B3"` for a UNIQUE that spilled to 3 rows) — but we cannot
know the spill size at save time. The fix:

1. Set `ref="<anchor>"` (just the single anchor cell, e.g. `ref="F2"`).
2. Set `<calcPr fullCalcOnLoad="1"/>` in `xl/workbook.xml`.

When Excel opens the file it sees the workbook is dynamic-array-aware
(item 2), the cell is a dynamic-array anchor (items 4 + 5), and a full
recalculation is required (item 5b). It evaluates the formula, expands
the spill range, and rewrites `ref=` to the correct extent on next save.

## Bonus fix: removing `xl/calcChain.xml`

openpyxl writes a stale `xl/calcChain.xml` that references formula cells
in an order that no longer matches the worksheet — particularly with
dynamic arrays. Excel responds with
*"Removed Records: Formula from /xl/worksheets/sheet1.xml part"*.

Solution: drop `xl/calcChain.xml` from the saved zip. Excel regenerates a
correct chain on next open.

**Other causes of "Removed Records: Formula"**

- Storing the `A1#` spill operator verbatim (see §1a above).
- Storing a dynamic-array `<f>` element with `t="array"` but without the
  matching `cm="1"` on the parent `<c>` (see §4 below).

## Round-trip preservation pass

When an existing dynamic-array formula is loaded by openpyxl and saved again,
openpyxl preserves `<f t="array" ref="...">` but silently drops `cm="1"` from
the parent `<c>`. The postprocessor therefore runs two stamping passes:

1. **New formulas** (`_DA_NEW_RE`): cells whose `<f>` has no `t=` attribute —
   these were written by `_set_cell` via openpyxl and need the full treatment.
2. **Existing formulas** (`_DA_EXISTING_RE`): cells whose `<f t="array">`
   survived the round-trip but whose `<c>` lost `cm="1"` — re-add it.

## Known limitation: bare Excel-Table names in dynamic-array formulas

Using the bare table name as an array argument — e.g.
`=FILTER(DataTable, DataTable[Category]="Fruit", "No matches")` — is valid in
Excel's formula bar but fails to resolve during OOXML load-time evaluation,
leaving `#NAME?` until the user manually re-enters the formula. Always use a
column-qualified or `[#Data]` structured reference instead:

```
=FILTER(DataTable[#Data], DataTable[Category]="Fruit", "No matches")
```

## End-to-end flow in `core/excel_writer.py`

```text
apply_changes(file_path, changes)
├─ load workbook
├─ for each change: dispatch to handler
│   └─ _set_cell:  formula → _add_xlfn_prefix() → cell.value
│       _add_xlfn_prefix() does two rewrites in a single pass:
│         • FUNCNAME( → _xlfn.FUNCNAME(  (post-Excel-2010 functions)
│         • A1#       → _xlfn.ANCHORARRAY(A1)  (spilled-range operator)
├─ workbook.calculation.fullCalcOnLoad = True
├─ workbook.save(tmp_path)
└─ _postprocess_xlsx(tmp_path)        # single zip rewrite pass
    ├─ remove xl/calcChain.xml
    ├─ inject xcalcf:calcFeatures into xl/workbook.xml
    ├─ add xl/metadata.xml + Content_Types entry + rels entry (if not already present)
    └─ for each xl/worksheets/sheet*.xml:
        Pass 1 (_DA_NEW_RE):      stamp cm="1" / t="array" ref="<anchor>"
                                  on new dynamic-array <c><f> pairs (no t= attr)
        Pass 2 (_DA_EXISTING_RE): re-add cm="1" to existing <c> elements whose
                                  <f t="array"> survived the round-trip but lost cm
```

## Functions recognised as dynamic arrays

These trigger both `_xlfn.` prefixing and `cm="1"` / `t="array"` stamping:

`FILTER`, `UNIQUE`, `SORT`, `SORTBY`, `SEQUENCE`, `RANDARRAY`,
`XLOOKUP`, `XMATCH`, `LET`, `LAMBDA`, `MAP`, `REDUCE`, `SCAN`,
`MAKEARRAY`, `BYROW`, `BYCOL`.

The `_xlfn.` set is broader and covers every Excel-2013+ function
(see `_XLFN_FUNCTIONS` in `core/excel_writer.py`).

## Reference sources

- A real Excel 365-saved `.xlsx` containing a single `=UNIQUE(...)`
  formula was used as the ground truth for every required XML fragment;
  see commit history for `dynamic_working.xlsx` analysis.
- [Microsoft Open Specifications: Dynamic Arrays](https://learn.microsoft.com/openspecs/office_standards/ms-xlsx/) — `xcalcf` and `xda` namespaces.
- OOXML standard ECMA-376, Part 1, §18 — `<f>`, `<c>`, `cellMetadata`.
