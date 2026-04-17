# API Reference

The public surface re-exported from both `xlsx_parser` and `ks_xlsx_parser`:

```python
from ks_xlsx_parser import (
    parse_workbook,      # single file  → ParseResult
    compare_workbooks,   # N files      → GeneralizedTemplate
    export_importer,     # template     → generated Python class
    ParseResult,
    StageVerifier,       # per-stage debugging
    VerificationReport,
    ExcellentStage,
    __version__,
)
```

The package is fully type-annotated; `py.typed` is shipped.

## `parse_workbook()`

Parse a single Excel workbook.

```python
def parse_workbook(
    path: str | Path | None = None,
    content: bytes | None = None,
    filename: str | None = None,
    max_cells_per_sheet: int = 2_000_000,
) -> ParseResult: ...
```

| Argument | Type | Default | Purpose |
|---|---|---|---|
| `path` | `str \| Path \| None` | `None` | Path to a `.xlsx` / `.xlsm` file. Either `path` or `content` is required. |
| `content` | `bytes \| None` | `None` | Raw file bytes. Use when reading from an HTTP upload or S3 object. |
| `filename` | `str \| None` | `None` | Display name to attach to the result (shown in source URIs and logs). Defaults to `path.name` if `path` is set, else `"<in-memory>"`. |
| `max_cells_per_sheet` | `int` | `2_000_000` | Safety ceiling. Sheets with more cells are truncated with a `WARNING`-level `ParseError` on the result. |

**Returns:** [`ParseResult`](Data-Models#parseresult).

**Raises:** never — load errors become `ParseError` entries on
`result.workbook.errors` so a single bad file can't bring down a pipeline.

**Example — from a filename:**

```python
result = parse_workbook(path="workbook.xlsx")
```

**Example — from bytes:**

```python
with open("workbook.xlsx", "rb") as f:
    content = f.read()
result = parse_workbook(content=content, filename="workbook.xlsx")
```

## `compare_workbooks()`

Align multiple workbooks that share a template (e.g. Q1/Q2/Q3 reports) and
compute a `GeneralizedTemplate` capturing structural similarities and
degrees-of-freedom.

```python
def compare_workbooks(
    paths: list[str | Path],
    dof_threshold: int = 50,
) -> GeneralizedTemplate: ...
```

| Argument | Type | Default | Purpose |
|---|---|---|---|
| `paths` | `list[str \| Path]` | required | Two or more workbooks to align. |
| `dof_threshold` | `int` | `50` | Cells with more than this many unique values across inputs are marked as free-form data rather than fixed-template. |

**Returns:** `GeneralizedTemplate` — a tree of `TemplateNode` objects.

**Example:**

```python
from ks_xlsx_parser import compare_workbooks

template = compare_workbooks(
    ["report_q1.xlsx", "report_q2.xlsx", "report_q3.xlsx"],
    dof_threshold=50,
)
```

## `export_importer()`

Generate a reusable Python importer class from a generalised template.

```python
def export_importer(
    template: GeneralizedTemplate,
    output_path: str | Path,
    class_name: str = "GeneratedImporter",
) -> Path: ...
```

| Argument | Type | Default | Purpose |
|---|---|---|---|
| `template` | `GeneralizedTemplate` | required | Output of `compare_workbooks()`. |
| `output_path` | `str \| Path` | required | File to write. |
| `class_name` | `str` | `"GeneratedImporter"` | Name of the generated class. |

**Returns:** the `Path` written to.

**Example:**

```python
from ks_xlsx_parser import compare_workbooks, export_importer

template = compare_workbooks(["q1.xlsx", "q2.xlsx", "q3.xlsx"])
export_importer(template, "quarterly_importer.py",
                class_name="QuarterlyReportImporter")
```

The generated class has one `import_one(path: str) -> QuarterlyReport`
method that pulls the same fields from every future workbook matching
the template.

## `StageVerifier`

Step-by-step debugging of the parse pipeline.

```python
from ks_xlsx_parser import StageVerifier, ExcellentStage

verifier = StageVerifier(path="workbook.xlsx")
report = verifier.run()

for stage in ExcellentStage:
    stage_result = report.get_stage(stage)
    print(stage.value, stage_result.ok, stage_result.duration_ms)

print(report.to_markdown())   # human-readable summary
```

`ExcellentStage` is an enum of the 11 stages in the pipeline (see
[Pipeline Internals](Pipeline-Internals)). Each stage produces a
`StageResult` with:

- `stage` — which stage
- `ok` — did it pass invariants?
- `duration_ms` — wall time
- `diagnostics` — structured list of issues found
- `output_summary` — one-line description of what the stage produced

## CLI

The package also installs an `xlsx-parser-api` console entry point that
launches the FastAPI web server — see the [Web API](Web-API) page.

## Import paths

Two module names point at the same package:

- `from xlsx_parser import ...` — original import path.
- `from ks_xlsx_parser import ...` — alias matching the PyPI
  distribution name (dashes normalised to underscores).

Use whichever reads better. Both will always work.
