# Architecture

`excel-parser` runs an 8-stage pipeline: **parse → analyse → annotate → segment → render → serialise → verify → compare/export**. The whole graph is deterministic and side-effect-free — you can run the same workbook through it 1,000 times and get the same chunk IDs and hashes.

```mermaid
%%{init: {'theme':'base', 'themeVariables': {
  'primaryColor':'#10B981','primaryTextColor':'#fff','primaryBorderColor':'#047857',
  'lineColor':'#94A3B8','secondaryColor':'#22C55E','tertiaryColor':'#34D399',
  'background':'#FFFFFF','mainBkg':'#10B981','clusterBkg':'#F0FDF4'
}}}%%
flowchart TD
    IN([📄 .xlsx bytes])
    PARSE[["① parsers/<br/>OOXML drivers<br/><i>openpyxl + lxml</i>"]]
    MODELS[["② models/<br/>Pydantic DTOs<br/><i>Workbook · Sheet · Cell · Table · Chart</i>"]]
    FORMULA[["③ formula/<br/>lexer + parser<br/><i>cross-sheet · table · array</i>"]]
    ANALYSIS[["④ analysis/<br/>dependency graph<br/><i>cycles · impact</i>"]]
    CHARTS[["⑤ charts/<br/>OOXML chart extraction"]]
    ANNOT[["⑥ annotation/<br/>semantic roles · KPIs"]]
    SEG[["⑦ chunking/<br/>adaptive segmenter"]]
    REND[["⑧ rendering/<br/>HTML + pipe-text<br/>token counts"]]
    STORE[["🗄️ storage/<br/>JSON · DB rows · vectors"]]
    VER[["✅ verification/<br/>stage assertions"]]
    CMP[["🔀 comparison/<br/>multi-workbook templates"]]
    EXP[["🧬 export/<br/>generated importer"]]
    OUT([🤖 LLM-ready chunks<br/>with citations])

    IN --> PARSE --> MODELS
    MODELS --> FORMULA
    MODELS --> ANALYSIS
    MODELS --> CHARTS
    FORMULA --> ANALYSIS
    ANALYSIS --> ANNOT
    CHARTS --> ANNOT
    ANNOT --> SEG --> REND --> STORE
    MODELS --> VER
    STORE --> OUT
    STORE -.-> CMP -.-> EXP

    %% All-green palette: deepest for entry, lightest for auxiliary stages,
    %% emerald for the headline output node.
    classDef entry   fill:#064E3B,stroke:#022C22,color:#fff,stroke-width:2px;
    classDef parse   fill:#065F46,stroke:#022C22,color:#fff,stroke-width:2px;
    classDef model   fill:#047857,stroke:#064E3B,color:#fff,stroke-width:2px;
    classDef analyze fill:#059669,stroke:#065F46,color:#fff,stroke-width:2px;
    classDef render  fill:#16A34A,stroke:#166534,color:#fff,stroke-width:2px;
    classDef output  fill:#22C55E,stroke:#15803D,color:#fff,stroke-width:2px;
    classDef aux     fill:#A7F3D0,stroke:#047857,color:#065F46,stroke-width:2px;

    class IN entry
    class PARSE parse
    class MODELS model
    class FORMULA,ANALYSIS,CHARTS analyze
    class ANNOT,SEG,REND render
    class STORE,OUT output
    class VER,CMP,EXP aux
```

> The importable module is `excel_parser`; `excel_parser` is a re-export matching the PyPI package name. The package is fully type-annotated (`py.typed` is shipped).

## The 8 stages

| Stage | Module | What it does |
|---|---|---|
| ① Parse | [`parsers/`](../../src/parsers/) | OOXML driver wrapper around `openpyxl` + `lxml`. Emits raw `WorkbookDTO` with cells, merges, hidden rows/cols, conditional formats. |
| ② Models | [`models/`](../../src/models/) | Strict pydantic DTOs for every workbook construct. The contract every downstream stage operates on. |
| ③ Formula | [`formula/`](../../src/formula/) | Lexer + parser for Excel formulas, handling cross-sheet refs, structured-table refs, and array formulas. |
| ④ Analysis | [`analysis/`](../../src/analysis/) | Directed dependency graph between cells, cycle detection, impact analysis. |
| ⑤ Charts | [`charts/`](../../src/charts/) | OOXML chart extraction across 10 chart types (bar/line/pie/scatter/area/radar/bubble/...). |
| ⑥ Annotation | [`annotation/`](../../src/annotation/) | Cell-level semantic roles + KPI detection. Marks header/data/label/output cells. |
| ⑦ Chunking | [`chunking/`](../../src/chunking/) | Adaptive segmenter — connected-components + gap detection + title merging — produces RAG-ready blocks. |
| ⑧ Rendering | [`rendering/`](../../src/rendering/) | HTML and pipe-text rendering per block, token-count estimation, retrieval-friendly raw numeric output. |
| 🗄️ Storage | [`storage/`](../../src/storage/) | Serialiser for JSON / DB rows / vectors. |
| ✅ Verification | [`verification/`](../../src/verification/) | Stage-level invariant assertions — catch parser regressions deterministically. |
| 🔀 Comparison | [`comparison/`](../../src/comparison/) | Compare templates across multiple workbooks to derive a `GeneralizedTemplate`. |
| 🧬 Export | [`export/`](../../src/export/) | Code-generate a Python importer from a generalised template. |

## Where to look next

- **API surface** → [API Reference](API-Reference.md)
- **Stage-by-stage internals** → [Pipeline Internals](Pipeline-Internals.md)
- **DTO field reference** → [Data Models](Data-Models.md)
- **HTTP wrapper** → [Web API](Web-API.md)
