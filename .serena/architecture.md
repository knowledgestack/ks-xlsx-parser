# XLSXParser — Architecture Document

*Production enterprise-focused auditable Excel parser for validation and deployment.*

---

## 1. Purpose

Parse `.xlsx` workbooks into structured, loss-minimizing representations with full auditability. The system **must not miss any data** in Excel files. It combines:

- **Traditional (deterministic) parsing**: openpyxl, lxml, direct OOXML
- **LLM-assisted processing**: stage-specific skills for disambiguation, validation, and enrichment

Output is chunked, traceable, and exportable as JSON for downstream use and user validation.

---

## 2. High-Level Architecture

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                    XLSXParser — Enterprise Auditable Parser                   │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│   ┌────────┐   ┌────────┐   ┌──────────┐   ┌─────────┐   ┌───────┐   ┌─────┐│
│   │  Load  │──▶│ Parse  │──▶│ Normalize│──▶│ Segment │──▶│ Render│──▶│Store││
│   └────────┘   └────────┘   └──────────┘   └─────────┘   └───────┘   └─────┘│
│       │            │              │              │            │         │   │
│       ▼            ▼              ▼              ▼            ▼         ▼   │
│   [LLM Skill] [LLM Skill]   [LLM Skill]   [LLM Skill] [LLM Skill] [LLM Skill]│
│                                                                              │
├─────────────────────────────────────────────────────────────────────────────┤
│  Temporal Workflows │ Multiple async Excel files │ Validation loops          │
├─────────────────────────────────────────────────────────────────────────────┤
│  Output: JSON export │ Chunked for RAG │ Document for user validation        │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## 3. Pipeline Stages (with LLM Skills)

Each stage has a dedicated, detailed LLM skill to perform that task:

| Stage | Traditional Role | LLM Skill Role |
|-------|------------------|----------------|
| **Load** | Compute hash, open workbook, detect format | Validate file integrity, flag anomalies |
| **Parse** | Extract cells, formulas, charts, tables, comments | Resolve ambiguity, classify edge cases |
| **Normalize** | Merges, bounding boxes, dependency graph | Semantic merge validation |
| **Segment** | Gap-based blocks, table detection | Semantic block boundaries, context |
| **Render** | HTML/text output, token counts | Content summarization, coherence check |
| **Store** | JSON serialization, chunk hashes | Schema validation, completeness check |

---

## 4. Operational Model

### Temporal Integration

- Parser runs as Temporal activities/workflows
- Multiple Excel files processed asynchronously in parallel
- Per-file isolation: hash → parse → chunk → export
- Workflow-level orchestration for batch jobs

### Testing & Validation

- **Stage verification**: Maps to Excellent Algorithm (11 stages)
- **Validation loops**: Re-run parsing and compare outputs to ensure correctness
- **Determinism checks**: Same input → same output (hash-stable)
- **Coverage metrics**: Block counts, cell counts, formula counts per stage

### Final Deliverable

- Consolidated document with:
  - Workbook metadata (hash, filename, sheets, cells)
  - Chunk listing (source_uri, block_type, coordinates, token_count)
  - Rendered content per chunk
  - Error/warning summary
  - Exportable JSON for programmatic access

**User validation**: The document is the primary artifact for human review before downstream use.

---

## 5. Data Flow

```
.xlsx file (path or bytes)
    │
    ▼
WorkbookParser.parse() → WorkbookDTO
    │
    ▼
ChunkBuilder.build_all() → list[ChunkDTO]
    │
    ▼
WorkbookSerializer → JSON / Postgres / Vector store
    │
    ▼
ParseResult.to_json() → Final document for validation
```

---

## 6. Key Principles

| Principle | Implementation |
|-----------|----------------|
| **No data loss** | Partial parse with error collection; extract maximum even from malformed workbooks |
| **Auditability** | `workbook_hash`, `chunk_id`, `content_hash`; traceable to sheet/row/col |
| **Determinism** | xxhash64 for hashes; identical input → byte-identical output |
| **Parallelism** | Sheet-level independence; thread-safe parsing |
| **Enterprise-grade** | Structured logging, configurable limits, redaction hooks |

---

## 7. Export & Access

- **Primary API**: `parse_workbook(path | content) → ParseResult`
- **JSON export**: `result.to_json()` → dict with `workbook`, `chunks`, `total_chunks`, `total_tokens`
- **Storage**: Postgres (workbooks, sheets, blocks, cells, charts) + vector store (embeddings)
- **Chunk format**: `chunk_id`, `source_uri`, `sheet_name`, `block_type`, `top_left`, `bottom_right`, `render_text`, `token_count`

---

## 8. LLM Skill Structure (Per Stage)

Each stage skill should specify:

1. **Input**: What the traditional stage produced
2. **Task**: Specific validation, disambiguation, or enrichment
3. **Output**: Structured result (e.g., updated block classification, confidence score)
4. **Constraints**: No hallucination; only operate on extracted data

---

*This document is the architecture reference for XLSXParser. Use it for validation, onboarding, and Temporal workflow design.*
