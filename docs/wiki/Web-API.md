# Web API

`ks-xlsx-parser` ships a FastAPI application with a drag-and-drop UI, so
any service that can hit HTTP can use the parser without a Python dep.

## Install

```bash
pip install "ks-xlsx-parser[api]"
```

The `[api]` extra pulls in `fastapi`, `uvicorn[standard]`, and
`python-multipart`.

## Run it

```bash
# console entry point, listens on :8080
xlsx-parser-api

# or directly with uvicorn
uvicorn xlsx_parser.api:app --reload --port 8080
```

Open <http://localhost:8080> for the drag-and-drop UI.

## Endpoints

### `POST /parse`

Parse an uploaded `.xlsx` and return the full result.

**Request** — `multipart/form-data` with a single `file` field:

```bash
curl -X POST http://localhost:8080/parse \
     -F "file=@workbook.xlsx"
```

**Response** — JSON:

```jsonc
{
  "parse_result": {
    "workbook": { /* full WorkbookDTO */ },
    "chunks":   [ /* ChunkDTOs */ ]
  },
  "verification_markdown": "# Verification report\n…",
  "verification": { /* structured VerificationReport */ }
}
```

**Python client:**

```python
import requests

with open("workbook.xlsx", "rb") as f:
    r = requests.post("http://localhost:8080/parse", files={"file": f})
data = r.json()

for chunk in data["parse_result"]["chunks"]:
    print(chunk["source_uri"], chunk["token_count"])
```

**TypeScript / Node client:**

```typescript
import { readFile } from "node:fs/promises";

const buf = await readFile("workbook.xlsx");
const body = new FormData();
body.append("file", new Blob([buf]), "workbook.xlsx");

const res = await fetch("http://localhost:8080/parse", { method: "POST", body });
const data = await res.json();

for (const chunk of data.parse_result.chunks) {
  console.log(chunk.source_uri, chunk.token_count);
}
```

### `GET /` — UI

Drag-and-drop browser UI for ad-hoc parsing. Useful for triage but not
intended as a production endpoint.

### `GET /healthz`

Returns `{"status": "ok"}`. Point your load balancer here.

## Deployment notes

- **Stateless** — no DB, no disk writes. Horizontal scale at will.
- **Timeouts** — set your reverse proxy's upstream timeout above
  `max_cells_per_sheet` × your parse rate. A 2 M-cell sheet parses in
  single-digit seconds on a modern CPU; add margin for CF / formula
  evaluation on dense models.
- **Memory** — peak ~ `cells × ~1 KB`. A 2 M-cell sheet peaks around 2 GB.
- **No macros, no external links** — the parser never executes workbook
  content. Safe to run against untrusted uploads.
- **ZIP bomb protection** — incoming `.xlsx` is size-checked before
  openpyxl sees it.

## Running inside another FastAPI app

```python
from fastapi import FastAPI
from ks_xlsx_parser.api import app as xlsx_app

app = FastAPI()
app.mount("/xlsx", xlsx_app)
```

Now `POST /xlsx/parse` behaves the same as `POST /parse` above.

## Configuration

The API respects a handful of environment variables:

| Variable | Default | Purpose |
|---|---|---|
| `XLSX_PARSER_MAX_CELLS` | `2000000` | Per-sheet cell cap passed to `parse_workbook`. |
| `XLSX_PARSER_MAX_FILE_MB` | `100` | Reject uploads larger than this before parsing. |
| `XLSX_PARSER_PORT` | `8080` | Port the console entry point listens on. |

Production users typically front it with Nginx or Caddy for TLS + auth.

## MCP server (roadmap)

An MCP server wrapping the same parse surface is on the roadmap so that
Claude Desktop, Cursor, Windsurf, and Zed can call it without any glue
code. Track progress or vote on the
[roadmap discussion](https://github.com/knowledgestack/ks-xlsx-parser/discussions).
