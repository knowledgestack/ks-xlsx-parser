# Security Policy

## Supported Versions

We provide security fixes for the latest released version on PyPI.

| Version | Supported |
| ------- | --------- |
| Latest  | ✅        |
| Older   | ❌        |

## Reporting a Vulnerability

**Please do not report security vulnerabilities through public GitHub issues.**

Instead, report them via GitHub's [Private Vulnerability Reporting](https://github.com/knowledgestack/ks-xlsx-parser/security/advisories/new)
feature. This lets us triage privately before disclosing.

If you cannot use GitHub's private reporting, email the maintainers
(address in `pyproject.toml`'s `authors` field) with:

1. A description of the issue
2. Steps to reproduce (ideally a minimal `.xlsx` fixture)
3. The affected version(s)
4. Any suggested mitigation

We aim to acknowledge every report within **72 hours** and will keep you
informed throughout the triage.

## What counts as a vulnerability

ks-xlsx-parser processes untrusted `.xlsx` input, so we treat the following as
in-scope:

- Arbitrary code execution via a crafted workbook
- Denial-of-service (zip bomb, memory exhaustion, pathological formula chains)
  that cannot be prevented by documented limits
- Path traversal or SSRF via external references
- Infinite loops in parser or dependency-graph code on well-formed input

Out of scope:

- Bugs triggered only by workbooks larger than documented `max_cells_per_sheet`
- Slow parse times on genuinely huge workbooks (we provide timeouts and limits)

## Disclosure

Once a fix is ready, we publish a GitHub Security Advisory with a CVE if
applicable and credit the reporter (unless anonymity is preferred).
