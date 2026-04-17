# Contributing to XLSXParser

**First: welcome.** 👋 If you got here and aren't sure what to do, open a
[Discussion](https://github.com/arnav2/XLSXParser/discussions) — we'd rather
talk than have you leave. Every good-first-issue, every weird `.xlsx` fixture,
every three-line doc patch is welcome.

This project only moves forward because people take 20 minutes to file a good
bug or send a small PR. If that's you, thank you.

## Ways to help (in order of preference for first-time contributors)

1. **Run `make testbench` and report a file that breaks.** We actively want
   edge-case `.xlsx` fixtures — use the
   [Parser edge case issue template](https://github.com/arnav2/XLSXParser/issues/new?template=parser_edge_case.yml).
2. **Add a new workbook to `testBench/`.** Either drop a file under
   `testBench/stress/` or add a builder to `scripts/build_testbench.py`. If
   the parser crashes on it, even better.
3. **Fix one of the flagged issues** in [`docs/PARSER_KNOWN_ISSUES.md`](docs/PARSER_KNOWN_ISSUES.md).
4. **Improve docs.** The README, the architecture diagram, the examples —
   if something confused you, it confuses everyone.
5. **Open a [Show & Tell](https://github.com/arnav2/XLSXParser/discussions/new?category=show-and-tell)**
   if you shipped something with the parser. Seriously, it helps us prioritise.

## Development setup

```bash
git clone https://github.com/arnav2/XLSXParser.git
cd XLSXParser
make install               # pip install -e ".[dev,api]"
make test                  # fast, default suite
make testbench-build       # regenerate 1000-file stress corpus (~1 min)
make testbench             # round-trip every workbook; parallel
```

Prerequisites: Python 3.10+, `pip`, optionally `make`. We use `ruff` for
linting/formatting — install it with the `[dev]` extra.

## Pull-request checklist

Your PR should:

1. Have tests. `pytest` must stay green: `make test`.
2. Keep `make testbench` at 1054/1054 (or explain the delta in the PR description).
3. Pass `ruff check` (`make lint`) and be formatted with `make format`.
4. Include one sentence in the PR description that starts with *"This change…"*.
5. Use [conventional-commit style](https://www.conventionalcommits.org/)
   commit messages: `feat:`, `fix:`, `perf:`, `refactor:`, `docs:`, `test:`,
   `chore:`.

We lean toward **smaller PRs with more context** over big bundles. A five-line
fix with a one-paragraph explanation is almost always mergeable.

## Reporting issues

Use the [issue templates](https://github.com/arnav2/XLSXParser/issues/new/choose).
For security issues, please use the
[private advisory flow](https://github.com/arnav2/XLSXParser/security/advisories/new)
— not a public issue.

Helpful things to include:

- Output of `python -c "import xlsx_parser; print(xlsx_parser.__version__)"`
- Python version (`python --version`)
- OS
- Minimal `.xlsx` that reproduces the bug (or a generator that builds one)
- Full traceback

## Code style at a glance

- Type hints everywhere that's practical.
- Tests live in `tests/`; programmatic workbook fixtures live in `tests/conftest.py`.
- Cross-validation against calamine uses the `crossval` marker.
- Long-running bench tests use `@pytest.mark.testbench` and are skipped by default.
- Keep public-API changes additive; if you can't, note it in the PR and the
  maintainers will line up the deprecation.

## Community

- Discussions: <https://github.com/arnav2/XLSXParser/discussions>
- Issues: <https://github.com/arnav2/XLSXParser/issues>
- Security: <https://github.com/arnav2/XLSXParser/security/advisories>

By participating you agree to follow our [Code of Conduct](CODE_OF_CONDUCT.md).

## Thanks

Really. Every contribution makes this project sustainable.
