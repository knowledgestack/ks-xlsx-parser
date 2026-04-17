# Maintainer's Handbook

Notes for whoever has the `Admin` bit on the repo. Everything below is
reversible — err on the side of "turn it on and see what breaks."

## Repository settings to apply before going public

### GitHub → Settings → General

- **Default branch**: `main`
- **Features**: enable **Discussions**, **Issues**, **Projects**, **Wiki (optional)**.
- **Pull Requests**:
  - ✅ Allow squash merging (default)
  - ✅ Allow merge commits
  - ❌ Allow rebase merging (keeps history tidy; squash is preferred)
  - ✅ Automatically delete head branches
  - ✅ Always suggest updating pull request branches

### Settings → Branches → Branch protection rule for `main`

Enable:

- ✅ Require a pull request before merging
  - Required approvals: **1** (raise to 2 once there are multiple maintainers)
  - ✅ Dismiss stale approvals when new commits are pushed
  - ✅ Require review from Code Owners (once `CODEOWNERS` is populated)
- ✅ Require status checks to pass before merging
  - Status checks (once CI runs once on a PR so GitHub knows them):
    - `tests (ubuntu-latest / py3.10)`
    - `tests (ubuntu-latest / py3.11)`
    - `tests (ubuntu-latest / py3.12)`
    - `tests (macos-latest / py3.12)`
    - `testBench round-trip (ubuntu / py3.12)`
  - ✅ Require branches to be up to date before merging
- ✅ Require conversation resolution before merging
- ✅ Require signed commits (soft lock — can relax if it slows contributors)
- ✅ Require linear history
- ✅ Require deployments to succeed before merging (optional, once we have previews)
- ✅ Lock branch (read-only) — skip; we want merges
- ❌ Allow force pushes
- ❌ Allow deletions
- ✅ Do not allow bypassing the above settings (even admins)

### Settings → Actions → General

- Actions permissions: **Allow all actions and reusable workflows**
- Workflow permissions: **Read repository contents and packages permissions**
  (the release workflow requests its own write perms per-job)
- ✅ Require approval for first-time contributors

### Settings → Code security and analysis

- ✅ Dependency graph
- ✅ Dependabot alerts
- ✅ Dependabot security updates
- ✅ Secret scanning
- ✅ Secret scanning push protection
- ✅ Private vulnerability reporting

### Settings → Discussions

Create categories (click *New Category* for each):

- **📣 Announcements** (maintainer-posts only) — releases and project news
- **💡 Ideas** (open) — before-it-becomes-an-issue feature brainstorms
- **🎯 Show and tell** (open) — projects built with XLSXParser
  - Attach the template in `.github/DISCUSSION_TEMPLATE/show-and-tell.yml`
- **🙏 Q&A** (open, answerable) — usage and "does it handle X" questions
- **🧪 testBench findings** (open) — edge cases that shouldn't be issues yet

### Releases

Pushing a `vX.Y.Z` tag triggers `.github/workflows/release.yml` which will:

1. Build the wheel + sdist
2. Build `dist/testBench-v<version>.zip`
3. Attach all three to the GitHub Release
4. Publish to PyPI via Trusted Publishing

One-time PyPI setup: go to PyPI → *your project* → *Publishing* → *Add a new
pending publisher* with:

- Owner: `arnav2`
- Repository name: `XLSXParser`
- Workflow name: `release.yml`
- Environment name: `pypi`

Then create a GitHub environment named `pypi` (Settings → Environments) with
**required reviewers = at least one maintainer** so a release cannot happen
without a human click.

## Release checklist

1. Bump `version` in `pyproject.toml` and `src/xlsx_parser/__init__.py`.
2. `make testbench` → expect 1054/1054.
3. `make test` → clean.
4. Commit with `chore(release): vX.Y.Z`.
5. `git tag -s vX.Y.Z -m "vX.Y.Z"` (signed tag; required by branch protection).
6. `git push && git push --tags` — the tag triggers the release workflow.
7. Approve the `pypi` environment deployment in the Actions tab.
8. Announce in **📣 Announcements**.

## Labels to create

Run once with `gh`:

```bash
gh label create "good first issue" --color 7057ff --description "Ideal for first-time contributors" --force
gh label create "help wanted" --color 008672 --description "Extra attention is needed" --force
gh label create "bug" --color d73a4a --force
gh label create "enhancement" --color a2eeef --force
gh label create "edge-case" --color fbca04 --description "Parser fails/degrades on a specific workbook" --force
gh label create "performance" --color ef6f33 --force
gh label create "docs" --color 0075ca --force
gh label create "needs-triage" --color ededed --force
gh label create "show-and-tell" --color 6366f1 --force
```

## Code owners (once there's more than one maintainer)

Put a `.github/CODEOWNERS` with:

```
# Everything
*                             @arnav2

# Parser internals
/src/xlsx_parser/parsers/     @arnav2
/src/xlsx_parser/formula/     @arnav2
/src/xlsx_parser/analysis/    @arnav2

# Docs
/docs/                        @arnav2
README.md                     @arnav2
```

Add teammates as co-owners as they come on.
