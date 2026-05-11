# Release process

This document is the **operational** companion to [`.github/workflows/release.yml`](../.github/workflows/release.yml). The workflow is tag-triggered (`v*.*.*`); pushing such a tag builds wheel + sdist, creates a GitHub Release, and publishes to PyPI. **All three actions are partially or fully irreversible** — PyPI in particular does not allow re-publishing a version. Run through this checklist before tagging.

## One-time setup

These have to be done once per repository and persist across releases.

### 1. GitHub Environment: `pypi`

The release workflow's `pypi` job declares `environment: pypi`. That environment must exist on the repo, otherwise the OIDC token exchange with PyPI's Trusted Publisher endpoint fails.

```bash
# Create the empty environment (idempotent)
gh api -X PUT repos/knowledgestack/ks-xlsx-parser/environments/pypi
```

Optional but recommended after creation: open https://github.com/knowledgestack/ks-xlsx-parser/settings/environments/pypi and add a **required reviewer** so a tag push needs explicit approval before the PyPI publish step runs. This is a safety net — once a tag is pushed the release workflow auto-fires; a reviewer gate gives you one last "are you sure?" before the irreversible PyPI publish.

### 2. PyPI Trusted Publisher binding

This **cannot** be done via API or a PR — it requires logging into pypi.org as a maintainer of `ks-xlsx-parser`.

1. Go to https://pypi.org/manage/project/ks-xlsx-parser/settings/publishing/
2. Click **Add a new publisher** → **GitHub**
3. Fill in:
   - **PyPI Project Name:** `ks-xlsx-parser`
   - **Owner:** `knowledgestack`
   - **Repository:** `ks-xlsx-parser`
   - **Workflow filename:** `release.yml`
   - **Environment name:** `pypi`
4. Save.

Verify with:

```bash
# Should list any publishers tied to the project (requires you to be logged in)
open "https://pypi.org/manage/project/ks-xlsx-parser/settings/publishing/"
```

If you're spinning up a new project that doesn't exist on PyPI yet, the publisher has to be configured as a **pending publisher** under your account first. Same form, accessible at https://pypi.org/manage/account/publishing/.

### 3. Branch protection on `main`

CI on a PR validates {ubuntu, macOS} × Python {3.10, 3.11, 3.12} before merge. Add a branch protection rule on `main` requiring those status checks to pass before a PR is mergeable:

```bash
gh api -X PUT repos/knowledgestack/ks-xlsx-parser/branches/main/protection \
  -F required_status_checks[strict]=true \
  -F 'required_status_checks[contexts][]=tests (ubuntu-latest / py3.10)' \
  -F 'required_status_checks[contexts][]=tests (ubuntu-latest / py3.11)' \
  -F 'required_status_checks[contexts][]=tests (ubuntu-latest / py3.12)' \
  -F 'required_status_checks[contexts][]=tests (macos-latest / py3.10)' \
  -F 'required_status_checks[contexts][]=tests (macos-latest / py3.11)' \
  -F 'required_status_checks[contexts][]=tests (macos-latest / py3.12)' \
  -F enforce_admins=false \
  -F required_pull_request_reviews[required_approving_review_count]=1 \
  -F restrictions= 2>/dev/null
```

Or set in the UI: https://github.com/knowledgestack/ks-xlsx-parser/settings/branches

## Per-release checklist

For every new version `X.Y.Z`:

1. **Decide the version number.** Follow [SemVer](https://semver.org/). Breaking API change → major bump. New feature, no breakage → minor. Bugfix only → patch.
2. **Bump version in two places** (kept in sync to avoid drift):
   - `pyproject.toml` — `version = "X.Y.Z"`
   - `src/ks_xlsx_parser/__init__.py` — `__version__ = "X.Y.Z"`
3. **Write the CHANGELOG entry** under a new `## [X.Y.Z] — YYYY-MM-DD` heading in [`CHANGELOG.md`](../CHANGELOG.md). Use the section labels documented at the top of that file (Added / Changed / Fixed / Performance / Docs / Internal / ⚠️ BREAKING).
4. **(Optional but recommended) Write hand-curated release notes** at `docs/launch/RELEASE_NOTES_vX.Y.Z.md`. If present, the release workflow picks it up automatically as the GitHub Release body; otherwise GitHub auto-generates from commits.
5. **Run `make test` locally** and verify all tests pass.
6. **Open a PR**, wait for CI green on all matrix cells, get a review, merge to `main`.
7. **Tag from `main`:**
   ```bash
   git checkout main && git pull
   git tag -a vX.Y.Z -m "vX.Y.Z — <short tagline>"
   git push origin vX.Y.Z
   ```
8. **Watch the workflow.** https://github.com/knowledgestack/ks-xlsx-parser/actions — the `Release` workflow should run `build` → `github-release` → `pypi`. If the `pypi` job is gated on a reviewer, approve it in the Actions UI.
9. **Verify post-release:**
   - PyPI: https://pypi.org/project/ks-xlsx-parser/X.Y.Z/ resolves and `pip install ks-xlsx-parser==X.Y.Z` works in a fresh venv.
   - GitHub Release: https://github.com/knowledgestack/ks-xlsx-parser/releases/tag/vX.Y.Z shows the release notes + wheel + sdist.
   - The `[Unreleased]` heading at the top of `CHANGELOG.md` is reset to "Nothing yet" for the next cycle (manual; do this in a follow-up PR).

## Common failure modes

**`pypi` job fails with "trusted publisher" error.** The PyPI Trusted Publisher binding is missing or pointing at the wrong workflow / environment. Fix at pypi.org (step 2 above). The GitHub Release will already exist — once the binding is correct, re-run the failed job via `gh run rerun --failed`.

**`pypi` job fails with "version already exists".** Someone published this version manually or a previous CI run partially succeeded. PyPI does not allow republishing. Bump to the next patch (`X.Y.Z+1`), update CHANGELOG / `RELEASE_NOTES`, delete the local tag, re-tag, force-delete the remote tag if it exists (`git push --delete origin vX.Y.Z`), and re-push.

**CI fails on a matrix cell after a tag is pushed.** The tag-trigger workflow doesn't depend on CI — it runs in parallel. If CI is red, delete the tag (`git push --delete origin vX.Y.Z`), fix the issue on a PR, then re-tag from the merged fix. The GitHub Release artifact from the doomed run can be deleted via `gh release delete vX.Y.Z`.

**Wheel doesn't include something it should.** Check `MANIFEST.in` / `pyproject.toml`'s `[tool.setuptools.packages.find]`. The `py.typed` marker, `*.pyi` stubs, and any data files (`.json`, `.xlsx` fixtures) only ship if explicitly declared.

## Hotfix / yank

If a published release has a critical bug:

```bash
# Yank from PyPI (hides from `pip install` but doesn't delete — required for cache-poisoning safety)
# UI: https://pypi.org/manage/project/ks-xlsx-parser/release/X.Y.Z/

# Tag a hotfix
git checkout main
# … apply fix, bump to X.Y.Z+1, update CHANGELOG …
git tag -a vX.Y.Z+1 -m "vX.Y.Z+1 — hotfix for <issue>"
git push origin vX.Y.Z+1
```

Yanked versions remain installable if pinned explicitly; `pip install ks-xlsx-parser` without a version constraint skips them. This is the safe default for accidental release of broken code.

## Why we use Trusted Publishing instead of an API token

PyPI API tokens stored as repo secrets are forge-able by anyone with write access to the workflow file. Trusted Publishing uses GitHub OIDC: only a real run of the configured workflow on the configured environment can mint a publish token, and the token is short-lived. No long-lived secret lives in the repo. Trade-off: one-time PyPI-side setup (step 2) and the environment binding (step 1) are both required.
