# Releasing `docx-scalpel` to PyPI

`docx-scalpel` is published via **PyPI Trusted Publishing (OIDC)** — no API tokens stored in GitHub secrets. The publish job in [`.github/workflows/python-publish.yml`](../.github/workflows/python-publish.yml) exchanges an OIDC token signed by GitHub for a short-lived PyPI upload credential at publish time.

## One-time setup (repo-level)

Two things must exist before the first publish:

### 1. GitHub Actions environments

Create two environments in **Settings → Environments**:

| Name | Purpose | Suggested protection |
|---|---|---|
| `testpypi` | Dry-run publishes via `workflow_dispatch` | None — used freely |
| `pypi` | Production publishes (auto on `docx-scalpel-v*` tag push) | Required reviewer, restrict to `main` branch |

The names are case-sensitive and must match the `environment.name` fields in the workflow.

### 2. PyPI / TestPyPI trusted publishers

For **each** of TestPyPI and PyPI, register a pending trusted publisher (the `docx-scalpel` project doesn't exist yet on first upload).

- TestPyPI: https://test.pypi.org/manage/account/publishing/
- PyPI: https://pypi.org/manage/account/publishing/

Fill out the form exactly:

| Field | Value |
|---|---|
| PyPI project name | `docx-scalpel` |
| Owner | `JSv4` |
| Repository name | `Docxodus` |
| Workflow filename | `python-publish.yml` |
| Environment name | `testpypi` (or `pypi`, matching the row) |

After the first successful publish, the "pending" publisher becomes a regular trusted publisher on the project page.

## Routine release

**docx-scalpel ships on its own tag pattern, decoupled from Docxodus core / npm / binaries.**

| Tag pattern | Workflow | Publishes |
|---|---|---|
| `v*` (e.g. `v1.2.3`) | `publish.yml` | NuGet (Docxodus + redline + docx2html + docx2oc) + npm (`docxodus`) + GitHub Release binaries |
| `docx-scalpel-v*` (e.g. `docx-scalpel-v0.1.0a1`) | `python-publish.yml` | PyPI (`docx-scalpel`) |

The two tag namespaces are independent. A docx-scalpel point-release doesn't drag Docxodus core through a version bump, and a Docxodus core release doesn't force unrelated PyPI churn. Each docx-scalpel wheel bundles a `docxodus-pyhost` built from the same commit the tag points to, so there's no runtime version-pinning between the two distributions.

### Recommended sequence

1. **Dry-run on TestPyPI** (manual)
   - Actions → `python-publish` → Run workflow
   - Branch: whatever you're cutting from (default branch for stable, feature branch for testing the pipeline itself)
   - `version`: PEP 440 string, e.g. `0.1.0a1`
   - `target`: `testpypi`
   - Confirm the upload by installing in a fresh venv:
     ```bash
     pip install --index-url https://test.pypi.org/simple/ \
         --extra-index-url https://pypi.org/simple/ \
         docx-scalpel==0.1.0a1
     python -c "from docx_scalpel import ping; print(ping())"
     ```

2. **Real release on PyPI** (tag-driven)
   - Bump the version in `python/pyproject.toml`. CI overwrites this at build time, but keeping the in-tree value current is helpful for editable installs and PR readability.
     ```bash
     # python/pyproject.toml: version = "0.1.0a1"
     git commit -am "chore(python): bump docx-scalpel to 0.1.0a1"
     git tag docx-scalpel-v0.1.0a1
     git push origin main docx-scalpel-v0.1.0a1
     ```
   - The tag fires `python-publish.yml`. `resolve` extracts `0.1.0a1` from `${GITHUB_REF#refs/tags/docx-scalpel-v}` and resolves target = `pypi`.

### Pre-releases

PyPI accepts PEP 440 pre-release identifiers — `0.1.0a1`, `0.1.0b2`, `0.1.0rc1`, `0.1.0.dev3`. Use the tag form `docx-scalpel-v0.1.0a1` and pip will only install pre-releases when `--pre` is passed.

## Wheel scope (v1)

Today the workflow ships:

- `docx_scalpel-${VER}-py3-none-manylinux_2_28_x86_64.whl` — linux-x64, bundled `docxodus-pyhost`
- `docx_scalpel-${VER}.tar.gz` — sdist (no bundled binary; install requires `DOCXODUS_HOST` env var or a dev clone of Docxodus)

To install on linux-x64 the user does `pip install docx-scalpel` and gets the bundled host transparently. On other platforms today, the sdist install path requires building the host out-of-band — temporary until we add the other RIDs.

## Adding more RIDs

The matrix is the obvious one. Adding `linux-arm64`, `osx-x64`, `osx-arm64`, `win-x64` is a mechanical extension of `build-wheel-linux-x64`:

1. Duplicate the `build-wheel-linux-x64` job per RID, parameterized by:
   - `runs-on`: `ubuntu-22.04-arm` (linux-arm64), `macos-13` (osx-x64), `macos-14` (osx-arm64), `windows-latest` (win-x64)
   - `-r <rid>` for `dotnet publish`
   - Platform tag passed to `wheel tags`:
     - `manylinux_2_28_aarch64`
     - `macosx_11_0_x86_64`
     - `macosx_11_0_arm64`
     - `win_amd64`
2. Linux ARM may need a manylinux container to ensure glibc compliance — see `auditwheel` if it does.
3. Windows: the host binary is `docxodus-pyhost.exe`; everything else is identical. Use a bash shell for the wheel/sed steps (`shell: bash` or rewrite in PowerShell).
4. Add the new artifact name to the `publish` job's `pattern: dist-*` glob (already wildcard-covered).

Each per-RID wheel is independent — the publish job collects all artifacts from `dist-*` and uploads them as one PyPI release.

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| `OIDC token exchange failed` in publish step | Trusted publisher not registered, or the environment name in the workflow doesn't match what's configured on PyPI |
| `403 Forbidden` from PyPI | Same as above, or the project name was claimed by someone else after the pending-publisher form was filled out |
| `pyhost ping fails in CI` smoke test | `dotnet publish` succeeded but produced a binary missing a runtime dep; check the .NET SDK version in `setup-dotnet` matches the one in `global.json` |
| Wheel installs but `find_host()` raises `DocxodusHostNotFoundError` | Wheel was built without the binary staged in `src/docx_scalpel/_bin/`; the workflow's "Stage binary into wheel layout" step did not run or wrote to the wrong path |
| `auditwheel` complains about glibc | Build on a manylinux container (`quay.io/pypa/manylinux_2_28_x86_64`) instead of the `ubuntu-latest` runner |

## Why trusted publishing

- No long-lived API tokens to rotate or leak via a compromised GitHub Secret.
- Token is scoped to one workflow run on one repo on one environment — exfiltration window is minutes, not the lifetime of a secret.
- No password to share among maintainers; access is governed by who can push tags / approve environment deployments.

See [PyPI's Trusted Publishers docs](https://docs.pypi.org/trusted-publishers/) for the full rationale and security model.
