# GitHub Actions workflows

These two workflows belong in `.github/workflows/`. They are kept here instead
because the bridge that wrote this repository to disk refuses to write into
`.github/`, which is a sensible guard: a workflow file is code that GitHub runs
automatically on your account, and nothing automated should be able to install
one without you seeing it first.

Read them, then move them yourself:

```bash
mkdir -p .github/workflows
mv ci-workflows/build.yml ci-workflows/pages.yml .github/workflows/
rmdir ci-workflows 2>/dev/null || rm -rf ci-workflows
```

## What each one does

**`build.yml`** runs on every push and pull request. It installs `openpyxl`,
rebuilds the page from the workbook, and fails if the result differs from the
committed `index.html`. The date stamp is normalised before the comparison, so
only a real change trips it. It then runs `tools/check_page.py`, which fails if
the page has picked up an external asset reference, an em dash or an en dash, or
has lost the workbook it carries. Nothing is deployed and nothing is written back
to the repository; the job only reads and reports.

**`pages.yml`** publishes the repository to GitHub Pages on every push to `main`.
It needs Pages turned on under **Settings, Pages** with the source set to
**GitHub Actions**. If you would rather serve the page from the branch directly,
you do not need this file at all: set the Pages source to **Deploy from a
branch**, pick `main` and `/ (root)`, and delete it. `index.html` sits at the
repository root, so either route serves it with no build step.

Neither workflow needs a secret, and neither can write to the repository.
