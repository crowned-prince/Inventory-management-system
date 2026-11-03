# GitHub Website Upload Guide — AUVQEN IMS v1.1.0

This guide is for the GitHub website only. It does not require GitHub Desktop or command-line Git.

## Before starting

1. Extract `AUVQEN-IMS-GitHub-Ready-v1.1.0.zip`.
2. Open the extracted `AUVQEN-IMS` folder.
3. In File Explorer, enable **View → Show → Hidden items** so `.gitignore`, `.gitattributes`, `.editorconfig`, and `.github` are visible.
4. Keep the original workbook unchanged. Review the workbook and sample-data modules for private records before public distribution.
5. Use a private repository while reviewing proprietary code and data.

## Browser upload limits

GitHub's browser uploader accepts up to 100 files per upload operation and 25 MiB per file. The supplied workbook is approximately 1.35 MiB, so it is below the browser's per-file limit. Do not upload the entire `src` directory in one operation: its four child directories contain 110 source files altogether.

## Upload sequence

### 1. Root metadata (12 files)

From the extracted `AUVQEN-IMS` folder, select and upload these files to the repository root:

- `.editorconfig`
- `.gitattributes`
- `.gitignore`
- `CHANGELOG.md`
- `CONTRIBUTING.md`
- `LICENSE`
- `README.md`
- `SECURITY.md`
- `SOURCE-MANIFEST.json`
- `SOURCE-MANIFEST.md`
- `UPLOAD-PLAN.md`
- `VERSION`

The count is **12 files**. Use commit message: `Add repository metadata and project overview`.

### 2. Documentation folder (9 files)

Upload the contents of local `docs/` to a GitHub directory named `docs/`. If it does not exist, create it first with **Add file → Create new file**, filename `docs/.gitkeep`, add a short line, and commit. Then open `docs/` in GitHub and upload its local contents.

Contents: eight Markdown documents and `AUVQEN-IMS-User-Guide-v1.1.0.docx`. Commit message: `Add project documentation and user guide`.

### 3. Assets folder (3 files)

Upload local `assets/` to a GitHub directory named `assets/`. It contains `README.md`, `logo/.gitkeep`, and `screenshots/.gitkeep`. Create `assets/` first if needed. Commit message: `Add project asset placeholders`.

### 4. Source tree

Create the `src/` directory using **Add file → Create new file** and filename `src/.gitkeep`, then commit. Open `src/`. Upload each local child directory separately, one upload operation at a time:

| Local folder | Files | Commit message |
|---|---:|---|
| `src/classes/` | 2 | `Add VBA class modules` |
| `src/forms/` | 50 | `Add VBA UserForms and paired resources` |
| `src/modules/` | 39 | `Add standard VBA modules` |
| `src/workbook/` | 19 | `Add worksheet and workbook code exports` |

Use GitHub's **Add file → Upload files** page from inside `src/`. Drag one child folder at a time from File Explorer. GitHub's website supports dragging a folder into the upload area. Confirm the uploaded paths show `src/classes/`, `src/forms/`, `src/modules/`, or `src/workbook/` before committing.

If folder dragging is unavailable in your browser, create each destination with **Add file → Create new file** using a placeholder path such as `src/classes/.gitkeep`; open that destination folder, then select and upload the files inside the matching local folder. Keep every `.frm` together with its identically named `.frx`.

### 5. GitHub templates (3 files)

Upload local `.github/` to the repository root. It contains two issue templates under `ISSUE_TEMPLATE/` and one pull-request template. Create the `.github/` destination first if needed. Commit message: `Add GitHub issue and pull request templates`.

### 6. Main workbook (1 file)

Upload `IMS-v1.1.0.xlsm` to the repository root. Commit message: `Add AUVQEN IMS v1.1.0 workbook`.

## Final verification

Confirm that the repository root contains the README, configuration files, manifest, documentation, assets, source tree, GitHub templates, and workbook. Confirm that all 25 `.frm` files have their matching `.frx` files. Open `SOURCE-MANIFEST.md` and use it as the file inventory.

## GitHub web-interface note

GitHub documents that browser uploads support up to 100 files per operation and 25 MiB per file. It also notes that uploads through the web interface do not apply `.gitattributes` line-ending conversions. The `.gitattributes` file is still useful to collaborators who later use Git clients that honor it.
