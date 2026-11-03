# GitHub Browser Upload Plan — AUVQEN IMS v1.1.0

This package is arranged so each folder can be uploaded separately through GitHub's browser interface. GitHub currently allows up to 100 files per browser upload and a maximum of 25 MiB per file.

## Upload order

1. Root metadata files: `README.md`, `.gitignore`, `.gitattributes`, `.editorconfig`, `VERSION`, `CHANGELOG.md`, `LICENSE`, `CONTRIBUTING.md`, `SECURITY.md`, `SOURCE-MANIFEST.md`, and `UPLOAD-PLAN.md`.
2. `docs/` (contains the Markdown guides and the DOCX user-guide draft).
3. `assets/` (currently contains README and `.gitkeep` placeholders only).
4. `src/classes/` (2 files).
5. `src/forms/` (50 files: 25 `.frm` and 25 matching `.frx`).
6. `src/modules/` (39 `.bas` files).
7. `src/workbook/` (19 `.cls` files).
8. `.github/` (3 issue/pull-request template files).
9. `IMS-v1.1.0.xlsm` (1 file, approximately 1.35 MiB in the supplied archive).

Each listed folder contains fewer than 100 files. Do not upload the entire `src/` directory in one operation: it contains 110 source files and would exceed the browser's 100-file limit.

## Browser workflow for each batch

- Open the repository's target directory on GitHub.
- Select **Add file → Upload files**.
- Drag the contents of the matching local directory into the upload area (not the enclosing parent folder if that would create an extra nesting level).
- Check the displayed filenames and wait for completion.
- Enter the commit message suggested in the step-by-step guide and select **Commit changes**.
- Repeat for the next directory.

GitHub's web upload ignores `.gitattributes` line-ending conversion. Keep the file in the repository for collaborators using Git clients that honor it.

## Suggested commit messages

- `Add repository metadata and project overview`
- `Add project documentation and user guide`
- `Add project asset placeholders`
- `Add VBA class modules`
- `Add VBA UserForms and paired resources`
- `Add standard VBA modules`
- `Add worksheet and workbook code exports`
- `Add GitHub issue and pull request templates`
- `Add AUVQEN IMS v1.1.0 workbook`
