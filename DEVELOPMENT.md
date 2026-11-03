# Development Guide

## Source layout

- `src/modules/` — 39 standard VBA modules (`.bas`).
- `src/forms/` — 25 UserForms (`.frm`) and 25 matching binary resources (`.frx`).
- `src/classes/` — 2 class modules (`.cls`).
- `src/workbook/` — 19 worksheet/workbook code exports (`.cls`).
- `IMS-v1.1.0.xlsm` — complete macro-enabled workbook.

## Safe development workflow

1. Preserve the original release workbook.
2. Make changes in a separate development copy.
3. Test changes with synthetic data in the development copy.
4. Export all affected VBA components.
5. Keep UserForm `.frm` and `.frx` files together.
6. Update `SOURCE-MANIFEST.md` after each export.
7. Update documentation and `CHANGELOG.md` when changes are verified.
8. Test the full workbook in desktop Excel before publishing a release.

## GitHub web upload note

Uploading files through GitHub's browser interface does not apply `.gitattributes` line-ending conversions. The attributes file is still useful for collaborators who use Git clients that honor it. Review the official GitHub documentation before choosing an upload method.

## Do not

- Replace the only copy of the workbook.
- mix files from different application versions in one source tree.
- delete `.frx` resources that belong to a UserForm.
- commit production data, credentials, or Office temporary files.
