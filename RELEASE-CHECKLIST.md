# Release Checklist

## Source integrity

- [ ] Confirm the application version in `VERSION`.
- [ ] Confirm the release workbook matches the intended version.
- [ ] Re-export the complete VBA project.
- [ ] Confirm each `.frm` has its matching `.frx`.
- [ ] Confirm workbook and worksheet modules are included.
- [ ] Regenerate `SOURCE-MANIFEST.md`.

## Privacy and security

- [ ] Use a clean release workbook or explicitly approve the included sample data.
- [ ] Check the workbook for customer, supplier, employee, or financial records.
- [ ] Check source code for credentials, secrets, personal contact details, and hard-coded test accounts.
- [ ] Remove temporary files and private backups.
- [ ] Confirm repository visibility is intentional.

## Quality

- [ ] Complete the relevant checks in `TESTING.md`.
- [ ] Confirm installation instructions match the release.
- [ ] Update `README.md` and `CHANGELOG.md`.
- [ ] Check the release ZIP can be extracted and opened on a test machine.

## Publish

- [ ] Commit the verified source snapshot.
- [ ] Create a tag matching the application version, for example `v1.1.0`.
- [ ] Attach the clean workbook or full release package.
- [ ] Review release notes and attached files before publishing.
