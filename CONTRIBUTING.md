# Contributing to AUVQEN IMS

Thank you for helping improve AUVQEN IMS. Access to this repository does not automatically grant permission to redistribute the application or its source code.

## Before making changes

1. Confirm that you have permission to contribute.
2. Work on a separate copy of the workbook.
3. Keep a recoverable backup before changing workbook structure or data.
4. Read the development and architecture documents.

## VBA source rules

- Keep each `.frm` file together with its matching `.frx` file.
- Keep exported modules in their existing directories.
- Do not rename exported components without updating the workbook and source manifest.
- Do not commit temporary Office lock files, private workbooks, customer data, credentials, or production backups.
- Update documentation when a real workflow or component changes.

## Testing before proposing a change

- Test the affected workflow in a disposable test copy.
- Test relevant neighboring workflows, especially those that share inventory or user data.
- Record the Excel version and environment used.
- Record expected and actual results without including private business data.
- Do not describe a change as verified unless the test was actually performed.

## Pull requests

Include a clear summary, reason for the change, affected files, test steps, test results, and any required data or workbook migration notes.
