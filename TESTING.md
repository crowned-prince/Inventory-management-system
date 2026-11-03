# Testing Guide

The source export includes files named `modTestPlan.bas`, `modTestData.bas`, and `modSampleData.bas`. Their presence is not proof that automated or manual tests passed. Record actual results when tests are run.

## Test environment record

Record:

- IMS version
- Windows version
- Microsoft Excel version and bitness, if known
- Test date
- Test workbook copy (never include production data)
- Tester and test result

## Suggested verification checklist

- [ ] Workbook opens in the intended desktop Excel version.
- [ ] First-run/setup behaviour is verified on a clean test copy.
- [ ] Login, logout, and password-change flows are checked.
- [ ] Product and category create/edit flows are checked.
- [ ] Stock adjustments and low-stock behaviour are checked.
- [ ] Purchase and supplier workflows are checked.
- [ ] Sales and customer workflows are checked.
- [ ] Expense, profit, chart, and report outputs are checked against known test data.
- [ ] User roles and access restrictions are checked.
- [ ] Audit-log behaviour is checked.
- [ ] Backup creation, verification, and restore are tested using disposable files.
- [ ] Invalid input and error paths are checked.
- [ ] Workbook remains usable after closing and reopening.

## Recording results

For every check, record the expected result, actual result, pass/fail status, and any issue reference. Use only synthetic data in screenshots and test files intended for GitHub.
