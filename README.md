# AUVQEN Inventory Management System (IMS)

**Version:** 1.1.0  
**Application format:** Microsoft Excel macro-enabled workbook (`.xlsm`)  
**Implementation:** Excel VBA

AUVQEN IMS is an Excel/VBA desktop inventory-management project. This repository contains the macro-enabled workbook, exported VBA components, and project documentation for the v1.1.0 source snapshot.

> **Repository status:** This is a source snapshot and documentation baseline. The presence of a module or form in the export does not, by itself, certify that every feature has been independently tested or is production-ready.

## Project contents

| Location | Contents |
|---|---|
| `IMS-v1.1.0.xlsm` | Main macro-enabled Excel workbook |
| `src/modules/` | 39 exported standard VBA modules (`.bas`) |
| `src/forms/` | 25 UserForms (`.frm`) and their 25 matching resource files (`.frx`) |
| `src/classes/` | 2 exported VBA class modules (`.cls`) |
| `src/workbook/` | 19 exported worksheet/workbook code modules (`.cls`) |
| `assets/` | Locations reserved for approved logo and screenshot assets |
| `docs/` | Installation, user, architecture, development, testing, and release documentation |
| `.github/` | Issue and pull-request templates |

## Source inventory

The exported source contains components whose names refer to areas such as inventory, products, sales, purchases, suppliers, customers, expenses, reporting, users, security, backups, dashboards, and setup. Consult `SOURCE-MANIFEST.md` and `docs/COMPONENT-CATALOG.md` for the exact exported file list. These names are a source inventory, not a guarantee that each workflow has passed testing.

## Requirements

- Windows environment compatible with the workbook
- Desktop Microsoft Excel with VBA support
- A trusted, clean copy of the workbook

Compatibility should be checked using the actual Excel version and Windows environment intended for deployment. Excel for the web is not a substitute for desktop Excel VBA execution.

## Getting started

1. Download the application package from the appropriate private release or approved distribution location.
2. Extract the package completely.
3. Keep a backup of any existing workbook and business data.
4. Open `IMS-v1.1.0.xlsm` in desktop Microsoft Excel.
5. Review the workbook's trust and macro prompts. Enable macros only when you have verified that the file is from a trusted source and you understand the risk.
6. Follow the application's own first-run/setup screens. Their exact wording and sequence should be verified against the shipped workbook.

See [Installation](docs/INSTALLATION.md) and the [User Guide](docs/USER-GUIDE.md).

## Development and source exports

VBA source is exported into `src/` so it can be reviewed and versioned separately from the binary workbook. Keep each `.frm` file with its matching `.frx` resource file. Keep the workbook as the runnable application; exported source files alone do not recreate the complete workbook, its sheets, formatting, formulas, names, or embedded data.

See [Development Guide](docs/DEVELOPMENT.md), [Architecture](docs/ARCHITECTURE.md), and [Testing](docs/TESTING.md).

## Versioning

The repository version is recorded in [`VERSION`](VERSION). Release notes are recorded in [`CHANGELOG.md`](CHANGELOG.md). The exporter version is separate from the IMS application version.

## Data and security

- Do not commit customer records, real supplier/customer contact details, passwords, recovery codes, or production backups.
- Review both the workbook and sample/test-data modules before publishing this repository publicly.
- Keep the repository private while reviewing proprietary source code and data.
- Report suspected vulnerabilities privately; see [SECURITY.md](SECURITY.md).

## License

See [`LICENSE`](LICENSE). The included notice is a proprietary-license starting point and should be reviewed before commercial distribution.

---

**Developed under AUVQEN**
