# User Guide — Verification Draft

This guide is a structured starting point based on the exported workbook and VBA component names. The actual screen labels, buttons, validation rules, and workflow order must be verified in the running v1.1.0 workbook before this document is treated as a final customer manual.

## 1. Opening the application

1. Open the approved `.xlsm` release in desktop Microsoft Excel.
2. Review Office security prompts and enable macros only when the workbook is trusted.
3. Follow the first-run/setup or login screen shown by the workbook.
4. Do not use real business records during initial testing; use a disposable test copy.

The export contains forms named `frmWelcome`, `frmFirstTimeSetup`, and `frmLogin`. Their presence indicates the exported UI components, but their exact sequence should be confirmed in the workbook.

## 2. Main areas represented in the source export

The exported forms and modules include names relating to:

- Dashboard and charts
- Products, stock, and categories
- Sales and purchases
- Suppliers and customers
- Expenses and profit reporting
- User administration and passwords
- Audit log, backup, settings, and application health

See [Component Catalog](COMPONENT-CATALOG.md) for the exact source inventory. Confirm which screens are enabled and their actual business rules in the application.

## 3. Daily workflow — complete after validation

The project owner should document the verified steps for:

1. Signing in and signing out.
2. Creating and editing products.
3. Receiving stock and recording purchases.
4. Recording sales and returns, if supported.
5. Managing suppliers and customers.
6. Reviewing low-stock information.
7. Generating and exporting reports.
8. Managing users and permissions.
9. Creating and verifying backups.

Do not infer button names or business rules from form filenames. Add screenshots only after checking that they contain no real customer or financial information.

## 4. Data safety

- Use the application's backup process where available and independently retain a safe copy before major changes.
- Test restore operations using a disposable copy first.
- Avoid editing internal worksheet structures unless you understand the application's dependencies.
- Keep production files and credentials private.

## 5. Troubleshooting

Record the application version, Excel version, Windows version, the screen involved, and a sanitized description of the problem. Never attach a workbook containing private business records to a public issue.
