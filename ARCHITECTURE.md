# Architecture Overview

This description is based on the folders and filenames in the supplied export. It is a source-layout overview, not a verified runtime dependency diagram.

## Main workbook

`IMS-v1.1.0.xlsm` is the runnable Excel macro-enabled workbook. It contains workbook sheets, formulas, named items, embedded resources, and VBA project data that are not fully represented by exporting VBA source files alone.

## VBA source groups

### Standard modules

`src/modules/` contains 39 `.bas` files. Their filenames indicate areas including application startup, data access, products, inventory, sales, purchases, suppliers, customers, reporting, users, security, validation, backups, dashboards, setup, and test/sample data.

### UserForms

`src/forms/` contains 25 `.frm` files and their matching 25 `.frx` files. The `.frm` files contain form definitions and code; `.frx` files hold associated binary form resources. Keep each pair together.

### Class modules

`src/classes/` contains two exported `.cls` class modules.

### Workbook and worksheet code

`src/workbook/` contains 19 exported `.cls` files for worksheet and workbook code modules. The workbook itself remains necessary to restore the complete project context.

## Version control

Text-based source files can be reviewed as text. The `.xlsm` workbook and `.frx` files are binary and should be treated as indivisible files. GitHub's web uploader does not apply `.gitattributes` line-ending normalization.
