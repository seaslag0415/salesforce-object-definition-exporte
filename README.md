# getObjectDefine

Excel VBA tool for exporting Salesforce object metadata and permissions.

## Overview

This workbook connects to Salesforce through Salesforce CLI and exports the selected object's definition into Excel sheets.

The macro exports:

- Object basic information
- Field definitions
- Object access permissions by profile and permission set
- Field access permissions by profile and permission set
- Error details to an error log sheet

## Requirements

- Microsoft Excel with macros enabled
- Salesforce CLI (`sf`) installed and available from the command line
- A Salesforce user that can read object and permission metadata

## Workbook

The macro-enabled workbook is stored in:

```text
workbook/オブジェクト定義取得.xlsm
```

## Usage

1. Open `workbook/オブジェクト定義取得.xlsm`.
2. Enable macros.
3. On the `取得対象` sheet, select the connection target.
4. Enter the Salesforce object API name in `C5`.
5. Click `接続`.
6. Complete browser authentication when Salesforce CLI opens the browser.

Generated sheets are recreated on each run. The `取得対象` and `エラー` sheets are preserved.

## Source Layout

```text
src/vba/
  modObjectDefineMain.bas
  modJsonParser.bas
  sheet_target.cls

tools/
  install_object_define_macro.ps1
```

## Notes

- Authentication is delegated to Salesforce CLI using `sf org login web`.
- Object and field definitions are retrieved with `sf sobject describe`.
- Object and field permissions are retrieved with SOQL through `sf data query`.
- Permission set display names are supplemented by querying `PermissionSet` using the permission parent IDs.
