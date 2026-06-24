# getObjectDefine

Excel VBA tool for exporting Salesforce object metadata and permissions.

Salesforce CLI is used for authentication and data retrieval, so the workbook does not store Salesforce credentials.

## Features

- Export object basic information
- Export field definitions
- Export object access permissions by profile and permission set
- Export field access permissions by profile and permission set
- Log errors to the `エラー` sheet
- Recreate output sheets on each run while preserving `取得対象` and `エラー`

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
- Permission set display names are supplemented by querying `PermissionSet` using permission parent IDs.

## Disclaimer

This tool is provided as-is, without warranty of any kind.
Use it at your own risk.

The author is not responsible for any damage, data loss, security issue, Salesforce configuration change, or business impact caused by using this tool.
Please review the VBA code and test it in a safe environment before use.

本ツールは現状有姿で提供されます。
利用は自己責任で行ってください。

本ツールの利用により発生した損害、データ損失、セキュリティ上の問題、Salesforce設定への影響、業務上の影響について、作者は責任を負いません。
利用前にVBAコードを確認し、安全な環境で検証してください。

## License

MIT License. See [LICENSE](LICENSE).
