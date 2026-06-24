Attribute VB_Name = "modObjectDefineMain"
Option Explicit

Private Const TARGET_SHEET_NAME As String = "取得対象"
Private Const PROD_URL As String = "https://login.salesforce.com"
Private Const SANDBOX_URL As String = "https://test.salesforce.com"
Private Const ERROR_SHEET_NAME As String = "エラー"

Private m_CurrentStep As String
Private m_LastCommand As String
Private m_LastStdout As String
Private m_LastStderr As String

Public Sub RunObjectDefinitionExport()
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(TARGET_SHEET_NAME)

    Dim loginKind As String
    Dim loginUrl As String
    Dim objectApiName As String

    loginKind = Trim$(CStr(ws.Range("C3").Value))
    loginUrl = Trim$(CStr(ws.Range("D3").Value))
    objectApiName = Trim$(CStr(ws.Range("C5").Value))

    If loginKind = "" Then Err.Raise vbObjectError + 100, , "接続先を選択してください。"
    If loginUrl = "" Then Err.Raise vbObjectError + 101, , "接続先URLを入力してください。"
    If objectApiName = "" Then Err.Raise vbObjectError + 102, , "オブジェクトAPI参照名を入力してください。"

    Application.ScreenUpdating = False
    EnsureErrorSheet False
    DeleteOutputSheets

    m_CurrentStep = "Salesforce CLI確認"
    Application.StatusBar = "Salesforce CLIを確認しています..."

    Dim outputText As String
    Dim errorText As String
    Dim exitCode As Long

    exitCode = RunCommandCapture("sf --version", outputText, errorText)
    If exitCode <> 0 Then
        Err.Raise vbObjectError + 103, , "Salesforce CLI の実行に失敗しました。" & vbCrLf & CommandErrorDetail(outputText, errorText)
    End If

    Dim aliasName As String
    aliasName = "ExcelObjectDefine_" & Format$(Now, "yyyymmdd_hhnnss")

    m_CurrentStep = "Salesforceログイン"
    Application.StatusBar = "ブラウザ認証を開始しています..."
    exitCode = RunCommandCapture("sf org login web --instance-url " & QuoteArg(loginUrl) & " --alias " & QuoteArg(aliasName), outputText, errorText)
    If exitCode <> 0 Then
        Err.Raise vbObjectError + 104, , "Salesforceログインに失敗しました。" & vbCrLf & CommandErrorDetail(outputText, errorText)
    End If

    m_CurrentStep = "オブジェクト定義取得"
    Application.StatusBar = "オブジェクト定義を取得しています..."
    Dim describeJson As String
    exitCode = RunCommandCapture("sf sobject describe --target-org " & QuoteArg(aliasName) & " --sobject " & QuoteArg(objectApiName) & " --json", describeJson, errorText)
    If exitCode <> 0 Then
        Err.Raise vbObjectError + 105, , "オブジェクト定義の取得に失敗しました。" & vbCrLf & CommandErrorDetail(describeJson, errorText)
    End If

    Dim describeRoot As Object
    Dim describeResult As Object
    Set describeRoot = ParseJson(describeJson)
    Set describeResult = GetDictObject(describeRoot, "result")

    m_CurrentStep = "オブジェクト権限取得"
    Application.StatusBar = "オブジェクト権限を取得しています..."
    Dim objectPermJson As String
    Dim soql As String
    soql = "SELECT ParentId, Parent.Name, Parent.Label, Parent.IsOwnedByProfile, Parent.Profile.Name, SObjectType, PermissionsRead, PermissionsCreate, PermissionsEdit, PermissionsDelete, PermissionsViewAllRecords, PermissionsModifyAllRecords FROM ObjectPermissions WHERE SObjectType = '" & EscapeSoql(objectApiName) & "'"
    exitCode = RunCommandCapture("sf data query --target-org " & QuoteArg(aliasName) & " --query=" & QuoteArg(soql) & " --json", objectPermJson, errorText)
    If exitCode <> 0 Then
        Err.Raise vbObjectError + 106, , "オブジェクト権限の取得に失敗しました。" & vbCrLf & CommandErrorDetail(objectPermJson, errorText)
    End If

    m_CurrentStep = "項目権限取得"
    Application.StatusBar = "項目権限を取得しています..."
    Dim fieldPermJson As String
    soql = "SELECT ParentId, Parent.Name, Parent.Label, Parent.IsOwnedByProfile, Parent.Profile.Name, SObjectType, Field, PermissionsRead, PermissionsEdit FROM FieldPermissions WHERE SObjectType = '" & EscapeSoql(objectApiName) & "'"
    exitCode = RunCommandCapture("sf data query --target-org " & QuoteArg(aliasName) & " --query=" & QuoteArg(soql) & " --json", fieldPermJson, errorText)
    If exitCode <> 0 Then
        Err.Raise vbObjectError + 107, , "項目権限の取得に失敗しました。" & vbCrLf & CommandErrorDetail(fieldPermJson, errorText)
    End If

    m_CurrentStep = "権限セット情報取得"
    Application.StatusBar = "権限セット情報を取得しています..."
    Dim objectPermRoot As Object
    Dim fieldPermRoot As Object
    Dim permissionSetJson As String
    Dim permissionSetRoot As Object
    Dim permissionSetMap As Object

    Set objectPermRoot = ParseJson(objectPermJson)
    Set fieldPermRoot = ParseJson(fieldPermJson)

    soql = BuildPermissionSetQuery(objectPermRoot, fieldPermRoot)
    If Len(soql) > 0 Then
        exitCode = RunCommandCapture("sf data query --target-org " & QuoteArg(aliasName) & " --query=" & QuoteArg(soql) & " --json", permissionSetJson, errorText)
        If exitCode <> 0 Then
            Err.Raise vbObjectError + 108, , "権限セット情報の取得に失敗しました。" & vbCrLf & CommandErrorDetail(permissionSetJson, errorText)
        End If
        Set permissionSetRoot = ParseJson(permissionSetJson)
        Set permissionSetMap = BuildPermissionSetMap(permissionSetRoot)
    Else
        Set permissionSetMap = CreateObject("Scripting.Dictionary")
    End If

    m_CurrentStep = "Excelシート出力"
    Application.StatusBar = "Excelシートへ出力しています..."
    WriteBasicInfo describeResult
    WriteFieldDefinitions describeResult
    WriteObjectPermissions objectPermRoot, permissionSetMap
    WriteFieldPermissions fieldPermRoot, describeResult, permissionSetMap

    ThisWorkbook.Worksheets("オブジェクト基本情報").Activate
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "オブジェクト定義の取得が完了しました。", vbInformation
    Exit Sub

EH:
    WriteErrorLog m_CurrentStep, Err.Description, m_LastCommand, m_LastStdout, m_LastStderr
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "処理中にエラーが発生しました。詳細は「" & ERROR_SHEET_NAME & "」シートを確認してください。", vbExclamation, "オブジェクト定義取得"
End Sub

Public Sub SetupObjectDefineWorkbook()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(TARGET_SHEET_NAME)

    ws.Range("A1").Value = "オブジェクト定義取得"
    ws.Range("B3").Value = "接続先"
    ws.Range("B5").Value = "オブジェクトAPI参照名"

    With ws.Range("C3")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="本番,Sandbox,カスタムURL"
        .Validation.IgnoreBlank = True
        .Validation.InCellDropdown = True
        If Len(Trim$(CStr(.Value))) = 0 Then .Value = "本番"
    End With

    UpdateLoginUrlFromSelection
    ws.Range("C3:D3").Borders.LineStyle = xlContinuous
    ws.Range("C5").Borders.LineStyle = xlContinuous
    ws.Columns("A:E").AutoFit
    EnsureErrorSheet True
End Sub

Public Sub UpdateLoginUrlFromSelection()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(TARGET_SHEET_NAME)

    Select Case Trim$(CStr(ws.Range("C3").Value))
        Case "本番"
            ws.Range("D3").Value = PROD_URL
        Case "Sandbox"
            ws.Range("D3").Value = SANDBOX_URL
        Case "カスタムURL"
            ' ユーザー入力を保持するため何もしない
    End Select
End Sub

Private Function RunCommandCapture(ByVal commandLine As String, ByRef stdoutText As String, ByRef stderrText As String) As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    m_LastCommand = commandLine
    m_LastStdout = ""
    m_LastStderr = ""

    Dim tempFolder As String
    tempFolder = Environ$("TEMP")

    Dim token As String
    token = Format$(Now, "yyyymmddhhnnss") & "_" & CStr(Int(Rnd() * 100000))

    Dim outPath As String
    Dim errPath As String
    Dim cmdPath As String
    outPath = tempFolder & "\sf_stdout_" & token & ".txt"
    errPath = tempFolder & "\sf_stderr_" & token & ".txt"
    cmdPath = tempFolder & "\sf_run_" & token & ".cmd"

    WriteTextFileDefault cmdPath, "@echo off" & vbCrLf & commandLine & " > " & QuoteArg(outPath) & " 2> " & QuoteArg(errPath) & vbCrLf & "exit /b %ERRORLEVEL%" & vbCrLf

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    RunCommandCapture = sh.Run("cmd.exe /C " & QuoteArg(cmdPath), 0, True)

    stdoutText = ReadTextFileIfExists(outPath)
    stderrText = ReadTextFileIfExists(errPath)
    m_LastStdout = stdoutText
    m_LastStderr = stderrText

    On Error Resume Next
    fso.DeleteFile outPath, True
    fso.DeleteFile errPath, True
    fso.DeleteFile cmdPath, True
    On Error GoTo 0
End Function

Private Sub WriteTextFileDefault(ByVal filePath As String, ByVal text As String)
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(filePath, True, False)
    ts.Write text
    ts.Close
End Sub

Private Function ReadTextFileIfExists(ByVal filePath As String) As String
    If Len(Dir$(filePath)) = 0 Then
        ReadTextFileIfExists = ""
        Exit Function
    End If

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    ReadTextFileIfExists = stm.ReadText
    stm.Close
End Function

Private Function CommandErrorDetail(ByVal stdoutText As String, ByVal stderrText As String) As String
    Dim detail As String
    detail = ""

    If Len(Trim$(stderrText)) > 0 Then
        detail = detail & stderrText
    End If

    If Len(Trim$(stdoutText)) > 0 Then
        If Len(detail) > 0 Then detail = detail & vbCrLf
        detail = detail & stdoutText
    End If

    If Len(Trim$(detail)) = 0 Then
        detail = "Salesforce CLIから詳細エラーが返されませんでした。"
    End If

    CommandErrorDetail = detail
End Function

Private Sub EnsureErrorSheet(ByVal resetSheet As Boolean)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ERROR_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = ERROR_SHEET_NAME
    ElseIf resetSheet Then
        ws.Cells.Clear
    End If

    If Len(CStr(ws.Range("A1").Value)) = 0 Or resetSheet Then
        ws.Cells.Clear
        ws.Range("A1:F1").Merge
        ws.Range("A1").Value = "エラーログ"
        ws.Range("A2:F2").Merge
        ws.Range("A2").Value = "マクロ実行時のエラー詳細を出力します。"
        ws.Range("A4:F4").Value = Array("日時", "処理", "エラー内容", "実行コマンド", "標準出力", "標準エラー")
        FormatOutputSheet ws, "F"
        ws.Columns("A:A").ColumnWidth = 20
        ws.Columns("B:B").ColumnWidth = 22
        ws.Columns("C:F").ColumnWidth = 60
    End If
End Sub

Private Sub WriteErrorLog(ByVal stepName As String, ByVal errorMessage As String, ByVal commandLine As String, ByVal stdoutText As String, ByVal stderrText As String)
    On Error Resume Next

    EnsureErrorSheet False

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(ERROR_SHEET_NAME)

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 5 Then nextRow = 5

    ws.Cells(nextRow, 1).Value = Format$(Now, "yyyy/mm/dd hh:nn:ss")
    ws.Cells(nextRow, 2).Value = stepName
    ws.Cells(nextRow, 3).Value = errorMessage
    ws.Cells(nextRow, 4).Value = commandLine
    ws.Cells(nextRow, 5).Value = stdoutText
    ws.Cells(nextRow, 6).Value = stderrText

    ws.Rows(nextRow).WrapText = True
    ws.Rows(nextRow).VerticalAlignment = xlTop
    ws.Range(ws.Cells(nextRow, 1), ws.Cells(nextRow, 6)).Borders.LineStyle = xlContinuous
    ThisWorkbook.Save
    ws.Activate
End Sub

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = """" & Replace(value, """", """""") & """"
End Function

Private Function EscapeSoql(ByVal value As String) As String
    EscapeSoql = Replace(value, "'", "\'")
End Function

Private Sub WriteBasicInfo(ByVal describeResult As Object)
    Dim ws As Worksheet
    Set ws = PrepareSheet("オブジェクト基本情報")

    ws.Range("A1:E1").Merge
    ws.Range("A1").Value = "オブジェクト基本情報"
    ws.Range("A2:E2").Merge
    ws.Range("A2").Value = "取得対象オブジェクトの describe 結果から、主要な基本情報を出力するシート"
    ws.Range("A4:E4").Value = Array("区分", "項目", "値", "説明", "取得元")

    Dim rows As Collection
    Set rows = New Collection
    AddInfoRow rows, "識別情報", "オブジェクトAPI参照名", DictValue(describeResult, "name"), "Salesforce上のオブジェクトAPI名", "describe.name"
    AddInfoRow rows, "識別情報", "表示ラベル", DictValue(describeResult, "label"), "画面上の単数形ラベル", "describe.label"
    AddInfoRow rows, "識別情報", "表示ラベル（複数）", DictValue(describeResult, "labelPlural"), "画面上の複数形ラベル", "describe.labelPlural"
    AddInfoRow rows, "識別情報", "キー接頭辞", DictValue(describeResult, "keyPrefix"), "レコードIDの先頭3文字", "describe.keyPrefix"
    AddInfoRow rows, "種別", "カスタムオブジェクト", BoolText(DictValue(describeResult, "custom")), "カスタムオブジェクトの場合 TRUE", "describe.custom"
    AddInfoRow rows, "種別", "カスタム設定", BoolText(DictValue(describeResult, "customSetting")), "カスタム設定の場合 TRUE", "describe.customSetting"
    AddInfoRow rows, "利用可否", "作成可能", BoolText(DictValue(describeResult, "createable")), "ログインユーザがこのオブジェクトを作成できるか", "describe.createable"
    AddInfoRow rows, "利用可否", "参照可能", BoolText(DictValue(describeResult, "queryable")), "ログインユーザがこのオブジェクトを参照できるか", "describe.queryable"
    AddInfoRow rows, "利用可否", "更新可能", BoolText(DictValue(describeResult, "updateable")), "ログインユーザがこのオブジェクトを更新できるか", "describe.updateable"
    AddInfoRow rows, "利用可否", "削除可能", BoolText(DictValue(describeResult, "deletable")), "ログインユーザがこのオブジェクトを削除できるか", "describe.deletable"
    AddInfoRow rows, "利用可否", "検索可能", BoolText(DictValue(describeResult, "searchable")), "SOSL検索対象にできるか", "describe.searchable"
    AddInfoRow rows, "機能", "レコードタイプ対応", IIf(CollectionCount(DictValue(describeResult, "recordTypeInfos")) > 0, "TRUE", "FALSE"), "レコードタイプを持つか", "describe.recordTypeInfos"
    AddInfoRow rows, "機能", "活動許可", BoolText(DictValue(describeResult, "allowActivities")), "活動に関連付け可能か", "describe.allowActivities"
    AddInfoRow rows, "機能", "フィード有効", BoolText(DictValue(describeResult, "feedEnabled")), "Chatterフィードが有効か", "describe.feedEnabled"
    AddInfoRow rows, "補足", "取得日時", Format$(Now, "yyyy/mm/dd hh:nn:ss"), "マクロ実行日時", "Excelマクロ"

    WriteCollectionRows ws, rows, 5, 5
    FormatOutputSheet ws, "E"
End Sub

Private Sub WriteFieldDefinitions(ByVal describeResult As Object)
    Dim ws As Worksheet
    Set ws = PrepareSheet("項目定義")

    ws.Range("A1:S1").Merge
    ws.Range("A1").Value = "項目定義"
    ws.Range("A2:S2").Merge
    ws.Range("A2").Value = "項目そのものの定義情報を1項目1行で出力するシート。プロファイル/権限セットごとの項目アクセスは「項目アクセス権限」シートに出力する。"
    ws.Range("A4:S4").Value = Array("No", "オブジェクトAPI参照名", "項目API参照名", "表示ラベル", "データ型", "桁数", "小数桁", "必須", "一意", "外部ID", "自動採番", "作成可", "更新可", "参照先", "選択リスト値", "数式", "ヘルプテキスト", "説明", "取得元")

    Dim objectName As String
    objectName = CStr(DictValue(describeResult, "name"))

    Dim fields As Collection
    Set fields = DictValue(describeResult, "fields")

    Dim rows As Collection
    Set rows = New Collection

    Dim i As Long
    Dim f As Object
    For i = 1 To fields.Count
        Set f = fields(i)
        Dim row(1 To 19) As Variant
        row(1) = i
        row(2) = objectName
        row(3) = DictValue(f, "name")
        row(4) = DictValue(f, "label")
        row(5) = DictValue(f, "type")
        row(6) = DictValue(f, "length")
        row(7) = DictValue(f, "scale")
        row(8) = BoolText(Not CBoolDefault(DictValue(f, "nillable"), True))
        row(9) = BoolText(DictValue(f, "unique"))
        row(10) = BoolText(DictValue(f, "externalId"))
        row(11) = BoolText(DictValue(f, "autoNumber"))
        row(12) = BoolText(DictValue(f, "createable"))
        row(13) = BoolText(DictValue(f, "updateable"))
        row(14) = JoinCollection(DictValue(f, "referenceTo"), "; ")
        row(15) = JoinPicklistValues(DictValue(f, "picklistValues"))
        row(16) = DictValue(f, "calculatedFormula")
        row(17) = DictValue(f, "inlineHelpText")
        row(18) = ""
        row(19) = "fields[].*"
        rows.Add row
    Next

    WriteCollectionRows ws, rows, 5, 19
    FormatOutputSheet ws, "S"
    ws.Columns("P:P").Interior.Color = RGB(255, 247, 230)
End Sub

Private Sub WriteObjectPermissions(ByVal queryRoot As Object, ByVal permissionSetMap As Object)
    Dim ws As Worksheet
    Set ws = PrepareSheet("オブジェクトアクセス権限")

    ws.Range("A1:M1").Merge
    ws.Range("A1").Value = "オブジェクトアクセス権限"
    ws.Range("A2:M2").Merge
    ws.Range("A2").Value = "ObjectPermissions を元に、オブジェクトへのアクセス権限をプロファイル・権限セット単位で出力するシート"
    ws.Range("A4:M4").Value = Array("No", "種別", "プロファイル名", "権限セットAPI名", "権限セット表示名", "対象オブジェクト", "参照", "作成", "編集", "削除", "すべて表示", "すべて変更", "権限ID")

    Dim records As Collection
    Set records = GetQueryRecords(queryRoot)

    Dim rows As Collection
    Set rows = New Collection

    Dim i As Long
    Dim rec As Object
    For i = 1 To records.Count
        Set rec = records(i)
        Dim parent As Object
        Set parent = GetDictObject(rec, "Parent")
        Dim parentId As String
        parentId = CStr(FirstValue(rec, "ParentId", "Parentid"))

        Dim row(1 To 13) As Variant
        Dim ownerType As String
        ownerType = PermissionOwnerTypeFromMap(rec, parent, permissionSetMap, parentId)
        row(1) = i
        row(2) = ownerType
        row(3) = ProfileNameFromMap(rec, parent, permissionSetMap, parentId)
        row(4) = IIf(ownerType = "権限セット", PermissionSetMapValue(permissionSetMap, parentId, "Name", RelationshipValue(rec, parent, "Name", "Parent.Name")), "")
        row(5) = IIf(ownerType = "権限セット", PermissionSetMapValue(permissionSetMap, parentId, "Label", RelationshipValue(rec, parent, "Label", "Parent.Label")), "")
        row(6) = FirstValue(rec, "SObjectType", "SobjectType")
        row(7) = BoolText(DictValue(rec, "PermissionsRead"))
        row(8) = BoolText(DictValue(rec, "PermissionsCreate"))
        row(9) = BoolText(DictValue(rec, "PermissionsEdit"))
        row(10) = BoolText(DictValue(rec, "PermissionsDelete"))
        row(11) = BoolText(DictValue(rec, "PermissionsViewAllRecords"))
        row(12) = BoolText(DictValue(rec, "PermissionsModifyAllRecords"))
        row(13) = parentId
        rows.Add row
    Next

    WriteCollectionRows ws, rows, 5, 13
    FormatOutputSheet ws, "M"
End Sub

Private Sub WriteFieldPermissions(ByVal queryRoot As Object, ByVal describeResult As Object, ByVal permissionSetMap As Object)
    Dim ws As Worksheet
    Set ws = PrepareSheet("項目アクセス権限")

    ws.Range("A1:M1").Merge
    ws.Range("A1").Value = "項目アクセス権限"
    ws.Range("A2:M2").Merge
    ws.Range("A2").Value = "FieldPermissions を元に、項目への参照/編集が許可されているプロファイル・権限セットを1行ずつ出力するシート"
    ws.Range("A4:M4").Value = Array("No", "オブジェクトAPI参照名", "項目API参照名", "項目表示ラベル", "種別", "プロファイル名", "権限セットAPI名", "権限セット表示名", "参照可", "編集可", "権限ID", "項目キー", "取得元")

    Dim labelMap As Object
    Set labelMap = BuildFieldLabelMap(describeResult)

    Dim records As Collection
    Set records = GetQueryRecords(queryRoot)

    Dim rows As Collection
    Set rows = New Collection

    Dim i As Long
    Dim rec As Object
    For i = 1 To records.Count
        Set rec = records(i)

        Dim parent As Object
        Set parent = GetDictObject(rec, "Parent")
        Dim parentId As String
        parentId = CStr(FirstValue(rec, "ParentId", "Parentid"))

        Dim fieldKey As String
        Dim fieldApiName As String
        fieldKey = CStr(DictValue(rec, "Field"))
        fieldApiName = FieldNameFromKey(fieldKey)

        Dim row(1 To 13) As Variant
        Dim ownerType As String
        ownerType = PermissionOwnerTypeFromMap(rec, parent, permissionSetMap, parentId)
        row(1) = i
        row(2) = FirstValue(rec, "SObjectType", "SobjectType")
        row(3) = fieldApiName
        row(4) = DictValue(labelMap, fieldApiName)
        row(5) = ownerType
        row(6) = ProfileNameFromMap(rec, parent, permissionSetMap, parentId)
        row(7) = IIf(ownerType = "権限セット", PermissionSetMapValue(permissionSetMap, parentId, "Name", RelationshipValue(rec, parent, "Name", "Parent.Name")), "")
        row(8) = IIf(ownerType = "権限セット", PermissionSetMapValue(permissionSetMap, parentId, "Label", RelationshipValue(rec, parent, "Label", "Parent.Label")), "")
        row(9) = BoolText(DictValue(rec, "PermissionsRead"))
        row(10) = BoolText(DictValue(rec, "PermissionsEdit"))
        row(11) = parentId
        row(12) = fieldKey
        row(13) = "FieldPermissions"
        rows.Add row
    Next

    WriteCollectionRows ws, rows, 5, 13
    FormatOutputSheet ws, "M"
End Sub

Private Function PrepareSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    ws.Cells.Font.Name = "メイリオ"
    ws.Cells.Font.Size = 10
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    Set PrepareSheet = ws
End Function

Private Sub DeleteOutputSheets()
    Dim i As Long
    Application.DisplayAlerts = False
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Dim sheetName As String
        sheetName = ThisWorkbook.Worksheets(i).Name
        If sheetName <> TARGET_SHEET_NAME And sheetName <> ERROR_SHEET_NAME Then
            ThisWorkbook.Worksheets(i).Delete
        End If
    Next
    Application.DisplayAlerts = True
End Sub

Private Sub FormatOutputSheet(ByVal ws As Worksheet, ByVal lastColumn As String)
    ws.Range("A1:" & lastColumn & "1").Interior.Color = RGB(31, 78, 121)
    ws.Range("A1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2:" & lastColumn & "2").Interior.Color = RGB(234, 242, 248)
    ws.Range("A2").Font.Color = RGB(31, 78, 121)
    ws.Range("A4:" & lastColumn & "4").Interior.Color = RGB(217, 234, 247)
    ws.Range("A4:" & lastColumn & "4").Font.Bold = True
    ws.Range("A4:" & lastColumn & "4").Borders.LineStyle = xlContinuous

    Dim used As Range
    Set used = ws.UsedRange
    used.WrapText = True
    used.VerticalAlignment = xlTop
    used.Borders.LineStyle = xlContinuous
    used.Columns.AutoFit

    ws.Rows(4).AutoFilter
    ws.Range("A5").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub AddInfoRow(ByVal rows As Collection, ByVal category As String, ByVal itemName As String, ByVal itemValue As Variant, ByVal description As String, ByVal sourceName As String)
    Dim row(1 To 5) As Variant
    row(1) = category
    row(2) = itemName
    row(3) = ToCellText(itemValue)
    row(4) = description
    row(5) = sourceName
    rows.Add row
End Sub

Private Sub WriteCollectionRows(ByVal ws As Worksheet, ByVal rows As Collection, ByVal startRow As Long, ByVal colCount As Long)
    If rows.Count = 0 Then Exit Sub

    Dim arr() As Variant
    ReDim arr(1 To rows.Count, 1 To colCount)

    Dim r As Long
    Dim c As Long
    Dim rowValues As Variant
    For r = 1 To rows.Count
        rowValues = rows(r)
        For c = 1 To colCount
            arr(r, c) = rowValues(c)
        Next
    Next

    ws.Cells(startRow, 1).Resize(rows.Count, colCount).Value = arr
End Sub

Private Function GetQueryRecords(ByVal queryRoot As Object) As Collection
    Dim result As Object
    Set result = GetDictObject(queryRoot, "result")
    Set GetQueryRecords = DictValue(result, "records")
End Function

Private Function BuildPermissionSetQuery(ByVal objectPermRoot As Object, ByVal fieldPermRoot As Object) As String
    Dim ids As Object
    Set ids = CreateObject("Scripting.Dictionary")

    AddParentIds ids, objectPermRoot
    AddParentIds ids, fieldPermRoot

    If ids.Count = 0 Then
        BuildPermissionSetQuery = ""
        Exit Function
    End If

    Dim parts() As String
    ReDim parts(0 To ids.Count - 1)

    Dim i As Long
    Dim k As Variant
    i = 0
    For Each k In ids.Keys
        parts(i) = "'" & EscapeSoql(CStr(k)) & "'"
        i = i + 1
    Next

    BuildPermissionSetQuery = "SELECT Id, Name, Label, IsOwnedByProfile, Profile.Name FROM PermissionSet WHERE Id IN (" & Join(parts, ",") & ")"
End Function

Private Sub AddParentIds(ByVal ids As Object, ByVal queryRoot As Object)
    On Error GoTo EH

    Dim records As Collection
    Set records = GetQueryRecords(queryRoot)

    Dim i As Long
    Dim rec As Object
    Dim parentId As String
    For i = 1 To records.Count
        Set rec = records(i)
        parentId = CStr(FirstValue(rec, "ParentId", "Parentid"))
        If Len(parentId) > 0 Then
            If Not ids.Exists(parentId) Then ids.Add parentId, True
        End If
    Next
    Exit Sub
EH:
End Sub

Private Function BuildPermissionSetMap(ByVal queryRoot As Object) As Object
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    Dim records As Collection
    Set records = GetQueryRecords(queryRoot)

    Dim i As Long
    Dim rec As Object
    Dim idValue As String
    For i = 1 To records.Count
        Set rec = records(i)
        idValue = CStr(DictValue(rec, "Id"))
        If Len(idValue) > 0 Then
            map.Add idValue, rec
        End If
    Next

    Set BuildPermissionSetMap = map
End Function

Private Function BuildFieldLabelMap(ByVal describeResult As Object) As Object
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    Dim fields As Collection
    Set fields = DictValue(describeResult, "fields")

    Dim i As Long
    Dim f As Object
    For i = 1 To fields.Count
        Set f = fields(i)
        map(CStr(DictValue(f, "name"))) = DictValue(f, "label")
    Next

    Set BuildFieldLabelMap = map
End Function

Private Function FieldNameFromKey(ByVal fieldKey As String) As String
    Dim p As Long
    p = InStrRev(fieldKey, ".")
    If p > 0 Then
        FieldNameFromKey = Mid$(fieldKey, p + 1)
    Else
        FieldNameFromKey = fieldKey
    End If
End Function

Private Function PermissionOwnerType(ByVal rec As Object, ByVal parent As Object) As String
    If CBoolDefault(RelationshipValue(rec, parent, "IsOwnedByProfile", "Parent.IsOwnedByProfile"), False) Then
        PermissionOwnerType = "プロファイル"
    Else
        PermissionOwnerType = "権限セット"
    End If
End Function

Private Function PermissionOwnerTypeFromMap(ByVal rec As Object, ByVal parent As Object, ByVal permissionSetMap As Object, ByVal parentId As String) As String
    Dim v As Variant
    v = PermissionSetMapValue(permissionSetMap, parentId, "IsOwnedByProfile", "")

    If Len(CStr(v)) > 0 Then
        If CBoolDefault(v, False) Then
            PermissionOwnerTypeFromMap = "プロファイル"
        Else
            PermissionOwnerTypeFromMap = "権限セット"
        End If
    Else
        PermissionOwnerTypeFromMap = PermissionOwnerType(rec, parent)
    End If
End Function

Private Function ProfileName(ByVal rec As Object, ByVal parent As Object) As String
    If Not CBoolDefault(RelationshipValue(rec, parent, "IsOwnedByProfile", "Parent.IsOwnedByProfile"), False) Then
        ProfileName = ""
        Exit Function
    End If

    Dim profile As Object
    Set profile = GetDictObject(parent, "Profile")
    ProfileName = CStr(FirstNonBlank(DictValue(profile, "Name"), DictValue(rec, "Parent.Profile.Name")))
End Function

Private Function ProfileNameFromMap(ByVal rec As Object, ByVal parent As Object, ByVal permissionSetMap As Object, ByVal parentId As String) As String
    If PermissionOwnerTypeFromMap(rec, parent, permissionSetMap, parentId) <> "プロファイル" Then
        ProfileNameFromMap = ""
        Exit Function
    End If

    Dim profileName As Variant
    profileName = PermissionSetMapValue(permissionSetMap, parentId, "Profile.Name", "")
    If Len(CStr(profileName)) > 0 Then
        ProfileNameFromMap = CStr(profileName)
    Else
        ProfileNameFromMap = ProfileName(rec, parent)
    End If
End Function

Private Function PermissionSetMapValue(ByVal permissionSetMap As Object, ByVal parentId As String, ByVal key As String, ByVal fallbackValue As Variant) As Variant
    If permissionSetMap Is Nothing Then
        PermissionSetMapValue = fallbackValue
        Exit Function
    End If

    If Len(parentId) = 0 Or Not permissionSetMap.Exists(parentId) Then
        PermissionSetMapValue = fallbackValue
        Exit Function
    End If

    Dim rec As Object
    Set rec = permissionSetMap(parentId)

    Dim profile As Object
    If key = "Profile.Name" Then
        Set profile = GetDictObject(rec, "Profile")
        PermissionSetMapValue = FirstNonBlank(DictValue(profile, "Name"), DictValue(rec, "Profile.Name"), fallbackValue)
    Else
        PermissionSetMapValue = FirstNonBlank(DictValue(rec, key), fallbackValue)
    End If
End Function

Private Function GetDictObject(ByVal dict As Object, ByVal key As String) As Object
    If dict Is Nothing Then Exit Function

    Dim actualKey As Variant
    actualKey = ResolveKey(dict, key)
    If CStr(actualKey) <> "" Then
        If IsObject(dict(actualKey)) Then Set GetDictObject = dict(actualKey)
    End If
End Function

Private Function DictValue(ByVal dict As Object, ByVal key As String) As Variant
    If dict Is Nothing Then
        DictValue = ""
        Exit Function
    End If

    Dim actualKey As Variant
    actualKey = ResolveKey(dict, key)

    If CStr(actualKey) <> "" Then
        If IsNull(dict(actualKey)) Then
            DictValue = ""
        ElseIf IsObject(dict(actualKey)) Then
            Set DictValue = dict(actualKey)
        Else
            DictValue = dict(actualKey)
        End If
    Else
        DictValue = ""
    End If
End Function

Private Function ResolveKey(ByVal dict As Object, ByVal key As String) As Variant
    If dict Is Nothing Then
        ResolveKey = ""
        Exit Function
    End If

    If dict.Exists(key) Then
        ResolveKey = key
        Exit Function
    End If

    Dim k As Variant
    For Each k In dict.Keys
        If StrComp(CStr(k), key, vbTextCompare) = 0 Then
            ResolveKey = k
            Exit Function
        End If
    Next

    ResolveKey = ""
End Function

Private Function FirstValue(ByVal dict As Object, ParamArray keys() As Variant) As Variant
    Dim i As Long
    Dim v As Variant

    For i = LBound(keys) To UBound(keys)
        v = DictValue(dict, CStr(keys(i)))
        If Not IsObject(v) And Len(CStr(v)) > 0 Then
            FirstValue = v
            Exit Function
        End If
    Next

    FirstValue = ""
End Function

Private Function FirstNonBlank(ParamArray values() As Variant) As Variant
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If Not IsObject(values(i)) And Len(CStr(values(i))) > 0 Then
            FirstNonBlank = values(i)
            Exit Function
        End If
    Next
    FirstNonBlank = ""
End Function

Private Function RelationshipValue(ByVal rec As Object, ByVal parent As Object, ByVal nestedKey As String, ByVal flatKey As String) As Variant
    RelationshipValue = FirstNonBlank(DictValue(parent, nestedKey), DictValue(rec, flatKey))
End Function

Private Function BoolText(ByVal value As Variant) As String
    If IsEmpty(value) Or IsNull(value) Or value = "" Then
        BoolText = ""
    ElseIf CBool(value) Then
        BoolText = "TRUE"
    Else
        BoolText = "FALSE"
    End If
End Function

Private Function CBoolDefault(ByVal value As Variant, ByVal defaultValue As Boolean) As Boolean
    If IsEmpty(value) Or IsNull(value) Or value = "" Then
        CBoolDefault = defaultValue
    Else
        CBoolDefault = CBool(value)
    End If
End Function

Private Function ToCellText(ByVal value As Variant) As String
    If IsObject(value) Then
        ToCellText = ""
    ElseIf IsNull(value) Or IsEmpty(value) Then
        ToCellText = ""
    Else
        ToCellText = CStr(value)
    End If
End Function

Private Function CollectionCount(ByVal value As Variant) As Long
    On Error GoTo EH
    If IsObject(value) Then CollectionCount = value.Count
    Exit Function
EH:
    CollectionCount = 0
End Function

Private Function JoinCollection(ByVal value As Variant, ByVal delimiter As String) As String
    If Not IsObject(value) Then
        JoinCollection = ""
        Exit Function
    End If

    Dim c As Collection
    Set c = value

    Dim parts() As String
    ReDim parts(1 To IIf(c.Count = 0, 1, c.Count))

    Dim i As Long
    For i = 1 To c.Count
        parts(i) = CStr(c(i))
    Next

    If c.Count = 0 Then
        JoinCollection = ""
    Else
        JoinCollection = Join(parts, delimiter)
    End If
End Function

Private Function JoinPicklistValues(ByVal value As Variant) As String
    If Not IsObject(value) Then
        JoinPicklistValues = ""
        Exit Function
    End If

    Dim c As Collection
    Set c = value

    Dim parts As Collection
    Set parts = New Collection

    Dim i As Long
    Dim item As Object
    For i = 1 To c.Count
        Set item = c(i)
        If CBoolDefault(DictValue(item, "active"), True) Then
            parts.Add CStr(DictValue(item, "value"))
        End If
    Next

    Dim arr() As String
    If parts.Count = 0 Then
        JoinPicklistValues = ""
        Exit Function
    End If

    ReDim arr(1 To parts.Count)
    For i = 1 To parts.Count
        arr(i) = parts(i)
    Next
    JoinPicklistValues = Join(arr, "; ")
End Function
