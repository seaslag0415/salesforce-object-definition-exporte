param(
    [Parameter(Mandatory = $true)]
    [string] $WorkbookPath
)

$ErrorActionPreference = 'Stop'

$workbookPath = $WorkbookPath
$moduleMain = Join-Path $PSScriptRoot 'modObjectDefineMain.bas'
$moduleJson = Join-Path $PSScriptRoot 'modJsonParser.bas'
$sheetCode = Join-Path $PSScriptRoot 'sheet_target.cls'

if (-not (Test-Path -LiteralPath $workbookPath)) {
    throw "Workbook not found: $workbookPath"
}

$excel = $null
$workbook = $null

try {
    Write-Output 'starting excel'
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    Write-Output 'opening workbook'
    $workbook = $excel.Workbooks.Open($workbookPath)
    Write-Output 'opened workbook'
    $vbProject = $workbook.VBProject

    function Remove-StandardModuleIfExists {
        param(
            [Parameter(Mandatory = $true)] $Project,
            [Parameter(Mandatory = $true)] [string] $Name
        )

        foreach ($component in @($Project.VBComponents)) {
            if ($component.Name -eq $Name -and $component.Type -eq 1) {
                $Project.VBComponents.Remove($component)
                break
            }
        }
    }

    function Convert-ToDefaultEncodingTemp {
        param(
            [Parameter(Mandatory = $true)] [string] $Path,
            [Parameter(Mandatory = $true)] [string] $Name
        )

        $text = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
        $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) $Name
        [System.IO.File]::WriteAllText($tempPath, $text, [System.Text.Encoding]::Default)
        return $tempPath
    }

    Write-Output 'removing old modules'
    Remove-StandardModuleIfExists -Project $vbProject -Name 'modObjectDefineMain'
    Remove-StandardModuleIfExists -Project $vbProject -Name 'modJsonParser'

    Write-Output 'importing modules'
    $moduleMainImport = Convert-ToDefaultEncodingTemp -Path $moduleMain -Name 'modObjectDefineMain_import.bas'
    $moduleJsonImport = Convert-ToDefaultEncodingTemp -Path $moduleJson -Name 'modJsonParser_import.bas'
    [void]$vbProject.VBComponents.Import($moduleMainImport)
    [void]$vbProject.VBComponents.Import($moduleJsonImport)

    Write-Output 'installing sheet event'
    $sheet = $workbook.Worksheets.Item(1)
    $sheetComponent = $vbProject.VBComponents.Item($sheet.CodeName)
    $codeModule = $sheetComponent.CodeModule
    if ($codeModule.CountOfLines -gt 0) {
        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
    }
    $codeModule.AddFromString((Get-Content -LiteralPath $sheetCode -Raw -Encoding UTF8))

    Write-Output 'configuring input sheet'
    $honban = [string]::Concat([char]26412, [char]30058)
    $customUrl = [string]::Concat([char]12459, [char]12473, [char]12479, [char]12512, 'URL')
    $connectText = [string]::Concat([char]25509, [char]32154)
    $objectDefineTitle = [string]::Concat([char]12458, [char]12502, [char]12472, [char]12455, [char]12463, [char]12488, [char]23450, [char]32681, [char]21462, [char]24471)
    $connectionLabel = [string]::Concat([char]25509, [char]32154, [char]20808)
    $objectApiLabel = [string]::Concat([char]12458, [char]12502, [char]12472, [char]12455, [char]12463, [char]12488, 'API', [char]21442, [char]29031, [char]21517)

    $sheet.Range('A1').Value2 = $objectDefineTitle
    $sheet.Range('B3').Value2 = $connectionLabel
    $sheet.Range('B5').Value2 = $objectApiLabel
    $sheet.Range('C3').Validation.Delete()
    $sheet.Range('C3').Validation.Add(3, 1, 1, ($honban + ',Sandbox,' + $customUrl))
    if ([string]::IsNullOrWhiteSpace([string]$sheet.Range('C3').Value2)) {
        $sheet.Range('C3').Value2 = $honban
    }
    if ($sheet.Range('C3').Value2 -eq $honban) {
        $sheet.Range('D3').Value2 = 'https://login.salesforce.com'
    }
    elseif ($sheet.Range('C3').Value2 -eq 'Sandbox') {
        $sheet.Range('D3').Value2 = 'https://test.salesforce.com'
    }
    $sheet.Range('C3:D3').Borders.LineStyle = 1
    $sheet.Range('C5').Borders.LineStyle = 1

    Write-Output 'installing button'
    foreach ($shape in @($sheet.Shapes)) {
        if ($shape.Name -eq 'btnObjectDefinitionExport') {
            $shape.Delete()
        }
    }

    $button = $sheet.Buttons().Add($sheet.Range('B7').Left, $sheet.Range('B7').Top, $sheet.Range('C7:D8').Width, $sheet.Range('C7:D8').Height)
    $button.Name = 'btnObjectDefinitionExport'
    $button.Characters().Text = $connectText
    $button.OnAction = "'" + $workbook.Name + "'!RunObjectDefinitionExport"

    $sheet.Columns('A:E').AutoFit()

    Write-Output 'saving workbook'
    $workbook.Save()
    Write-Output "Installed macro and button: $workbookPath"
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($true)
    }
    if ($excel -ne $null) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
}
