param(
    [string]$OutputName = 'vba-toolbox.xlam'
)

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$srcDir = Join-Path $projectDir 'src'
$distDir = Join-Path $projectDir 'dist'
$outputPath = Join-Path $distDir $OutputName

function Get-ModuleFiles {
    param([string]$Root)

    $patterns = @('*.bas', '*.cls', '*.frm')
    $items = foreach ($pattern in $patterns) {
        Get-ChildItem -Path $Root -Recurse -File -Filter $pattern | Sort-Object FullName
    }
    return @($items)
}

function Set-CodeModuleText {
    param(
        [object]$CodeModule,
        [string]$Code
    )

    if ($CodeModule.CountOfLines -gt 0) {
        $CodeModule.DeleteLines(1, $CodeModule.CountOfLines)
    }
    if (-not [string]::IsNullOrWhiteSpace($Code)) {
        $CodeModule.AddFromString($Code)
    }
}

function Release-ComObject {
    param([object]$ComObject)

    if ($null -eq $ComObject) { return }
    if ([System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

if (-not (Test-Path -LiteralPath $distDir)) {
    New-Item -ItemType Directory -Path $distDir -Force | Out-Null
}
if (Test-Path -LiteralPath $outputPath) {
    Remove-Item -LiteralPath $outputPath -Force
}

$moduleFiles = Get-ModuleFiles -Root $srcDir
if ($moduleFiles.Count -eq 0) {
    throw 'No VBA source files were found under src.'
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null
$vbProj = $null
$docComp = $null

try {
    $wb = $excel.Workbooks.Add()
    $vbProj = $wb.VBProject
    if ($null -eq $vbProj) {
        throw 'VBA project access is not trusted.'
    }
    try {
        $null = $vbProj.VBComponents.Count
    } catch {
        throw 'Enable Trust access to the VBA project object model in Excel.'
    }

    foreach ($file in $moduleFiles) {
        $vbProj.VBComponents.Import($file.FullName) | Out-Null
    }

    $thisWorkbookLines = @(
        'Option Explicit',
        '',
        'Private Sub Workbook_Open()',
        '    On Error Resume Next',
        '    ToolboxMain.InitAddin',
        'End Sub',
        '',
        'Private Sub Workbook_AddinInstall()',
        '    On Error Resume Next',
        '    ToolboxMain.InitAddin',
        'End Sub',
        '',
        'Private Sub Workbook_AddinUninstall()',
        '    On Error Resume Next',
        '    ToolboxMain.ShutdownAddin',
        'End Sub',
        '',
        'Private Sub Workbook_BeforeClose(Cancel As Boolean)',
        '    On Error Resume Next',
        '    ToolboxMain.ShutdownAddin',
        '    Me.Saved = True',
        'End Sub'
    )
    $thisWorkbookCode = [string]::Join("`r`n", $thisWorkbookLines)

    $docComp = $vbProj.VBComponents.Item('ThisWorkbook')
    Set-CodeModuleText -CodeModule $docComp.CodeModule -Code $thisWorkbookCode

    $wb.IsAddin = $true
    $wb.SaveAs($outputPath, 55)
    Write-Host "Built add-in: $outputPath" -ForegroundColor Green
} finally {
    Release-ComObject -ComObject $docComp
    Release-ComObject -ComObject $vbProj
    if ($null -ne $wb) {
        try { $wb.Close($false) } catch {}
        Release-ComObject -ComObject $wb
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject -ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$customUIPath = Join-Path $srcDir 'customUI14.xml'
if (Test-Path -LiteralPath $customUIPath) {
    Write-Host 'Injecting customUI ribbon...' -ForegroundColor Cyan
    $tempDir = Join-Path $env:TEMP ("vba_toolbox_build_" + (Get-Random))
    $zipPath = $outputPath + '.zip'

    Copy-Item -LiteralPath $outputPath -Destination $zipPath -Force
    Expand-Archive -Path $zipPath -DestinationPath $tempDir -Force
    Remove-Item -LiteralPath $zipPath -Force

    $cuiDir = Join-Path $tempDir 'customUI'
    New-Item -ItemType Directory -Path $cuiDir -Force | Out-Null
    Copy-Item -LiteralPath $customUIPath -Destination (Join-Path $cuiDir 'customUI14.xml') -Force

    $ctPath = Join-Path $tempDir '[Content_Types].xml'
    $ctXml = [xml](Get-Content -LiteralPath $ctPath)
    $ctNs = $ctXml.DocumentElement.NamespaceURI
    $existingCt = $ctXml.Types.Override | Where-Object { $_.PartName -eq '/customUI/customUI14.xml' }
    if (-not $existingCt) {
        $node = $ctXml.CreateElement('Override', $ctNs)
        $node.SetAttribute('PartName', '/customUI/customUI14.xml')
        $node.SetAttribute('ContentType', 'application/xml')
        $ctXml.DocumentElement.AppendChild($node) | Out-Null
        $ctXml.Save($ctPath)
    }

    $relsPath = Join-Path $tempDir '_rels\.rels'
    $relsXml = [xml](Get-Content -LiteralPath $relsPath)
    $relsNs = $relsXml.DocumentElement.NamespaceURI
    $existingRel = $relsXml.Relationships.Relationship | Where-Object { $_.Target -eq 'customUI/customUI14.xml' }
    if (-not $existingRel) {
        $relNode = $relsXml.CreateElement('Relationship', $relsNs)
        $relNode.SetAttribute('Id', 'rIdCustomUI')
        $relNode.SetAttribute('Type', 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility')
        $relNode.SetAttribute('Target', 'customUI/customUI14.xml')
        $relsXml.DocumentElement.AppendChild($relNode) | Out-Null
        $relsXml.Save($relsPath)
    }

    Remove-Item -LiteralPath $outputPath -Force
    Compress-Archive -Path (Join-Path $tempDir '*') -DestinationPath $zipPath -Force
    Move-Item -LiteralPath $zipPath -Destination $outputPath -Force
    Remove-Item -LiteralPath $tempDir -Recurse -Force
    Write-Host 'customUI injected.' -ForegroundColor Green
}
