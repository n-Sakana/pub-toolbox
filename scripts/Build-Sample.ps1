param()

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$sampleDir = Join-Path $projectDir 'sample'
$attachmentDir = Join-Path $sampleDir 'attachments'
$outputPath = Join-Path $sampleDir 'vba-toolbox-sample.xlsx'

if (-not (Test-Path -LiteralPath $sampleDir)) {
    New-Item -ItemType Directory -Path $sampleDir -Force | Out-Null
}
if (-not (Test-Path -LiteralPath $attachmentDir)) {
    New-Item -ItemType Directory -Path $attachmentDir -Force | Out-Null
}
if (Test-Path -LiteralPath $outputPath) {
    Remove-Item -LiteralPath $outputPath -Force
}

$attachmentA = Join-Path $attachmentDir 'sample-a.txt'
$attachmentB = Join-Path $attachmentDir 'sample-b.txt'
Set-Content -LiteralPath $attachmentA -Value 'sample attachment A' -Encoding ASCII
Set-Content -LiteralPath $attachmentB -Value 'sample attachment B' -Encoding ASCII

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null

try {
    $wb = $excel.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) {
        $wb.Worksheets.Item($wb.Worksheets.Count).Delete()
    }

    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'source'

    $headers = @('from', 'to', 'cc', 'bcc', 'subject', 'body', 'attachments')
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $ws.Cells.Item(1, $c + 1).Value2 = [string]$headers[$c]
    }

    $rows = @(
        @('', 'user1@example.com', '', '', 'Draft test 01', "Please review the attached note.`nThanks.", ''),
        @('', 'user2@example.com;user2b@example.com', 'manager@example.com', '', 'Draft test 02', 'Multiple recipients example.', ''),
        @('', 'user3@example.com', '', 'audit@example.com', 'Draft test 03', 'Single attachment example.', $attachmentA),
        @('', 'user4@example.com', '', '', 'Draft test 04', 'Two attachments example.', "$attachmentA;$attachmentB")
    )

    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $ws.Cells.Item($r + 2, $c + 1).Value2 = [string]$rows[$r][$c]
        }
    }

    $tableRange = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($rows.Count + 1, $headers.Count))
    $lo = $ws.ListObjects.Add(1, $tableRange, $null, 1)
    $lo.Name = 'source_data'
    $lo.TableStyle = 'TableStyleMedium2'
    [void]$ws.Columns.AutoFit()
    $ws.Columns.Item(5).ColumnWidth = 28
    $ws.Columns.Item(6).ColumnWidth = 36
    $ws.Columns.Item(7).ColumnWidth = 48

    $notes = $wb.Worksheets.Add()
    $notes.Name = 'howto'
    $notes.Range('A1').Value2 = '1. Install dist/vba-toolbox.xlam'
    $notes.Range('A2').Value2 = '2. Use the vba-tools ribbon tab and click Create Draft Sheet'
    $notes.Range('A3').Value2 = '3. Copy columns A:G from source into outlook_draft, or fill outlook_draft with formulas'
    $notes.Range('A4').Value2 = '4. Column from is optional. Leave it blank for the default Outlook account'
    $notes.Range('A5').Value2 = '5. Use ; for multiple mail addresses and attachment paths'
    $notes.Range('A6').Value2 = '6. sample\\attachments contains attachment files for testing'
    $notes.Range('A7').Value2 = '7. Click Run Drafts on the same ribbon tab'
    [void]$notes.Columns.AutoFit()

    $wb.SaveAs($outputPath, 51)
    Write-Host "Built sample workbook: $outputPath" -ForegroundColor Green
} finally {
    if ($null -ne $wb) {
        try { $wb.Close($false) } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    }
    try { $excel.Quit() } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
