param(
    [Parameter(Mandatory)][string]$FilePath
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $FilePath)) {
    Write-Host "エラー: ファイルが見つかりません: $FilePath" -ForegroundColor Red
    exit 1
}

$ext = [IO.Path]::GetExtension($FilePath).ToLower()
if ($ext -notin '.xls', '.xlsm', '.xlam') {
    Write-Host "エラー: 対応していない形式です: $ext" -ForegroundColor Red
    Write-Host "対応形式: .xls / .xlsm / .xlam"
    exit 1
}

# バックアップ作成
$bakPath = "$FilePath.bak"
Copy-Item $FilePath $bakPath -Force
Write-Host "バックアップ作成: $bakPath"

function Find-DPB([byte[]]$data) {
    $pattern = [System.Text.Encoding]::ASCII.GetBytes('DPB=')
    for ($i = 0; $i -le $data.Length - $pattern.Length; $i++) {
        $match = $true
        for ($j = 0; $j -lt $pattern.Length; $j++) {
            if ($data[$i + $j] -ne $pattern[$j]) { $match = $false; break }
        }
        if ($match) { return $i }
    }
    return -1
}

function Remove-PasswordXls([string]$path) {
    $data = [IO.File]::ReadAllBytes($path)
    $pos = Find-DPB $data
    if ($pos -eq -1) { return $false }
    # DPB= -> DPx=
    $data[$pos + 2] = 0x78
    [IO.File]::WriteAllBytes($path, $data)
    return $true
}

function Remove-PasswordOoxml([string]$path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $tempDir = Join-Path ([IO.Path]::GetTempPath()) "VBAPwdRemover_$(Get-Date -Format yyyyMMddHHmmss)"
    New-Item $tempDir -ItemType Directory -Force | Out-Null

    try {
        # ZIP 展開
        $extractDir = Join-Path $tempDir 'extracted'
        [IO.Compression.ZipFile]::ExtractToDirectory($path, $extractDir)

        # vbaProject.bin を検索
        $vbaProj = Get-ChildItem $extractDir -Recurse -Filter 'vbaProject.bin' | Select-Object -First 1
        if (-not $vbaProj) { return $false }

        # DPB= を書き換え
        $data = [IO.File]::ReadAllBytes($vbaProj.FullName)
        $pos = Find-DPB $data
        if ($pos -eq -1) { return $false }
        $data[$pos + 2] = 0x78
        [IO.File]::WriteAllBytes($vbaProj.FullName, $data)

        # 再 ZIP 化
        Remove-Item $path -Force
        [IO.Compression.ZipFile]::CreateFromDirectory($extractDir, $path)
        return $true
    }
    finally {
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# 実行
Write-Host "処理中: $FilePath"

$result = if ($ext -eq '.xls') {
    Remove-PasswordXls $FilePath
} else {
    Remove-PasswordOoxml $FilePath
}

if ($result) {
    Write-Host ""
    Write-Host "パスワード保護を無効化しました。" -ForegroundColor Green
    Write-Host ""
    Write-Host "次の手順で完全に解除してください:"
    Write-Host "  1. 対象ファイルを開く"
    Write-Host "  2. VBE (Alt+F11) を開く"
    Write-Host "  3. ツール > VBAProject のプロパティ > 保護タブ"
    Write-Host "  4. パスワード欄を空にして OK"
    Write-Host "  5. ファイルを保存"
} else {
    Write-Host ""
    Write-Host "VBAプロジェクトのパスワード情報が見つかりませんでした。" -ForegroundColor Yellow
    Write-Host "ファイルにVBAプロジェクトが含まれていない可能性があります。"
}
