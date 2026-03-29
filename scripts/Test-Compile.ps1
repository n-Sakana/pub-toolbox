$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
& (Join-Path $scriptDir 'Build-Addin.ps1') -OutputName 'vba-toolbox.test.xlam'
