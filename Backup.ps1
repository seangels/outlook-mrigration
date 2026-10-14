#Requires -Version 5.1
<#
.SYNOPSIS
Sao lưu file Rules .rwz và kiểm kê Outlook Classic đang mở (nếu có).
.EXAMPLE
powershell.exe -NoProfile -STA -File .\Backup.ps1 -RulesFile "D:\Backup cũ\Rules.rwz"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$RulesFile,
    [ValidateNotNullOrEmpty()][string]$DataRoot,
    [switch]$SkipOutlook
)
$ErrorActionPreference = 'Stop'
try {
    if ($PSVersionTable.PSEdition -ne 'Desktop') { throw 'Dùng Windows PowerShell 5.1 (powershell.exe), không dùng pwsh.' }
    # PS5.1 can evaluate parameter defaults before PSScriptRoot is available.
    # Resolve the default in the script body; explicit empty values stay invalid.
    $scriptDirectory = $PSScriptRoot
    if (-not $PSBoundParameters.ContainsKey('DataRoot')) { $DataRoot = $scriptDirectory }
    . (Join-Path $scriptDirectory 'Common.ps1')
    Invoke-RulesBackup -RulesFile $RulesFile -DataRoot $DataRoot -SkipOutlook:$SkipOutlook
}
catch { Write-Error $_ -ErrorAction Continue; exit 1 }
