#Requires -Version 5.1
<#
.SYNOPSIS
Kiểm kê, hướng dẫn import .rwz thủ công và kiểm tra Rules trong Outlook Classic.
.EXAMPLE
powershell.exe -NoProfile -STA -File .\Restore.ps1 -RulesFile ".\Rules\Rules.rwz"
.EXAMPLE
powershell.exe -NoProfile -STA -File .\Restore.ps1 -RulesFile ".\Rules\Rules.rwz" -Stage GuideImport -AccountIndex 1
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$RulesFile,
    [ValidateSet('Inspect', 'GuideImport', 'Verify')][string]$Stage = 'Inspect',
    [ValidateNotNullOrEmpty()][string]$DataRoot,
    [ValidateRange(1, 2147483647)][int]$AccountIndex
)
$ErrorActionPreference = 'Stop'
try {
    if ($PSVersionTable.PSEdition -ne 'Desktop') { throw 'Dùng Windows PowerShell 5.1 (powershell.exe), không dùng pwsh.' }
    # Initialize here rather than in param so the script path is available in PS5.1.
    $scriptDirectory = $PSScriptRoot
    if (-not $PSBoundParameters.ContainsKey('DataRoot')) { $DataRoot = $scriptDirectory }
    . (Join-Path $scriptDirectory 'Common.ps1')
    Invoke-RulesRestore -RulesFile $RulesFile -DataRoot $DataRoot -Stage $Stage -AccountIndex $AccountIndex
}
catch { Write-Error $_ -ErrorAction Continue; exit 1 }
