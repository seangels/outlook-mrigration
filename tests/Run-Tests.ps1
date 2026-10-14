#Requires -Version 5.1
[CmdletBinding()]
param([switch]$Quiet)
$ErrorActionPreference = 'Stop'
Import-Module Pester -MinimumVersion 3.4 -ErrorAction Stop
$result = Invoke-Pester -Script $PSScriptRoot -PassThru -Quiet:$Quiet
Write-Host ('Passed: {0}; Failed: {1}; Skipped: {2}' -f $result.PassedCount, $result.FailedCount, $result.SkippedCount)
if ($result.FailedCount -gt 0) { exit 1 }
