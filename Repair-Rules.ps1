#Requires -Version 5.1
<#
.SYNOPSIS
Lập bảng ánh xạ và sửa hàng loạt liên kết thư mục Move/Copy của Rules đã import.
.EXAMPLE
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\Repair-Rules.ps1 -Stage Plan -SourceReport "D:\Backup cũ\Report\..."
.EXAMPLE
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\Repair-Rules.ps1 -Stage Apply -MappingFile "D:\OutlookMigration\Report\...\folder-mapping.csv" -WhatIf
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [ValidateSet('Plan', 'Apply')][string]$Stage = 'Plan',
    [string]$SourceReport,
    [string]$MappingFile,
    [string]$BackupRulesFile,
    [ValidateNotNullOrEmpty()][string]$DataRoot,
    [ValidateRange(1, 2147483647)][int]$AccountIndex,
    [string]$SourceStoreId,
    [string]$TargetStoreId,
    [switch]$EnableRepairedRules
)
$ErrorActionPreference = 'Stop'
$connection = $null
try {
    if ($PSVersionTable.PSEdition -ne 'Desktop') { throw 'Dùng Windows PowerShell 5.1 (powershell.exe).' }
    $scriptDirectory = $PSScriptRoot
    if (-not $PSBoundParameters.ContainsKey('DataRoot')) { $DataRoot = $scriptDirectory }
    . (Join-Path $scriptDirectory 'Common.ps1')
    . (Join-Path $scriptDirectory 'Repair.Common.ps1')
    if ($Stage -eq 'Plan') {
        if ($WhatIfPreference) { throw 'Plan chỉ đọc Outlook và tạo bảng CSV. Dùng -WhatIf với -Stage Apply để xem trước thao tác lưu.' }
        if (-not $SourceReport) { $SourceReport = (Read-Host 'Đường dẫn thư mục Backup Report máy cũ hoặc rules.csv').Trim().Trim('"') }
        # Validate source before COM attachment so a bad path gives a useful error.
        $null = Import-RepairSourceReport $SourceReport
        $connection = Connect-RunningOutlook
        Invoke-RuleRepairPlan -Session $connection.Session -SourceReport $SourceReport -DataRoot $DataRoot -AccountIndex $AccountIndex -SourceStoreId $SourceStoreId -TargetStoreId $TargetStoreId
    }
    else {
        if (-not $MappingFile) { throw '-Stage Apply cần -MappingFile trỏ tới folder-mapping.csv đã xem lại.' }
        $null = Read-RuleRepairMapping $MappingFile
        $connection = Connect-RunningOutlook
        # The helper owns ShouldProcess; common WhatIf/Confirm preferences flow through.
        Invoke-RuleRepair -Session $connection.Session -MappingFile $MappingFile -DataRoot $DataRoot -BackupRulesFile $BackupRulesFile -EnableRepairedRules:$EnableRepairedRules
    }
}
catch { Write-Error $_ -ErrorAction Continue; exit 1 }
finally {
    if ($null -ne $connection) {
        Release-OutlookReference $connection.Session
        Release-OutlookReference $connection.Application
    }
}
