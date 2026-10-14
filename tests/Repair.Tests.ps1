#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
. (Join-Path $projectRoot 'Common.ps1')
$repairCommon = Join-Path $projectRoot 'Repair.Common.ps1'
if (Test-Path -LiteralPath $repairCommon) { . $repairCommon }

function New-RepairTestRule {
    param([string]$Name = 'Khách hàng', [int]$RuleIndex = 1)
    [pscustomobject]@{
        StoreId = 'current-rules'; RuleIndex = $RuleIndex; Name = $Name
        Enabled = $false; ExecutionOrder = $RuleIndex; RuleType = 0; IsLocalRule = $true
        EnabledActionTypes = '1;21'; EnabledConditionTypes = '3'; EnabledExceptionTypes = ''
        MoveStatus = 'MissingTarget'; MoveFolderPath = ''; MoveStoreId = ''; MoveEntryId = ''
        CopyStatus = 'NotEnabled'; CopyFolderPath = ''; CopyStoreId = ''; CopyEntryId = ''
        TechnicalStatus = 'NeedsRepair'; ManualReviewRequired = $true; Notes = 'Requires test email.'
    }
}

function New-RepairTestFolder {
    param([string]$Path = '\\PST mới\Customer\Sub', [string]$StoreId = 'new-pst', [string]$EntryId = 'folder-new-1')
    [pscustomobject]@{
        StoreId = $StoreId; EntryId = $EntryId; FolderPath = $Path
        Name = (($Path -split '\\')[-1]); Depth = 2
    }
}

Describe 'Bulk folder repair planning from old reports' {
    BeforeEach {
        $script:current = New-RepairTestRule
        $script:source = New-RepairTestRule
        $script:source.StoreId = 'old-rules'
        $script:source.Enabled = $true
        $script:source.MoveStatus = 'Resolved'
        $script:source.MoveFolderPath = '\\OldRoot\Customer\Sub'
        $script:source.MoveStoreId = 'old-pst'
        $script:source.MoveEntryId = 'folder-old-1'
        $script:source.TechnicalStatus = 'Checked'
        $script:folders = @(New-RepairTestFolder)
        $script:storeMap = @{ 'old-pst' = 'new-pst' }
    }

    It 'maps the entire relative path after the PST display root has changed' {
        $plan = New-RuleRepairPlan -CurrentRules @($current) -SourceRules @($source) -Folders $folders -TargetStoreMap $storeMap
        @($plan.Rows).Count | Should Be 1
        $row = $plan.Rows[0]
        $row.RowId | Should Be '1:MoveToFolder'
        $row.RuleName | Should Be 'Khách hàng'
        $row.Action | Should Be 'MoveToFolder'
        $row.SourceFolderPath | Should Be '\\OldRoot\Customer\Sub'
        $row.TargetStoreId | Should Be 'new-pst'
        $row.TargetFolderPath | Should Be '\\PST mới\Customer\Sub'
        $row.Apply | Should Be 'Yes'
        $row.MatchStatus | Should Be 'ExactRelativePath'
        $plan.SnapshotHash | Should Match '^[A-Fa-f0-9]{64}$'
    }

    It 'matches the exact relative path case-insensitively' {
        $folders[0].FolderPath = '\\PST mới\CUSTOMER\sub'
        $plan = New-RuleRepairPlan -CurrentRules @($current) -SourceRules @($source) -Folders $folders -TargetStoreMap $storeMap
        $plan.Rows[0].Apply | Should Be 'Yes'
        $plan.Rows[0].TargetFolderPath | Should Be '\\PST mới\CUSTOMER\sub'
    }

    It 'does not confuse identical leaves in different folder branches' {
        $folders = @((New-RepairTestFolder '\\PST mới\Other\Sub'), (New-RepairTestFolder '\\PST mới\Customer\Sub' 'new-pst' 'intended'))
        $plan = New-RuleRepairPlan -CurrentRules @($current) -SourceRules @($source) -Folders $folders -TargetStoreMap $storeMap
        $plan.Rows[0].Apply | Should Be 'Yes'
        $plan.Rows[0].TargetFolderPath | Should Be '\\PST mới\Customer\Sub'
    }

    It 'refuses a leaf-only match when the original branch is missing' {
        $folders = @(New-RepairTestFolder '\\PST mới\Other\Sub')
        $plan = New-RuleRepairPlan -CurrentRules @($current) -SourceRules @($source) -Folders $folders -TargetStoreMap $storeMap
        $plan.Rows[0].Apply | Should Be 'No'
        $plan.Rows[0].TargetFolderPath | Should Be ''
        $plan.Rows[0].TargetStoreId | Should Be 'new-pst'
    }

    It 'never selects an unchosen store even if that store has the original relative path' {
        $folders = @(New-RepairTestFolder '\\Unrelated\Customer\Sub' 'unrelated-store')
        $plan = New-RuleRepairPlan -CurrentRules @($current) -SourceRules @($source) -Folders $folders -TargetStoreMap $storeMap
        $plan.Rows[0].Apply | Should Be 'No'
        $plan.Rows[0].TargetStoreId | Should Be 'new-pst'
        $plan.Rows[0].TargetFolderPath | Should Be ''
    }

    It 'does not guess a target store when no source store mapping was chosen' {
        $plan = New-RuleRepairPlan -CurrentRules @($current) -SourceRules @($source) -Folders $folders -TargetStoreMap @{}
        $plan.Rows[0].Apply | Should Be 'No'
        $plan.Rows[0].TargetStoreId | Should Be ''
        $plan.Rows[0].TargetFolderPath | Should Be ''
    }

    It 'refuses an ambiguous full path within the chosen target store' {
        $folders = @((New-RepairTestFolder), (New-RepairTestFolder '\\AnotherRoot\Customer\Sub' 'new-pst' 'different-entry'))
        $plan = New-RuleRepairPlan -CurrentRules @($current) -SourceRules @($source) -Folders $folders -TargetStoreMap $storeMap
        $plan.Rows[0].Apply | Should Be 'No'
        $plan.Rows[0].TargetFolderPath | Should Be ''
    }

    It 'requires a unique source rule name instead of taking the first duplicate' {
        $duplicate = New-RepairTestRule 'Khách hàng' 2
        $duplicate.StoreId = 'old-rules'
        $plan = New-RuleRepairPlan -CurrentRules @($current) -SourceRules @($source, $duplicate) -Folders $folders -TargetStoreMap $storeMap
        $plan.Rows[0].Apply | Should Be 'No'
        $plan.Rows[0].TargetFolderPath | Should Be ''
    }

    It 'requires a unique current rule name instead of repairing both duplicates' {
        $duplicate = New-RepairTestRule 'Khách hàng' 2
        $plan = New-RuleRepairPlan -CurrentRules @($current, $duplicate) -SourceRules @($source) -Folders $folders -TargetStoreMap $storeMap
        @($plan.Rows).Count | Should Be 2
        @($plan.Rows | Where-Object { $_.Apply -eq 'Yes' }).Count | Should Be 0
    }

    It 'does not pair a same-named source rule with a different RuleType' {
        $source.RuleType = 1
        $plan = New-RuleRepairPlan -CurrentRules @($current) -SourceRules @($source) -Folders $folders -TargetStoreMap $storeMap
        $plan.Rows[0].Apply | Should Be 'No'
        $plan.Rows[0].TargetFolderPath | Should Be ''
    }

    It 'leaves an unmatched rule for manual mapping rather than guessing from its name' {
        $current.Name = 'Sub'
        $plan = New-RuleRepairPlan -CurrentRules @($current) -SourceRules @($source) -Folders $folders -TargetStoreMap $storeMap
        $plan.Rows[0].Apply | Should Be 'No'
        $plan.Rows[0].TargetFolderPath | Should Be ''
        $plan.Rows[0].SourceEnabled | Should Be 'Unknown'
    }

    It 'leaves a source rule with no readable old folder path unselected' {
        $source.MoveFolderPath = ''
        $plan = New-RuleRepairPlan -CurrentRules @($current) -SourceRules @($source) -Folders $folders -TargetStoreMap $storeMap
        $plan.Rows[0].Apply | Should Be 'No'
        $plan.Rows[0].TargetFolderPath | Should Be ''
    }

    It 'excludes folder actions that are disabled or already resolved' {
        $resolved = New-RepairTestRule 'Resolved' 2
        $resolved.MoveStatus = 'Resolved'
        $resolved.MoveFolderPath = '\\PST mới\Customer\Sub'
        $disabled = New-RepairTestRule 'Disabled' 3
        $disabled.MoveStatus = 'NotEnabled'
        $disabled.EnabledActionTypes = '21'
        $inconsistent = New-RepairTestRule 'Unreadable disabled action' 4
        $inconsistent.MoveStatus = 'Unverifiable'
        $inconsistent.EnabledActionTypes = '21'
        $plan = New-RuleRepairPlan -CurrentRules @($resolved, $disabled, $inconsistent) -SourceRules @($source) -Folders $folders -TargetStoreMap $storeMap
        @($plan.Rows).Count | Should Be 0
    }

    It 'can plan a folder action that COM previously marked unverifiable' {
        $current.MoveStatus = 'Unverifiable'
        $plan = New-RuleRepairPlan -CurrentRules @($current) -SourceRules @($source) -Folders $folders -TargetStoreMap $storeMap
        @($plan.Rows).Count | Should Be 1
        $plan.Rows[0].Apply | Should Be 'Yes'
    }

    It 'uses the copy source path for copy actions independently of move actions' {
        $current.EnabledActionTypes = '1;5;21'
        $current.CopyStatus = 'MissingTarget'
        $source.EnabledActionTypes = '1;5;21'
        $source.CopyStatus = 'Resolved'
        $source.CopyStoreId = 'old-pst'
        $source.CopyEntryId = 'old-copy'
        $source.CopyFolderPath = '\\OldRoot\Archive'
        $folders = @((New-RepairTestFolder), (New-RepairTestFolder '\\PST mới\Archive' 'new-pst' 'new-copy'))
        $plan = New-RuleRepairPlan -CurrentRules @($current) -SourceRules @($source) -Folders $folders -TargetStoreMap $storeMap
        @($plan.Rows).Count | Should Be 2
        $copy = @($plan.Rows | Where-Object { $_.Action -eq 'CopyToFolder' })[0]
        $copy.RowId | Should Be '1:CopyToFolder'
        $copy.Apply | Should Be 'Yes'
        $copy.SourceFolderPath | Should Be '\\OldRoot\Archive'
        $copy.TargetFolderPath | Should Be '\\PST mới\Archive'
    }

    It 'records the original enabled state without enabling the current rule' {
        $plan = New-RuleRepairPlan -CurrentRules @($current) -SourceRules @($source) -Folders $folders -TargetStoreMap $storeMap
        $plan.Rows[0].SourceEnabled | Should Be 'True'
        $current.Enabled | Should Be $false
        $source.Enabled = $false
        $plan = New-RuleRepairPlan -CurrentRules @($current) -SourceRules @($source) -Folders $folders -TargetStoreMap $storeMap
        $plan.Rows[0].SourceEnabled | Should Be 'False'
        $source.Enabled = $null
        $plan = New-RuleRepairPlan -CurrentRules @($current) -SourceRules @($source) -Folders $folders -TargetStoreMap $storeMap
        $plan.Rows[0].SourceEnabled | Should Be 'Unknown'
    }

    It 'produces stable rule identifiers if a second affected action is added' {
        $first = New-RuleRepairPlan -CurrentRules @($current) -SourceRules @($source) -Folders $folders -TargetStoreMap $storeMap
        $current.CopyStatus = 'MissingTarget'
        $current.EnabledActionTypes = '1;5;21'
        $second = New-RuleRepairPlan -CurrentRules @($current) -SourceRules @($source) -Folders $folders -TargetStoreMap $storeMap
        ($second.Rows | Where-Object { $_.Action -eq 'MoveToFolder' }).RowId | Should Be $first.Rows[0].RowId
    }
}

Describe 'Repair snapshot freshness' {
    It 'returns the same hash for equivalent raw report rows including CSV round trips' {
        $rules = @(New-RepairTestRule)
        $csvRules = @($rules | ConvertTo-Csv -NoTypeInformation | ConvertFrom-Csv)
        (Get-RepairSnapshotHash -Rules $rules) | Should Be (Get-RepairSnapshotHash -Rules $csvRules)
    }

    It 'detects an enabled state changed after a repair plan was created' {
        $rule = New-RepairTestRule
        $before = Get-RepairSnapshotHash -Rules @($rule)
        $rule.Enabled = $true
        (Get-RepairSnapshotHash -Rules @($rule)) | Should Not Be $before
    }

    It 'detects a changed execution order after a repair plan was created' {
        $rule = New-RepairTestRule
        $before = Get-RepairSnapshotHash -Rules @($rule)
        $rule.ExecutionOrder = 9
        (Get-RepairSnapshotHash -Rules @($rule)) | Should Not Be $before
    }

    It 'detects a changed destination after a repair plan was created' {
        $rule = New-RepairTestRule
        $before = Get-RepairSnapshotHash -Rules @($rule)
        $rule.MoveStatus = 'Resolved'
        $rule.MoveFolderPath = '\\PST mới\Different'
        $rule.MoveStoreId = 'new-pst'
        $rule.MoveEntryId = 'different-folder'
        (Get-RepairSnapshotHash -Rules @($rule)) | Should Not Be $before
    }
}
