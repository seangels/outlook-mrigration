$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
. (Join-Path $projectRoot 'Common.ps1')
$repairHelpers = Join-Path $projectRoot 'Repair.Common.ps1'
if (Test-Path -LiteralPath $repairHelpers) { . $repairHelpers }

function New-RepairTestCollection($Items) {
    $result = [pscustomobject]@{ Values = @($Items); Count = @($Items).Count }
    $result | Add-Member ScriptMethod Item { param($Index) $this.Values[$Index - 1] }
    $result
}

function New-RepairTestSession {
    $target = [pscustomobject]@{ Name = 'Khách hàng'; FolderPath = '\\PST mới\Khách hàng'; EntryID = 'target-1'; StoreID = 'target-store'; Folders = (New-RepairTestCollection @()) }
    $move = [pscustomobject]@{ Enabled = $true; ActionType = 1; Folder = $null }
    $copy = [pscustomobject]@{ Enabled = $false; ActionType = 5; Folder = $null }
    $actions = New-RepairTestCollection @($move, $copy)
    $actions | Add-Member NoteProperty MoveToFolder $move
    $actions | Add-Member NoteProperty CopyToFolder $copy
    $rule = [pscustomobject]@{ Name = 'Rule khách'; Enabled = $false; ExecutionOrder = 1; RuleType = 0; IsLocalRule = $true; Actions = $actions; Conditions = (New-RepairTestCollection @()); Exceptions = (New-RepairTestCollection @()) }
    $rules = New-RepairTestCollection @($rule)
    $rules | Add-Member NoteProperty SaveCalls 0
    $rules | Add-Member NoteProperty ThrowOnSave $false
    $rules | Add-Member ScriptMethod Save {
        param($ShowProgress)
        $this.SaveCalls++
        if ($this.ThrowOnSave) { throw 'Save failed; persisted state uncertain' }
    }
    $ruleRoot = [pscustomobject]@{ Name = 'Mailbox'; FolderPath = '\\Mailbox'; EntryID = 'rule-root'; StoreID = 'rule-store'; Folders = (New-RepairTestCollection @()) }
    $targetRoot = [pscustomobject]@{ Name = 'PST mới'; FolderPath = '\\PST mới'; EntryID = 'target-root'; StoreID = 'target-store'; Folders = (New-RepairTestCollection @($target)) }
    $store = [pscustomobject]@{ DisplayName = 'Mailbox'; StoreID = 'rule-store'; FilePath = 'C:\mail.pst'; Root = $ruleRoot; RulesList = $rules }
    $store | Add-Member ScriptMethod GetRootFolder { $this.Root }
    $store | Add-Member ScriptMethod GetRules { $this.RulesList }
    $targetStore = [pscustomobject]@{ DisplayName = 'PST mới'; StoreID = 'target-store'; FilePath = 'D:\archive.pst'; Root = $targetRoot }
    $targetStore | Add-Member ScriptMethod GetRootFolder { $this.Root }
    $account = [pscustomobject]@{ DisplayName = 'POP'; SmtpAddress = 'user@example.com'; AccountType = 2; DeliveryStore = $store }
    $session = [pscustomobject]@{ Accounts = (New-RepairTestCollection @($account)); Stores = (New-RepairTestCollection @($store, $targetStore)); Target = $target }
    $session | Add-Member ScriptMethod GetFolderFromID {
        param($EntryId, $StoreId)
        if ($EntryId -ne 'target-1' -or $StoreId -ne 'target-store') { throw 'Folder not found' }
        $this.Target
    }
    $session
}

function Copy-RepairTestEntrypoint {
    param([string]$Directory)
    New-Item -ItemType Directory -Path $Directory | Out-Null
    foreach ($name in @('Repair-Rules.ps1', 'Common.ps1', 'Repair.Common.ps1')) {
        Copy-Item -LiteralPath (Join-Path $projectRoot $name) -Destination (Join-Path $Directory $name)
    }
    $fixtureCode = foreach ($name in @('New-RepairTestCollection', 'New-RepairTestSession')) {
        'function ' + $name + ' {' + (Get-Command $name).Definition + '}'
    }
    $fixtureCode += 'function Connect-RunningOutlook { [pscustomobject]@{ Application = $null; Session = (New-RepairTestSession) } }'
    Add-Content -LiteralPath (Join-Path $Directory 'Common.ps1') -Value ($fixtureCode -join "`r`n") -Encoding UTF8
}

Describe 'Bulk repair preflight and writes' {
    BeforeEach {
        $script:session = New-RepairTestSession
        $script:caseRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $caseRoot | Out-Null
        $script:backupFile = Join-Path $caseRoot 'before.rwz'
        [IO.File]::WriteAllBytes($backupFile, [byte[]](1, 2, 3, 4))
        $script:current = (Get-StoreRuleReport -Session $session -StoreId 'rule-store').Rules
        $source = $current | Select-Object *
        $source.StoreId = 'old-rule-store'; $source.Enabled = $true
        $source.MoveFolderPath = '\\Old PST\Khách hàng'; $source.MoveStoreId = 'old-target'; $source.MoveStatus = 'Resolved'; $source.MoveEntryId = 'old-folder'
        $inventory = Get-OutlookInventory -Session $session
        $plan = New-RuleRepairPlan -CurrentRules $current -SourceRules @($source) -Folders $inventory.Folders -TargetStoreMap @{ 'old-target' = 'target-store' }
        $script:mapping = Save-RuleRepairPlan -Plan $plan -Directory $caseRoot -RuleStoreId 'rule-store'
    }

    It 'previews without saving or changing a rule even when no backup was supplied' {
        $result = Invoke-RuleRepair -Session $session -MappingFile $mapping -DataRoot (Join-Path $caseRoot 'out') -WhatIf
        $result.Status | Should Be 'Preview'
        $session.Stores.Item(1).RulesList.SaveCalls | Should Be 0
        $session.Stores.Item(1).RulesList.Item(1).Actions.MoveToFolder.Folder | Should BeNullOrEmpty
        (Test-Path -LiteralPath (Join-Path $caseRoot 'out')) | Should Be $false
    }

    It 'sets mapped folders and saves once while preserving rule and action enabled state' {
        $result = Invoke-RuleRepair -Session $session -MappingFile $mapping -DataRoot (Join-Path $caseRoot 'out') -BackupRulesFile $backupFile -Confirm:$false
        $result.Status | Should Be 'SavedVerified'
        $rules = $session.Stores.Item(1).RulesList
        $rules.SaveCalls | Should Be 1
        $rules.Item(1).Actions.MoveToFolder.Folder.EntryID | Should Be 'target-1'
        $rules.Item(1).Enabled | Should Be $false
        $rules.Item(1).Actions.MoveToFolder.Enabled | Should Be $true
        $rules.Item(1).ExecutionOrder | Should Be 1
        (Test-Path -LiteralPath (Join-Path $result.ReportDirectory 'repair-results.csv')) | Should Be $true
    }

    It 'rejects stale rule metadata before setting any folders' {
        $session.Stores.Item(1).RulesList.Item(1).Name = 'Changed after planning'
        { Invoke-RuleRepair -Session $session -MappingFile $mapping -DataRoot (Join-Path $caseRoot 'out') -BackupRulesFile $backupFile -Confirm:$false } | Should Throw
        $session.Stores.Item(1).RulesList.SaveCalls | Should Be 0
        $session.Stores.Item(1).RulesList.Item(1).Actions.MoveToFolder.Folder | Should BeNullOrEmpty
    }

    It 'rejects an invalid target path before mutating any actions' {
        $rows = @(Import-Csv -LiteralPath $mapping)
        $rows[0].TargetFolderPath = '\\PST mới\Missing'
        $rows | Export-Csv -LiteralPath $mapping -Encoding UTF8 -NoTypeInformation
        { Invoke-RuleRepair -Session $session -MappingFile $mapping -DataRoot (Join-Path $caseRoot 'out') -BackupRulesFile $backupFile -Confirm:$false } | Should Throw
        $session.Stores.Item(1).RulesList.SaveCalls | Should Be 0
        $session.Stores.Item(1).RulesList.Item(1).Actions.MoveToFolder.Folder | Should BeNullOrEmpty
    }

    It 'does not set the first action when the second selected target fails preflight' {
        $rule = $session.Stores.Item(1).RulesList.Item(1)
        $rule.Actions.CopyToFolder.Enabled = $true
        $current = (Get-StoreRuleReport -Session $session -StoreId 'rule-store').Rules
        $source = $current | Select-Object *
        $source.StoreId = 'old-rules'
        $source.MoveFolderPath = '\\Old\Khách hàng'; $source.MoveStoreId = 'old-pst'
        $source.CopyFolderPath = '\\Old\Khách hàng'; $source.CopyStoreId = 'old-pst'
        $inventory = Get-OutlookInventory -Session $session
        $plan = New-RuleRepairPlan -CurrentRules $current -SourceRules @($source) -Folders $inventory.Folders -TargetStoreMap @{ 'old-pst' = 'target-store' }
        $twoRowsDirectory = Join-Path $caseRoot 'two-rows'
        $twoMapping = Save-RuleRepairPlan -Plan $plan -Directory $twoRowsDirectory -RuleStoreId 'rule-store'
        $rows = @(Import-Csv -LiteralPath $twoMapping)
        $rows.Count | Should Be 2
        $rows[1].TargetFolderPath = '\\PST mới\Missing'
        $rows | Export-Csv -LiteralPath $twoMapping -NoTypeInformation -Encoding UTF8
        { Invoke-RuleRepair -Session $session -MappingFile $twoMapping -DataRoot (Join-Path $caseRoot 'out') -BackupRulesFile $backupFile -Confirm:$false } | Should Throw
        $rule.Actions.MoveToFolder.Folder | Should BeNullOrEmpty
        $rule.Actions.CopyToFolder.Folder | Should BeNullOrEmpty
        $session.Stores.Item(1).RulesList.SaveCalls | Should Be 0
    }

    It 'propagates WhatIf through the native entry script with DataRoot omitted' {
        $entryDirectory = Join-Path $caseRoot 'native-wrapper'
        Copy-RepairTestEntrypoint -Directory $entryDirectory
        $entry = Join-Path $entryDirectory 'Repair-Rules.ps1'
        & powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File $entry -Stage Apply -MappingFile $mapping -WhatIf *> (Join-Path $caseRoot 'native-whatif.log')
        $LASTEXITCODE | Should Be 0
        (Get-Content -LiteralPath (Join-Path $caseRoot 'native-whatif.log') -Raw -Encoding Unicode) | Should Match 'Preview'
        (Test-Path -LiteralPath (Join-Path $entryDirectory 'Report')) | Should Be $false
        (Test-Path -LiteralPath (Join-Path $entryDirectory 'Rules')) | Should Be $false
    }

    It 'requires a pre-repair rwz backup for actual writes' {
        { Invoke-RuleRepair -Session $session -MappingFile $mapping -DataRoot (Join-Path $caseRoot 'out') -Confirm:$false } | Should Throw
        $session.Stores.Item(1).RulesList.SaveCalls | Should Be 0
    }

    It 'does not retry Save or claim success when Save throws' {
        $session.Stores.Item(1).RulesList.ThrowOnSave = $true
        { Invoke-RuleRepair -Session $session -MappingFile $mapping -DataRoot (Join-Path $caseRoot 'out') -BackupRulesFile $backupFile -Confirm:$false } | Should Throw
        $session.Stores.Item(1).RulesList.SaveCalls | Should Be 1
        $journal = Get-ChildItem -LiteralPath (Join-Path $caseRoot 'out/Report') -Filter 'repair-results.csv' -Recurse | Select-Object -First 1
        (Import-Csv -LiteralPath $journal.FullName)[0].Status | Should Be 'SaveFailedUnknown'
    }

    It 'enables an originally enabled repaired rule only when explicitly requested' {
        $result = Invoke-RuleRepair -Session $session -MappingFile $mapping -DataRoot (Join-Path $caseRoot 'out') -BackupRulesFile $backupFile -EnableRepairedRules -Confirm:$false
        $result.Status | Should Be 'SavedVerified'
        $session.Stores.Item(1).RulesList.Item(1).Enabled | Should Be $true
    }

    It 'preserves a rule originally disabled in the source even when enabling was requested' {
        $sidecar = $mapping + '.plan.json'
        $planData = Get-Content -LiteralPath $sidecar -Raw -Encoding UTF8 | ConvertFrom-Json
        $planData.Rows[0].SourceEnabled = 'False'
        $planData | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $sidecar -Encoding UTF8
        $rows = @(Import-Csv -LiteralPath $mapping)
        $rows[0].SourceEnabled = 'False'
        $rows | Export-Csv -LiteralPath $mapping -Encoding UTF8 -NoTypeInformation
        $result = Invoke-RuleRepair -Session $session -MappingFile $mapping -DataRoot (Join-Path $caseRoot 'out') -BackupRulesFile $backupFile -EnableRepairedRules -Confirm:$false
        $session.Stores.Item(1).RulesList.Item(1).Enabled | Should Be $false
        $result.Results[0].EnableNote | Should Match 'giữ nguyên'
    }

    It 'never enables an old-machine rule while repairing its folder' {
        $rule = $session.Stores.Item(1).RulesList.Item(1)
        $rule.Conditions = New-RepairTestCollection @([pscustomobject]@{ Enabled = $true; ConditionType = 28 })
        $fresh = Get-StoreRuleReport -Session $session -StoreId 'rule-store'
        $sidecar = $mapping + '.plan.json'
        $planData = Get-Content -LiteralPath $sidecar -Raw -Encoding UTF8 | ConvertFrom-Json
        $planData.SnapshotHash = Get-RepairSnapshotHash $fresh.Rules
        $planData | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $sidecar -Encoding UTF8
        $result = Invoke-RuleRepair -Session $session -MappingFile $mapping -DataRoot (Join-Path $caseRoot 'out') -BackupRulesFile $backupFile -EnableRepairedRules -Confirm:$false
        $rule.Enabled | Should Be $false
        $rule.Conditions.Item(1).Enabled | Should Be $true
        $result.Results[0].EnableNote | Should Match 'OnOtherMachine'
    }

    It 'does not save after an action setter rejects a target' {
        $move = $session.Stores.Item(1).RulesList.Item(1).Actions.MoveToFolder
        $move | Add-Member ScriptProperty Folder { $null } { param($Value) throw 'Read-only target' } -Force
        { Invoke-RuleRepair -Session $session -MappingFile $mapping -DataRoot (Join-Path $caseRoot 'out') -BackupRulesFile $backupFile -Confirm:$false } | Should Throw
        $session.Stores.Item(1).RulesList.SaveCalls | Should Be 0
        $journal = Get-ChildItem -LiteralPath (Join-Path $caseRoot 'out/Report') -Filter 'repair-results.csv' -Recurse | Select-Object -First 1
        (Import-Csv -LiteralPath $journal.FullName)[0].Status | Should Be 'NotSaved'
    }

    It 'does not report success when Enabled cannot be read after saving a disabled rule' {
        $rules = $session.Stores.Item(1).RulesList
        $rules | Add-Member ScriptMethod Save {
            param($ShowProgress)
            $this.SaveCalls++
            $this.Item(1) | Add-Member ScriptProperty Enabled { throw 'Enabled unavailable after Save' } -Force
        } -Force
        { Invoke-RuleRepair -Session $session -MappingFile $mapping -DataRoot (Join-Path $caseRoot 'out') -BackupRulesFile $backupFile -Confirm:$false } | Should Throw
        $rules.SaveCalls | Should Be 1
        $journal = Get-ChildItem -LiteralPath (Join-Path $caseRoot 'out/Report') -Filter 'repair-results.csv' -Recurse | Select-Object -First 1
        (Import-Csv -LiteralPath $journal.FullName)[0].Status | Should Be 'VerificationFailed'
    }

    It 'rejects duplicate mapping rows instead of applying the same action twice' {
        $rows = @(Import-Csv -LiteralPath $mapping)
        @($rows[0], $rows[0]) | Export-Csv -LiteralPath $mapping -Encoding UTF8 -NoTypeInformation
        { Invoke-RuleRepair -Session $session -MappingFile $mapping -DataRoot (Join-Path $caseRoot 'out') -BackupRulesFile $backupFile -Confirm:$false } | Should Throw
        $session.Stores.Item(1).RulesList.SaveCalls | Should Be 0
    }
}

Describe 'Repair source reports and plan workflow' {
    It 'imports an old report directory with Unicode paths and keeps explicit source store identity' {
        $session = New-RepairTestSession
        $sourceRules = (Get-StoreRuleReport -Session $session -StoreId 'rule-store').Rules
        $caseRoot = Join-Path $TestDrive ('báo cáo nguồn ' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $caseRoot | Out-Null
        $sourceRules | Export-Csv -LiteralPath (Join-Path $caseRoot 'rules.csv') -NoTypeInformation -Encoding UTF8
        $source = Import-RepairSourceReport -Path $caseRoot
        $source.Rules[0].Name | Should Be 'Rule khách'
        $storeId = Select-RepairSourceStore -Source $source -SourceStoreId 'rule-store' -SmtpAddress 'user@example.com'
        $storeId | Should Be 'rule-store'
    }

    It 'rejects unrelated CSV input rather than treating it as a source Rules report' {
        $path = Join-Path $TestDrive 'unrelated.csv'
        'Name,Email', 'Someone,someone@example.com' | Set-Content -LiteralPath $path -Encoding UTF8
        { Import-RepairSourceReport -Path $path } | Should Throw
    }

    It 'requires an explicitly chosen destination store even when only one exists' {
        $source = [pscustomobject]@{ MoveStoreId = 'old-pst'; CopyStoreId = ''; MoveFolderPath = '\\Old\X'; CopyFolderPath = '' }
        $stores = @([pscustomobject]@{ StoreId = 'new-pst'; DisplayName = 'New'; FilePath = 'D:\new.pst' })
        Mock Read-Host { 'HUY' }
        { Select-RepairTargetStores -SourceRules @($source) -Stores $stores } | Should Throw
    }

    It 'runs the native Plan command with default data root and explicit source and target stores' {
        $caseRoot = Join-Path $TestDrive ('native-plan-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $caseRoot | Out-Null
        $entryDirectory = Join-Path $caseRoot 'scripts'
        Copy-RepairTestEntrypoint -Directory $entryDirectory
        $session = New-RepairTestSession
        $source = (Get-StoreRuleReport -Session $session -StoreId 'rule-store').Rules | Select-Object *
        $source.StoreId = 'old-rule-store'; $source.MoveFolderPath = '\\Old\Khách hàng'; $source.MoveStoreId = 'old-pst'; $source.MoveStatus = 'Resolved'
        $sourceFile = Join-Path $caseRoot 'source-rules.csv'
        $source | Export-Csv -LiteralPath $sourceFile -NoTypeInformation -Encoding UTF8
        $entry = Join-Path $entryDirectory 'Repair-Rules.ps1'
        & powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File $entry -Stage Plan -SourceReport $sourceFile -AccountIndex 1 -SourceStoreId 'old-rule-store' -TargetStoreId 'target-store' *> (Join-Path $caseRoot 'native-plan.log')
        $LASTEXITCODE | Should Be 0
        $maps = @(Get-ChildItem -LiteralPath (Join-Path $entryDirectory 'Report') -Filter 'folder-mapping.csv' -Recurse)
        $maps.Count | Should Be 1
        (Import-Csv -LiteralPath $maps[0].FullName)[0].TargetFolderPath | Should Be '\\PST mới\Khách hàng'
        (Import-Csv -LiteralPath $maps[0].FullName)[0].Apply | Should Be 'Yes'
    }
}
