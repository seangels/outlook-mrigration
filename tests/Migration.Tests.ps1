$ErrorActionPreference = 'Stop'
$commonPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Common.ps1'
if (Test-Path -LiteralPath $commonPath) { . $commonPath }

Describe 'Rules backup integrity' {
    BeforeEach {
        $script:caseRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $caseRoot | Out-Null
        $script:source = Join-Path $caseRoot 'quy tắc mẫu.rwz'
        [IO.File]::WriteAllBytes($source, [byte[]](1, 2, 3, 4))
    }

    It 'rejects a missing file, a wrong extension, and an empty file' {
        { Get-RulesFileValidation -Path (Join-Path $caseRoot 'missing.rwz') } | Should Throw
        $wrong = Join-Path $caseRoot 'wrong.txt'
        [IO.File]::WriteAllBytes($wrong, [byte[]](1))
        { Get-RulesFileValidation -Path $wrong } | Should Throw
        $empty = Join-Path $caseRoot 'empty.rwz'
        [IO.File]::WriteAllBytes($empty, [byte[]]@())
        { Get-RulesFileValidation -Path $empty } | Should Throw
    }

    It 'accepts an external rwz but marks its integrity unverified' {
        $result = Get-RulesFileValidation -Path $source
        $result.IntegrityStatus | Should Be 'Unverified'
        $result.Length | Should Be 4
        $result.SHA256 | Should Be '9F64A747E1B97F131FABB6B447296C9B6F0201E79FB3C5356E6C77E89B6A806A'
    }

    It 'copies Unicode paths without overwriting and writes a verifiable sidecar' {
        $destination = Join-Path $caseRoot 'bộ backup'
        $first = New-RulesBackup -Path $source -DataRoot $destination
        $second = New-RulesBackup -Path $source -DataRoot $destination
        $first.Path | Should Not Be $second.Path
        (Get-RulesFileValidation -Path $first.Path).IntegrityStatus | Should Be 'Verified'
        (Get-Item -LiteralPath $source).Length | Should Be 4
        foreach ($name in @('Rules', 'Signatures', 'Templates', 'PST', 'Report')) {
            (Test-Path -LiteralPath (Join-Path $destination $name) -PathType Container) | Should Be $true
        }
    }

    It 'rejects tampered backup bytes before they can be used' {
        $backup = New-RulesBackup -Path $source -DataRoot (Join-Path $caseRoot 'output')
        [IO.File]::WriteAllBytes($backup.Path, [byte[]](4, 3, 2, 1))
        { Get-RulesFileValidation -Path $backup.Path } | Should Throw
    }

    It 'rejects malformed manifests rather than treating them as external files' {
        [IO.File]::WriteAllText(($source + '.manifest.json'), '{bad json')
        { Get-RulesFileValidation -Path $source } | Should Throw
    }

    It 'resolves a relative data root against the PowerShell location' {
        Push-Location -LiteralPath $caseRoot
        try {
            $backup = New-RulesBackup -Path $source -DataRoot '.\backup tương đối'
            (Test-Path -LiteralPath (Join-Path $caseRoot 'backup tương đối/Rules')) | Should Be $true
            $expected = Join-Path (Join-Path $caseRoot 'backup tương đối/Rules') ([IO.Path]::GetFileName($backup.Path))
            ([IO.File]::ReadAllBytes($expected) -join ',') | Should Be '1,2,3,4'
        }
        finally { Pop-Location }
    }
}

function New-TestCollection($Items) {
    $collection = [pscustomobject]@{ Values = @($Items); Count = @($Items).Count }
    $collection | Add-Member ScriptMethod Item { param($Index) $this.Values[$Index - 1] }
    return $collection
}

function New-TestSession {
    $folder = [pscustomobject]@{
        Name = 'Khách hàng'; FolderPath = '\\PST mới\Khách hàng'; EntryID = 'folder-1'; StoreID = 'store-2'
        Folders = (New-TestCollection @())
    }
    $rootFolder = [pscustomobject]@{
        Name = 'PST mới'; FolderPath = '\\PST mới'; EntryID = 'root-2'; StoreID = 'store-2'
        Folders = (New-TestCollection @($folder))
    }
    $store = [pscustomobject]@{ DisplayName = 'PST mới'; StoreID = 'store-2'; FilePath = 'C:\Thư\mail.pst'; Root = $rootFolder; RulesList = $null }
    $store | Add-Member ScriptMethod GetRootFolder { $this.Root }
    $store | Add-Member ScriptMethod GetRules { $this.RulesList }
    $account = [pscustomobject]@{ DisplayName = 'POP mẫu'; SmtpAddress = 'user@example.com'; AccountType = 2; DeliveryStore = $store }
    $session = [pscustomobject]@{
        Accounts = (New-TestCollection @($account)); Stores = (New-TestCollection @($store)); Destination = $folder
    }
    $session | Add-Member ScriptMethod GetFolderFromID {
        param($EntryId, $StoreId)
        if ($EntryId -ne 'folder-1' -or $StoreId -ne 'store-2') { throw 'Folder unavailable' }
        $this.Destination
    }
    return $session
}

function New-TestRule($MoveAction) {
    $copy = [pscustomobject]@{ Enabled = $false; ActionType = 5; Folder = $null }
    $actions = New-TestCollection @($MoveAction, $copy)
    $actions | Add-Member NoteProperty MoveToFolder $MoveAction
    $actions | Add-Member NoteProperty CopyToFolder $copy
    return [pscustomobject]@{
        Name = 'Khách hàng'; Enabled = $true; ExecutionOrder = 1; RuleType = 0; IsLocalRule = $true
        Actions = $actions; Conditions = (New-TestCollection @()); Exceptions = (New-TestCollection @())
    }
}

Describe 'Read-only Outlook inventory and folder checks' {
    It 'preserves account and store identifiers and inventories nested Unicode folders' {
        $session = New-TestSession
        $inventory = Get-OutlookInventory -Session $session
        $inventory.Accounts.Count | Should Be 1
        $inventory.Accounts[0].DeliveryStoreId | Should Be 'store-2'
        $inventory.Stores[0].FilePath | Should Be 'C:\Thư\mail.pst'
        $inventory.Folders.Count | Should Be 2
        $inventory.Folders[1].FolderPath | Should Be '\\PST mới\Khách hàng'
    }

    It 'checks resolvable move targets without claiming that all rule behavior is verified' {
        $session = New-TestSession
        $move = [pscustomobject]@{ Enabled = $true; ActionType = 1; Folder = $session.Destination }
        $session.Stores.Item(1).RulesList = New-TestCollection @(New-TestRule $move)
        $report = Get-StoreRuleReport -Session $session -StoreId 'store-2'
        $report.Rules[0].TechnicalStatus | Should Be 'Checked'
        $report.Rules[0].MoveFolderPath | Should Be '\\PST mới\Khách hàng'
        $report.Rules[0].ManualReviewRequired | Should Be $true
    }

    It 'marks a missing enabled target as needing repair' {
        $session = New-TestSession
        $move = [pscustomobject]@{ Enabled = $true; ActionType = 1; Folder = $null }
        $session.Stores.Item(1).RulesList = New-TestCollection @(New-TestRule $move)
        $report = Get-StoreRuleReport -Session $session -StoreId 'store-2'
        $report.Rules[0].TechnicalStatus | Should Be 'NeedsRepair'
    }

    It 'marks an unresolved folder for manual review instead of trusting its display name' {
        $session = New-TestSession
        $invalid = [pscustomobject]@{ FolderPath = '\\PST mới\Khách hàng'; EntryID = 'stale-id'; StoreID = 'store-2' }
        $move = [pscustomobject]@{ Enabled = $true; ActionType = 1; Folder = $invalid }
        $session.Stores.Item(1).RulesList = New-TestCollection @(New-TestRule $move)
        $report = Get-StoreRuleReport -Session $session -StoreId 'store-2'
        $report.Rules[0].TechnicalStatus | Should Be 'ManualReview'
        $report.Rules[0].MoveStatus | Should Be 'Unverifiable'
    }

    It 'records GetRules failures without reporting an empty successful rule set' {
        $session = New-TestSession
        $session.Stores.Item(1) | Add-Member ScriptMethod GetRules { throw 'COM unavailable' } -Force
        $report = Get-StoreRuleReport -Session $session -StoreId 'store-2'
        $report.ReadSucceeded | Should Be $false
        $report.Issues.Count | Should BeGreaterThan 0
    }

    It 'continues past an unreadable rule and labels it as requiring manual review' {
        $session = New-TestSession
        $broken = [pscustomobject]@{}
        $broken | Add-Member ScriptProperty Name { throw 'Unreadable rule' }
        $move = [pscustomobject]@{ Enabled = $true; ActionType = 1; Folder = $session.Destination }
        $session.Stores.Item(1).RulesList = New-TestCollection @($broken, (New-TestRule $move))
        $report = Get-StoreRuleReport -Session $session -StoreId 'store-2'
        $report.Rules.Count | Should Be 2
        $report.Rules[0].TechnicalStatus | Should Be 'ManualReview'
        $report.Rules[1].TechnicalStatus | Should Be 'Checked'
    }

    It 'rejects an invalid account index without choosing a default store' {
        $session = New-TestSession
        $inventory = Get-OutlookInventory -Session $session
        { Select-MigrationAccount -Accounts $inventory.Accounts -AccountIndex 2 } | Should Throw
        (Select-MigrationAccount -Accounts $inventory.Accounts -AccountIndex 1).DisplayName | Should Be 'POP mẫu'
    }

    It 'preserves an account with no delivery store and refuses to substitute a PST' {
        $session = New-TestSession
        $session.Accounts.Item(1).DeliveryStore = $null
        $inventory = Get-OutlookInventory -Session $session
        $inventory.Accounts.Count | Should Be 1
        $inventory.Stores.Count | Should Be 1
        $inventory.Issues.Count | Should BeGreaterThan 0
        { Select-MigrationAccount -Accounts $inventory.Accounts -AccountIndex 1 } | Should Throw
    }

    It 'flags a rule bound to the old computer even when its folder resolves' {
        $session = New-TestSession
        $move = [pscustomobject]@{ Enabled = $true; ActionType = 1; Folder = $session.Destination }
        $rule = New-TestRule $move
        $rule.Conditions = New-TestCollection @([pscustomobject]@{ Enabled = $true; ConditionType = 28 })
        $session.Stores.Item(1).RulesList = New-TestCollection @($rule)
        $report = Get-StoreRuleReport -Session $session -StoreId 'store-2'
        $report.Rules[0].TechnicalStatus | Should Be 'NeedsRepair'
        $report.Issues[0].Message | Should Match 'OnOtherMachine'
    }

    It 'selects the requested account when POP accounts share the same delivery store' {
        $session = New-TestSession
        $second = [pscustomobject]@{ DisplayName = 'POP thứ hai'; SmtpAddress = 'second@example.com'; AccountType = 2; DeliveryStore = $session.Stores.Item(1) }
        $session.Accounts = New-TestCollection @($session.Accounts.Item(1), $second)
        $inventory = Get-OutlookInventory -Session $session
        $inventory.Accounts.Count | Should Be 2
        $selected = Select-MigrationAccount -Accounts $inventory.Accounts -AccountIndex 2
        $selected.SmtpAddress | Should Be 'second@example.com'
        $selected.DeliveryStoreId | Should Be 'store-2'
    }

    It 'keeps a folder enumeration failure visible instead of claiming a complete tree' {
        $session = New-TestSession
        $session.Stores.Item(1).Root.Folders | Add-Member ScriptMethod Item { throw 'Access denied' } -Force
        $inventory = Get-OutlookInventory -Session $session
        $inventory.Folders.Count | Should Be 1
        $inventory.Issues[0].Message | Should Match 'Access denied'
    }
}

Describe 'Migration workflows without a live mailbox' {
    BeforeEach {
        $script:caseRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $caseRoot | Out-Null
        $script:source = Join-Path $caseRoot 'quy tắc [1].rwz'
        [IO.File]::WriteAllBytes($source, [byte[]](1, 2, 3, 4))
        $script:session = New-TestSession
        $session.Stores.Item(1).RulesList = New-TestCollection @()
        Mock Connect-RunningOutlook { [pscustomobject]@{ Application = $null; Session = $script:session } }
    }

    It 'backs up successfully when Outlook is unavailable and records the reason' {
        Mock Connect-RunningOutlook { throw 'Classic Outlook is not running' }
        $result = Invoke-RulesBackup -RulesFile $source -DataRoot (Join-Path $caseRoot 'backup')
        $result.FileValidation.IntegrityStatus | Should Be 'Verified'
        $summary = Get-Content -LiteralPath (Join-Path $result.ReportDirectory 'Summary.txt') -Raw -Encoding UTF8
        $summary | Should Match 'Classic Outlook is not running'
        (Get-Content -LiteralPath (Join-Path $result.ReportDirectory 'issues.csv') -Raw -Encoding UTF8) | Should Match 'Outlook'
    }

    It 'inspects the explicitly chosen account and writes UTF8 reports without modifying source bytes' {
        $result = Invoke-RulesRestore -RulesFile $source -DataRoot (Join-Path $caseRoot 'restore') -AccountIndex 1
        $result.Stage | Should Be 'Inspect'
        $result.SelectedAccount.DeliveryStoreId | Should Be 'store-2'
        $csv = Join-Path $result.ReportDirectory 'folders.csv'
        $bytes = [IO.File]::ReadAllBytes($csv)
        ($bytes[0..2] -join ',') | Should Be '239,187,191'
        (Import-Csv -LiteralPath $csv)[1].Name | Should Be 'Khách hàng'
        ([IO.File]::ReadAllBytes($source) -join ',') | Should Be '1,2,3,4'
        # Header remains present even with zero rules; absence of rows is not a read error.
        (Get-Content -LiteralPath (Join-Path $result.ReportDirectory 'rules.csv') -Raw -Encoding UTF8) | Should Match 'TechnicalStatus'
    }

    It 'does not choose even the sole account without user input' {
        Mock Read-Host { 'HUY' }
        { Invoke-RulesRestore -RulesFile $source -DataRoot (Join-Path $caseRoot 'restore') } | Should Throw
    }

    It 'cancels before import instructions when the existing rules have not been backed up' {
        Mock Read-Host { 'HUY' }
        { Invoke-RulesRestore -RulesFile $source -DataRoot (Join-Path $caseRoot 'restore') -AccountIndex 1 -Stage GuideImport } | Should Throw
        $summaryPath = Get-ChildItem -LiteralPath (Join-Path $caseRoot 'restore/Report') -Filter Summary.txt -Recurse | Select-Object -First 1
        (Get-Content -LiteralPath $summaryPath.FullName -Raw -Encoding UTF8) | Should Match 'hủy'
    }

    It 'waits for import confirmation then reads fresh rule state for verification' {
        $script:promptNumber = 0
        Mock Read-Host {
            $script:promptNumber++
            if ($script:promptNumber -eq 1) { return 'SAO-LUU' }
            $move = [pscustomobject]@{ Enabled = $true; ActionType = 1; Folder = $script:session.Destination }
            $script:session.Stores.Item(1).RulesList = New-TestCollection @(New-TestRule $move)
            return 'DA-NHAP'
        }
        $result = Invoke-RulesRestore -RulesFile $source -DataRoot (Join-Path $caseRoot 'restore') -AccountIndex 1 -Stage GuideImport
        $result.RuleReport.Rules.Count | Should Be 1
        $result.RuleReport.Rules[0].TechnicalStatus | Should Be 'Checked'
        $before = Join-Path $result.ReportDirectory 'PreImport/rules.csv'
        @(Import-Csv -LiteralPath $before).Count | Should Be 0
    }

    It 'rejects checksum changes before connecting to Outlook' {
        $backup = New-RulesBackup -Path $source -DataRoot (Join-Path $caseRoot 'backup')
        [IO.File]::WriteAllBytes($backup.Path, [byte[]](9, 9, 9, 9))
        { Invoke-RulesRestore -RulesFile $backup.Path -DataRoot (Join-Path $caseRoot 'restore') -AccountIndex 1 } | Should Throw
        Assert-MockCalled Connect-RunningOutlook -Times 0 -Exactly -Scope It
    }

    It 'verifies the copy action independently of a disabled move action' {
        $move = [pscustomobject]@{ Enabled = $false; ActionType = 1; Folder = $null }
        $rule = New-TestRule $move
        $rule.Actions.CopyToFolder.Enabled = $true
        $rule.Actions.CopyToFolder.Folder = $session.Destination
        $session.Stores.Item(1).RulesList = New-TestCollection @($rule)
        $result = Invoke-RulesRestore -RulesFile $source -DataRoot (Join-Path $caseRoot 'restore') -AccountIndex 1 -Stage Verify
        $result.RuleReport.Rules[0].TechnicalStatus | Should Be 'Checked'
        $result.RuleReport.Rules[0].MoveStatus | Should Be 'NotEnabled'
        $result.RuleReport.Rules[0].CopyFolderPath | Should Be '\\PST mới\Khách hàng'
    }

    It 'stops if the target account changes while the user imports manually' {
        $move = [pscustomobject]@{ Enabled = $true; ActionType = 1; Folder = $session.Destination }
        $session.Stores.Item(1).RulesList = New-TestCollection @(New-TestRule $move)
        $script:promptNumber = 0
        Mock Read-Host {
            $script:promptNumber++
            if ($script:promptNumber -eq 1) { return 'SAO-LUU' }
            $script:session.Accounts.Item(1).SmtpAddress = 'different@example.com'
            return 'DA-NHAP'
        }
        { Invoke-RulesRestore -RulesFile $source -DataRoot (Join-Path $caseRoot 'restore') -AccountIndex 1 -Stage GuideImport } | Should Throw
        $summaryPath = Get-ChildItem -LiteralPath (Join-Path $caseRoot 'restore/Report') -Filter Summary.txt -Recurse | Where-Object { $_.Directory.Name -ne 'PreImport' } | Select-Object -First 1
        (Get-Content -LiteralPath $summaryPath.FullName -Raw -Encoding UTF8) | Should Match 'đã thay đổi'
        # A failed post-import check must not present the old rules as a fresh verification.
        @(Import-Csv -LiteralPath (Join-Path $summaryPath.Directory.FullName 'rules.csv')).Count | Should Be 0
    }
}

Describe 'Windows PowerShell entry scripts' {
    It 'defaults backup data to the script directory when DataRoot is omitted' {
        $caseRoot = Join-Path $TestDrive ('default-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $caseRoot | Out-Null
        $scriptDir = Join-Path $caseRoot 'bộ script'
        New-Item -ItemType Directory -Path $scriptDir | Out-Null
        $projectRoot = Split-Path $PSScriptRoot -Parent
        foreach ($name in @('Backup.ps1', 'Common.ps1')) {
            Copy-Item -LiteralPath (Join-Path $projectRoot $name) -Destination (Join-Path $scriptDir $name)
        }
        $source = Join-Path $caseRoot 'nguồn.rwz'
        [IO.File]::WriteAllBytes($source, [byte[]](1, 2, 3, 4))
        $entry = Join-Path $scriptDir 'Backup.ps1'
        & powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File $entry -RulesFile $source -SkipOutlook *> (Join-Path $caseRoot 'stdout.log')
        $LASTEXITCODE | Should Be 0
        $backups = @(Get-ChildItem -LiteralPath (Join-Path $scriptDir 'Rules') -Filter '*.rwz')
        $backups.Count | Should Be 1
        (Get-RulesFileValidation -Path $backups[0].FullName).IntegrityStatus | Should Be 'Verified'
    }

    It 'defaults restore reports to the script directory without attaching to real Outlook' {
        $caseRoot = Join-Path $TestDrive ('restore-default-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $caseRoot | Out-Null
        $scriptDir = Join-Path $caseRoot 'bộ script'
        New-Item -ItemType Directory -Path $scriptDir | Out-Null
        $projectRoot = Split-Path $PSScriptRoot -Parent
        foreach ($name in @('Restore.ps1', 'Common.ps1')) {
            Copy-Item -LiteralPath (Join-Path $projectRoot $name) -Destination (Join-Path $scriptDir $name)
        }
        # Replace only the external connection boundary in the temporary copy.
        Add-Content -LiteralPath (Join-Path $scriptDir 'Common.ps1') -Encoding UTF8 -Value "`r`nfunction Connect-RunningOutlook { throw 'TEST-OUTLOOK-UNAVAILABLE' }"
        $source = Join-Path $caseRoot 'nguồn.rwz'
        [IO.File]::WriteAllBytes($source, [byte[]](1, 2, 3, 4))
        $entry = Join-Path $scriptDir 'Restore.ps1'
        $savedPreference = $ErrorActionPreference
        try {
            $ErrorActionPreference = 'Continue'
            & powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File $entry -RulesFile $source -AccountIndex 1 *> (Join-Path $caseRoot 'stdout.log')
            $LASTEXITCODE | Should Be 1
        }
        finally { $ErrorActionPreference = $savedPreference }
        $reports = @(Get-ChildItem -LiteralPath (Join-Path $scriptDir 'Report') -Filter 'Summary.txt' -Recurse)
        $reports.Count | Should Be 1
        (Get-Content -LiteralPath $reports[0].FullName -Raw -Encoding UTF8) | Should Match 'TEST-OUTLOOK-UNAVAILABLE'
    }

    It 'runs Backup.ps1 against an external Unicode path without launching Outlook' {
        $caseRoot = Join-Path $TestDrive ('cli-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $caseRoot | Out-Null
        $source = Join-Path $caseRoot 'nguồn [1].rwz'
        [IO.File]::WriteAllBytes($source, [byte[]](1, 2, 3, 4))
        $dataRoot = Join-Path $caseRoot 'backup'
        $entry = Join-Path (Split-Path $PSScriptRoot -Parent) 'Backup.ps1'
        & powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File $entry -RulesFile $source -DataRoot $dataRoot -SkipOutlook *> (Join-Path $caseRoot 'stdout.log')
        $LASTEXITCODE | Should Be 0
        $backups = @(Get-ChildItem -LiteralPath (Join-Path $dataRoot 'Rules') -Filter '*.rwz')
        $backups.Count | Should Be 1
        (Get-RulesFileValidation -Path $backups[0].FullName).IntegrityStatus | Should Be 'Verified'
    }

    It 'returns a nonzero exit code for missing Rules input' {
        $entry = Join-Path (Split-Path $PSScriptRoot -Parent) 'Restore.ps1'
        $log = Join-Path $TestDrive 'missing-input.log'
        $missing = Join-Path $TestDrive 'not-found.rwz'
        # Native stderr is an error stream in PS5.1, so allow it while inspecting exit status.
        $savedPreference = $ErrorActionPreference
        try {
            $ErrorActionPreference = 'Continue'
            & powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File $entry -RulesFile $missing *> $log
            $LASTEXITCODE | Should Be 1
        }
        finally { $ErrorActionPreference = $savedPreference }
    }
}
