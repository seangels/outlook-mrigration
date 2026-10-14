# Dot-source Common.ps1 before this file. All COM operations run in the calling STA.
Set-StrictMode -Version 2.0

function Get-RepairSnapshotHash {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rules)
    $fields = @('StoreId', 'RuleIndex', 'Name', 'Enabled', 'ExecutionOrder', 'RuleType', 'IsLocalRule', 'EnabledActionTypes', 'EnabledConditionTypes', 'EnabledExceptionTypes', 'MoveStatus', 'MoveFolderPath', 'MoveStoreId', 'MoveEntryId', 'CopyStatus', 'CopyFolderPath', 'CopyStoreId', 'CopyEntryId')
    $normalized = @(
        foreach ($rule in $Rules) {
            $row = [ordered]@{}
            foreach ($field in $fields) { $row[$field] = [string]$rule.$field }
            [pscustomobject]$row
        }
    )
    $json = ConvertTo-Json -InputObject $normalized -Depth 6 -Compress
    $algorithm = [Security.Cryptography.SHA256]::Create()
    try { ([BitConverter]::ToString($algorithm.ComputeHash([Text.Encoding]::UTF8.GetBytes($json)))).Replace('-', '') }
    finally { $algorithm.Dispose() }
}

function Get-RepairRelativeFolderPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not $Path.StartsWith('\\')) { return $null }
    $withoutPrefix = $Path.Substring(2)
    $separator = $withoutPrefix.IndexOf('\')
    if ($separator -lt 0) { return '' }
    $withoutPrefix.Substring($separator + 1)
}

function New-RuleRepairPlan {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$CurrentRules,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$SourceRules,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Folders,
        [Parameter(Mandatory = $true)][hashtable]$TargetStoreMap
    )
    $rows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($current in $CurrentRules) {
        $sources = @($SourceRules | Where-Object { $_.Name -ceq $current.Name -and [string]$_.RuleType -eq [string]$current.RuleType })
        $duplicates = @($CurrentRules | Where-Object { $_.Name -ceq $current.Name -and [string]$_.RuleType -eq [string]$current.RuleType })
        foreach ($definition in @(@{ Name = 'MoveToFolder'; Prefix = 'Move'; Code = '1' }, @{ Name = 'CopyToFolder'; Prefix = 'Copy'; Code = '5' })) {
            $prefix = $definition.Prefix
            $status = [string]$current.($prefix + 'Status')
            if ($status -notin @('MissingTarget', 'Unverifiable') -or ([string]$current.EnabledActionTypes -split ';') -notcontains $definition.Code) { continue }
            $row = [ordered]@{
                RowId = "$($current.RuleIndex):$($definition.Name)"; RuleIndex = $current.RuleIndex; RuleName = $current.Name
                Action = $definition.Name; SourceFolderPath = ''; SourceEnabled = 'Unknown'; Apply = 'No'
                TargetStoreId = ''; TargetFolderPath = ''; MatchStatus = 'SourceRuleMissing'
            }
            if ($sources.Count -gt 1 -or $duplicates.Count -ne 1) { $row.MatchStatus = 'AmbiguousRule' }
            elseif ($sources.Count -eq 1) {
                $source = $sources[0]
                if ([string]$source.Enabled -iin @('True', 'False')) { $row.SourceEnabled = [string]$source.Enabled }
                $row.SourceFolderPath = [string]$source.($prefix + 'FolderPath')
                $oldStore = [string]$source.($prefix + 'StoreId')
                $relative = Get-RepairRelativeFolderPath $row.SourceFolderPath
                $row.MatchStatus = 'SourceFolderMissing'
                if ($oldStore -and $TargetStoreMap.ContainsKey($oldStore)) { $row.TargetStoreId = [string]$TargetStoreMap[$oldStore] }
                if ($null -ne $relative) {
                    $row.MatchStatus = 'StoreMappingMissing'
                    if ($row.TargetStoreId) {
                        $matches = @($Folders | Where-Object {
                            $_.StoreId -eq $row.TargetStoreId -and (Get-RepairRelativeFolderPath $_.FolderPath) -ieq $relative
                        })
                        $row.MatchStatus = 'TargetFolderMissing'
                        if ($matches.Count -gt 1) { $row.MatchStatus = 'AmbiguousFolder' }
                        elseif ($matches.Count -eq 1) {
                            $row.TargetFolderPath = [string]$matches[0].FolderPath
                            $row.MatchStatus = 'ExactRelativePath'; $row.Apply = 'Yes'
                        }
                    }
                }
            }
            [void]$rows.Add([pscustomobject]$row)
        }
    }
    [pscustomobject]@{ Rows = @($rows.ToArray()); SnapshotHash = (Get-RepairSnapshotHash $CurrentRules) }
}

function Save-RuleRepairPlan {
    param($Plan, [string]$Directory, [string]$RuleStoreId)
    [IO.Directory]::CreateDirectory((Resolve-MigrationDataRoot $Directory)) | Out-Null
    $path = Join-Path $Directory 'folder-mapping.csv'
    if (Test-Path -LiteralPath $path) { throw "Bảng ánh xạ đã tồn tại; tạo Plan mới thay vì ghi đè: $path" }
    Export-MigrationCsv @($Plan.Rows) @('RowId', 'RuleIndex', 'RuleName', 'Action', 'SourceFolderPath', 'SourceEnabled', 'Apply', 'TargetStoreId', 'TargetFolderPath', 'MatchStatus') $path
    [ordered]@{
        SchemaVersion = 1; PlanId = [guid]::NewGuid().ToString('N'); CreatedAt = (Get-Date).ToString('o')
        RuleStoreId = $RuleStoreId; SnapshotHash = $Plan.SnapshotHash; Rows = @($Plan.Rows)
    } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath ($path + '.plan.json') -Encoding UTF8 -ErrorAction Stop
    (Get-Item -LiteralPath $path).FullName
}

function Read-RuleRepairMapping {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Không tìm thấy bảng ánh xạ: $Path" }
    $sidecar = $Path + '.plan.json'
    if (-not (Test-Path -LiteralPath $sidecar -PathType Leaf)) { throw "Thiếu sidecar của bảng ánh xạ: $sidecar" }
    $plan = Get-Content -LiteralPath $sidecar -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
    if ($plan.SchemaVersion -ne 1 -or -not $plan.RuleStoreId -or $plan.SnapshotHash -notmatch '^[0-9A-Fa-f]{64}$') { throw 'Sidecar Plan không hợp lệ.' }
    $rows = @(Import-Csv -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop)
    if ($rows.Count -ne @($plan.Rows).Count) { throw 'Số dòng CSV đã thay đổi. Giữ nguyên các dòng và dùng Apply=No để bỏ qua.' }
    $seen = @{}
    foreach ($row in $rows) {
        if ($seen.ContainsKey($row.RowId)) { throw "RowId bị trùng: $($row.RowId)" }
        $seen[$row.RowId] = $true
        $originals = @($plan.Rows | Where-Object { $_.RowId -ceq $row.RowId })
        if ($originals.Count -ne 1) { throw "RowId không thuộc Plan: $($row.RowId)" }
        foreach ($field in @('RuleIndex', 'RuleName', 'Action', 'SourceFolderPath', 'SourceEnabled', 'MatchStatus')) {
            if ([string]$row.$field -cne [string]$originals[0].$field) { throw "Không sửa cột $field; chỉ sửa Apply, TargetStoreId, TargetFolderPath." }
        }
        if ($row.Apply -notin @('Yes', 'No')) { throw "Apply phải là Yes hoặc No: $($row.RowId)" }
        if ($row.Apply -eq 'Yes' -and (-not $row.TargetStoreId -or -not $row.TargetFolderPath)) { throw "Chưa chọn đủ kho và thư mục đích: $($row.RowId)" }
    }
    [pscustomobject]@{ Plan = $plan; Rows = $rows }
}

function Write-RepairSnapshot {
    param([string]$Directory, $Inventory, $RuleReport)
    [IO.Directory]::CreateDirectory($Directory) | Out-Null
    Export-MigrationCsv @($Inventory.Accounts) @('Index', 'DisplayName', 'SmtpAddress', 'AccountType', 'DeliveryStoreId', 'DeliveryStoreName') (Join-Path $Directory 'accounts.csv')
    Export-MigrationCsv @($Inventory.Stores) @('StoreId', 'DisplayName', 'FilePath') (Join-Path $Directory 'stores.csv')
    Export-MigrationCsv @($Inventory.Folders) @('StoreId', 'EntryId', 'FolderPath', 'Name', 'Depth') (Join-Path $Directory 'folders.csv')
    $RuleReport.Rules | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $Directory 'rules.json') -Encoding UTF8
    if (@($RuleReport.Rules).Count -gt 0) { $RuleReport.Rules | Export-Csv -LiteralPath (Join-Path $Directory 'rules.csv') -NoTypeInformation -Encoding UTF8 }
    else { Export-MigrationCsv @() @('StoreId', 'RuleIndex', 'Name', 'TechnicalStatus') (Join-Path $Directory 'rules.csv') }
    Export-MigrationCsv (@($Inventory.Issues) + @($RuleReport.Issues)) @('Area', 'Identity', 'Status', 'Message') (Join-Path $Directory 'issues.csv')
}

function Write-RepairResults {
    param([string]$Directory, [object[]]$Results, [string]$Status, [string[]]$Notes)
    Export-MigrationCsv $Results @('RowId', 'RuleIndex', 'RuleName', 'Action', 'BeforeFolderPath', 'TargetFolderPath', 'Status', 'WasEnabled', 'EnabledAfter', 'EnableNote', 'Detail') (Join-Path $Directory 'repair-results.csv')
    @(
        "OutlookMigration - Repair - $Status - $((Get-Date).ToString('o'))"
        "Số action được chọn: $($Results.Count)."
        'Chỉ sửa liên kết Move/Copy đã chọn; không chạy Rules trên email cũ, không tạo/xóa Rules hoặc thư mục.'
        'FolderVerified chỉ xác nhận liên kết và trạng thái đọc được sau Save; vẫn phải thử email mới đúng điều kiện.'
        'Rules.Save lưu cả bộ Rules của kho. Save lỗi có thể có trạng thái không xác định; script không tự retry/hoàn tác.'
        'Dùng bản .rwz đã xuất trước sửa nếu cần phục hồi thủ công. Snapshot JSON/CSV không thay thế .rwz.'
    ) + $Notes | Set-Content -LiteralPath (Join-Path $Directory 'Summary.txt') -Encoding UTF8
    Write-Host "Báo cáo sửa Rules: $Directory"
}

function Invoke-RuleRepair {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory = $true)]$Session,
        [Parameter(Mandatory = $true)][string]$MappingFile,
        [Parameter(Mandatory = $true)][string]$DataRoot,
        [string]$BackupRulesFile,
        [switch]$EnableRepairedRules
    )
    $mapping = Read-RuleRepairMapping $MappingFile
    $selected = @($mapping.Rows | Where-Object { $_.Apply -eq 'Yes' })
    if ($selected.Count -eq 0) {
        Write-Host 'Không có dòng Apply=Yes. Chưa sửa Rules.'
        return [pscustomobject]@{ Status = 'NoChanges'; Results = @(); ReportDirectory = '' }
    }
    $before = Get-StoreRuleReport -Session $Session -StoreId $mapping.Plan.RuleStoreId
    if (-not $before.ReadSucceeded -or (Get-RepairSnapshotHash $before.Rules) -ne $mapping.Plan.SnapshotHash) {
        throw 'Rules đã thay đổi hoặc không đọc được từ lúc tạo Plan. Chạy Plan mới trước khi sửa.'
    }
    $inventory = Get-OutlookInventory -Session $Session
    $work = New-Object 'System.Collections.Generic.List[object]'
    $results = New-Object 'System.Collections.Generic.List[object]'
    $storeCollection = $null; $store = $null; $rules = $null
    $directory = ''; $saveAttempted = $false; $saved = $false
    $notes = @()
    try {
        $storeCollection = $Session.Stores
        for ($i = 1; $i -le $storeCollection.Count; $i++) {
            $candidate = $storeCollection.Item($i)
            if ($candidate.StoreID -eq $mapping.Plan.RuleStoreId) { $store = $candidate; break }
            Release-OutlookReference $candidate
        }
        if ($null -eq $store) { throw 'Kho chứa Rules không còn mở.' }
        $rules = $store.GetRules()
        if ($null -eq $rules -or $rules.Count -ne $before.Rules.Count) { throw 'Số lượng Rules thay đổi; tạo Plan mới.' }

        # Resolve EVERY selected target and validate identities before any setter.
        foreach ($row in $selected) {
            $expectedRows = @($before.Rules | Where-Object { [string]$_.RuleIndex -eq $row.RuleIndex })
            if ($expectedRows.Count -ne 1) { throw "Không tìm thấy RuleIndex: $($row.RuleIndex)" }
            $expected = $expectedRows[0]
            $rule = $null; $actions = $null; $action = $null; $target = $null
            try {
                $rule = $rules.Item([int]$row.RuleIndex)
                foreach ($field in @('Name', 'Enabled', 'ExecutionOrder', 'RuleType', 'IsLocalRule')) {
                    if ([string]$rule.$field -cne [string]$expected.$field) { throw "Rule thay đổi ở $field; tạo Plan mới: $($row.RuleName)" }
                }
                if ($row.Action -notin @('MoveToFolder', 'CopyToFolder')) { throw 'Chỉ hỗ trợ action MoveToFolder/CopyToFolder.' }
                $actions = $rule.Actions
                $action = $actions.($row.Action)
                if ($null -eq $action -or -not $action.Enabled) { throw "Action không còn bật: $($row.RowId)" }
                $folders = @($inventory.Folders | Where-Object { $_.StoreId -eq $row.TargetStoreId -and $_.FolderPath -ieq $row.TargetFolderPath })
                if ($folders.Count -ne 1) { throw "Thư mục đích không có hoặc không duy nhất: $($row.TargetFolderPath)" }
                $target = $Session.GetFolderFromID($folders[0].EntryId, $folders[0].StoreId)
                if ($null -eq $target -or [string]$target.EntryID -ne [string]$folders[0].EntryId -or [string]$target.StoreID -ne [string]$folders[0].StoreId) {
                    throw "Không xác minh được định danh thư mục: $($row.TargetFolderPath)"
                }
                $prefix = 'Move'; if ($row.Action -eq 'CopyToFolder') { $prefix = 'Copy' }
                $journalRow = [pscustomobject]@{
                    RowId = $row.RowId; RuleIndex = [int]$row.RuleIndex; RuleName = $row.RuleName; Action = $row.Action
                    BeforeFolderPath = $expected.($prefix + 'FolderPath'); TargetFolderPath = [string]$target.FolderPath
                    Status = 'Planned'; WasEnabled = [bool]$rule.Enabled; EnabledAfter = [bool]$rule.Enabled; EnableNote = ''; Detail = ''
                }
                [void]$results.Add($journalRow)
                $original = @($mapping.Plan.Rows | Where-Object { $_.RowId -ceq $row.RowId })[0]
                [void]$work.Add([pscustomobject]@{
                    Rule = $rule; Actions = $actions; Action = $action; Target = $target; Journal = $journalRow
                    TargetEntryId = [string]$target.EntryID; TargetStoreId = [string]$target.StoreID; SourceEnabled = $original.SourceEnabled
                })
                # Ownership transfers to work; release all references in finally below.
                $rule = $null; $actions = $null; $action = $null; $target = $null
            }
            finally { Release-OutlookReference $target; Release-OutlookReference $action; Release-OutlookReference $actions; Release-OutlookReference $rule }
        }
        Write-Host ($results.ToArray() | Select-Object RuleName, Action, BeforeFolderPath, TargetFolderPath, WasEnabled | Format-Table -Wrap -AutoSize | Out-String)
        if (-not $PSCmdlet.ShouldProcess($store.DisplayName, "Gán $($work.Count) liên kết thư mục và lưu bộ Rules một lần")) {
            return [pscustomobject]@{ Status = 'Preview'; Results = @($results.ToArray()); ReportDirectory = '' }
        }
        if ([string]::IsNullOrWhiteSpace($BackupRulesFile)) {
            throw 'Trước khi Apply, xuất Rules hiện có trong Outlook ra .rwz riêng và truyền -BackupRulesFile. WhatIf không cần file này.'
        }
        $backup = New-RulesBackup -Path $BackupRulesFile -DataRoot $DataRoot
        $directory = New-MigrationReportDirectory $DataRoot 'RepairApply'
        Write-RepairSnapshot -Directory (Join-Path $directory 'BeforeApply') -Inventory $inventory -RuleReport $before
        $notes += "Bản sao .rwz người dùng cung cấp trước sửa: $($backup.Path); SHA256: $($backup.SHA256). Script không kiểm chứng thời điểm nội dung được xuất từ Outlook."
        Write-RepairResults $directory $results.ToArray() 'PreflightPassed' $notes

        foreach ($item in $work) {
            if ([Runtime.InteropServices.Marshal]::IsComObject($item.Action)) {
                # Folder is an object reference. A value assignment can silently
                # leave Outlook's target empty or resolve the wrong same-name folder.
                [void]$item.Action.GetType().InvokeMember('Folder', [Reflection.BindingFlags]::PutRefDispProperty, $null, $item.Action, [object[]]@($item.Target))
            }
            else { $item.Action.Folder = $item.Target }
            $assigned = $null
            try {
                $assigned = $item.Action.Folder
                if ($null -eq $assigned -or [string]$assigned.EntryID -ne $item.TargetEntryId -or [string]$assigned.StoreID -ne $item.TargetStoreId) {
                    throw "Outlook chưa nhận đúng thư mục đích; chưa Save: $($item.Journal.RuleName)"
                }
            }
            finally { Release-OutlookReference $assigned }
            $item.Journal.Status = 'Staged'
        }
        if ($EnableRepairedRules) {
            foreach ($group in @($work.ToArray() | Group-Object { $_.Journal.RuleIndex })) {
                $members = @($group.Group)
                $rule = $members[0].Rule
                $reason = ''
                if (@($members | Where-Object { $_.SourceEnabled -ine 'True' }).Count -gt 0) { $reason = 'Nguồn không xác nhận rule từng được bật; giữ nguyên Enabled.' }
                $checks = New-Object 'System.Collections.Generic.List[object]'
                $conditions = $null; $exceptions = $null
                try {
                    $conditions = $rule.Conditions; $exceptions = $rule.Exceptions
                    $conditionCheck = Get-RuleMemberSummary $conditions 'Condition' $checks ([string]$group.Name)
                    $exceptionCheck = Get-RuleMemberSummary $exceptions 'Condition' $checks ([string]$group.Name)
                    $move = Get-RuleFolderCheck $members[0].Actions 'MoveToFolder' $Session $checks ([string]$group.Name)
                    $copy = Get-RuleFolderCheck $members[0].Actions 'CopyToFolder' $Session $checks ([string]$group.Name)
                    if ($conditionCheck.OtherMachine -or -not $conditionCheck.ReadSucceeded -or -not $exceptionCheck.ReadSucceeded -or
                        $move.Status -notin @('Resolved', 'NotEnabled') -or $copy.Status -notin @('Resolved', 'NotEnabled')) {
                        $reason = 'Còn thư mục/điều kiện chưa xác minh hoặc OnOtherMachine; giữ nguyên Enabled và kiểm tra thủ công.'
                    }
                }
                catch { $reason = 'Không kiểm tra được điều kiện/thư mục: ' + $_.Exception.Message }
                finally { Release-OutlookReference $exceptions; Release-OutlookReference $conditions }
                if (-not $reason) { $rule.Enabled = $true; $reason = 'Đã yêu cầu bật lại rule từng được bật ở máy cũ.' }
                foreach ($item in $members) { $item.Journal.EnabledAfter = [bool]$rule.Enabled; $item.Journal.EnableNote = $reason }
            }
        }
        Write-RepairResults $directory $results.ToArray() 'StagedNotSaved' $notes
        # False hides Outlook's progress dialog; it is NOT a dry-run.
        $saveAttempted = $true
        $rules.Save($false)
        $saved = $true
        Release-OutlookReference $rules; $rules = $null
        $after = Get-StoreRuleReport -Session $Session -StoreId $mapping.Plan.RuleStoreId
        Write-RepairSnapshot -Directory (Join-Path $directory 'AfterApply') -Inventory $inventory -RuleReport $after
        if (-not $after.ReadSucceeded -or $after.Rules.Count -ne $before.Rules.Count) { throw 'Đã Save nhưng không xác minh được đầy đủ bộ Rules sau lưu.' }
        $verificationFailed = $false
        foreach ($item in $work) {
            $actuals = @($after.Rules | Where-Object { $_.RuleIndex -eq $item.Journal.RuleIndex })
            $prefix = 'Move'; if ($item.Journal.Action -eq 'CopyToFolder') { $prefix = 'Copy' }
            if ($actuals.Count -eq 1 -and $actuals[0].Enabled -is [bool] -and $actuals[0].($prefix + 'Status') -eq 'Resolved' -and
                $actuals[0].($prefix + 'StoreId') -eq $item.TargetStoreId -and $actuals[0].($prefix + 'EntryId') -eq $item.TargetEntryId -and
                [bool]$actuals[0].Enabled -eq $item.Journal.EnabledAfter) {
                $item.Journal.Status = 'FolderVerified'
            }
            else { $verificationFailed = $true; $item.Journal.Status = 'VerificationFailed'; $item.Journal.Detail = 'Đã Save nhưng đích/Enabled sau lưu không khớp; kiểm tra Outlook trước khi thử lại.' }
        }
        # Check readable metadata for every rule, including unselected rules.
        foreach ($old in $before.Rules) {
            $actual = @($after.Rules | Where-Object { $_.RuleIndex -eq $old.RuleIndex })[0]
            $changed = @($work | Where-Object { $_.Journal.RuleIndex -eq $old.RuleIndex })
            foreach ($field in @('Name', 'ExecutionOrder', 'RuleType', 'EnabledActionTypes', 'EnabledConditionTypes', 'EnabledExceptionTypes')) {
                if ([string]$old.$field -cne [string]$actual.$field) { $verificationFailed = $true; $notes += "Thuộc tính $field thay đổi sau Save ở rule $($old.Name); cần kiểm tra thủ công." }
            }
            if ($changed.Count -eq 0 -and (Get-RepairSnapshotHash @($old)) -ne (Get-RepairSnapshotHash @($actual))) {
                $verificationFailed = $true; $notes += "Rule ngoài bảng sửa có thông tin thay đổi: $($old.Name)."
            }
        }
        if ($verificationFailed) { throw 'Đã Save nhưng có khác biệt khi kiểm tra sau lưu. Xem BeforeApply, AfterApply và repair-results.csv.' }
        Write-RepairResults $directory $results.ToArray() 'SavedVerified' ($notes + @('Liên kết đã được đọc lại và xác minh. Tiếp tục thử email mới; xem AfterApply/issues.csv cho các lỗi ngoài liên kết vừa sửa.'))
        [pscustomobject]@{ Status = 'SavedVerified'; Results = @($results.ToArray()); ReportDirectory = $directory }
    }
    catch {
        $failure = $_
        $status = 'NotSaved'
        if ($saveAttempted -and -not $saved) { $status = 'SaveFailedUnknown' }
        elseif ($saved) { $status = 'VerificationFailed' }
        foreach ($result in $results) { $result.Status = $status; $result.Detail = $failure.Exception.Message }
        if ($directory) {
            try {
                $fresh = Get-StoreRuleReport -Session $Session -StoreId $mapping.Plan.RuleStoreId
                Write-RepairSnapshot -Directory (Join-Path $directory 'AfterApply') -Inventory $inventory -RuleReport $fresh
                Write-RepairResults $directory $results.ToArray() $status ($notes + @($failure.Exception.Message, 'Không có lần Save thử lại. Không dùng lại Plan này nếu trạng thái Rules đã thay đổi.'))
            }
            catch { Write-Warning ('Không ghi được đủ báo cáo lỗi: ' + $_.Exception.Message) }
        }
        throw $failure
    }
    finally {
        foreach ($item in $work) {
            Release-OutlookReference $item.Target; Release-OutlookReference $item.Action
            Release-OutlookReference $item.Actions; Release-OutlookReference $item.Rule
        }
        Release-OutlookReference $rules; Release-OutlookReference $store; Release-OutlookReference $storeCollection
    }
}

function Import-RepairSourceReport {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (Test-Path -LiteralPath $Path -PathType Container) { $Path = Join-Path $Path 'rules.csv' }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Không tìm thấy rules.csv máy cũ: $Path" }
    $file = Get-Item -LiteralPath $Path
    if ($file.Extension -ine '.csv') { throw 'SourceReport phải là thư mục Report chứa rules.csv hoặc file CSV đó.' }
    $header = Get-Content -LiteralPath $file.FullName -Encoding UTF8 -TotalCount 1
    $headerSample = @(ConvertFrom-Csv -InputObject ($header + "`r`n" + '"schema-probe"'))
    if ($headerSample.Count -ne 1) { throw 'CSV nguồn thiếu header.' }
    foreach ($field in @('StoreId', 'RuleIndex', 'Name', 'RuleType', 'Enabled', 'MoveFolderPath', 'MoveStoreId', 'CopyFolderPath', 'CopyStoreId')) {
        if ($headerSample[0].PSObject.Properties.Name -notcontains $field) { throw "CSV nguồn không phải báo cáo Rules hoặc thiếu cột $field." }
    }
    $rows = @(Import-Csv -LiteralPath $file.FullName -Encoding UTF8)
    $accounts = @(); $stores = @()
    foreach ($name in @('accounts', 'stores')) {
        $related = Join-Path $file.DirectoryName ($name + '.csv')
        if (Test-Path -LiteralPath $related -PathType Leaf) {
            $items = @(Import-Csv -LiteralPath $related -Encoding UTF8)
            if ($name -eq 'accounts') { $accounts = $items } else { $stores = $items }
        }
    }
    [pscustomobject]@{ Path = $file.FullName; Rules = $rows; Accounts = $accounts; Stores = $stores }
}

function Select-RepairSourceStore {
    param($Source, [string]$SourceStoreId, [string]$SmtpAddress)
    $ids = @($Source.Rules | Where-Object { $_.StoreId } | Select-Object -ExpandProperty StoreId -Unique)
    if ($SourceStoreId) {
        if ($ids -notcontains $SourceStoreId) { throw 'SourceStoreId không có trong báo cáo Rules máy cũ.' }
        return $SourceStoreId
    }
    if ($ids.Count -eq 0) {
        Write-Warning 'Báo cáo nguồn không có Rules đọc được. Bảng sẽ cần ánh xạ thủ công; lấy lại Report khi Outlook máy cũ đang mở nếu có thể.'
        return ''
    }
    $matched = @($Source.Accounts | Where-Object { $SmtpAddress -and $_.SmtpAddress -ieq $SmtpAddress -and $_.DeliveryStoreId -in $ids } | Select-Object -ExpandProperty DeliveryStoreId -Unique)
    if ($matched.Count -eq 1) { return $matched[0] }
    if ($ids.Count -eq 1) { return $ids[0] }
    Write-Host 'Báo cáo máy cũ chứa Rules từ nhiều kho. Chọn kho chứa bộ Rules cần chuyển:'
    for ($i = 0; $i -lt $ids.Count; $i++) {
        $details = @($Source.Stores | Where-Object { $_.StoreId -eq $ids[$i] })
        $label = $ids[$i]
        if ($details.Count -gt 0) { $label = "$($details[0].DisplayName) / $($details[0].FilePath) / $($ids[$i])" }
        Write-Host ("[{0}] {1}" -f ($i + 1), $label)
    }
    $answer = Read-Host 'Index kho nguồn (hoặc HUY)'
    $index = 0
    if (-not [int]::TryParse($answer, [ref]$index) -or $index -lt 1 -or $index -gt $ids.Count) { throw 'Đã hủy hoặc kho nguồn không hợp lệ.' }
    $ids[$index - 1]
}

function Select-RepairTargetStores {
    param([object[]]$SourceRules, [object[]]$Stores, [string]$TargetStoreId)
    $map = @{}
    $sourceIds = @(
        foreach ($rule in $SourceRules) {
            if ($rule.MoveStoreId -and $rule.MoveFolderPath) { [string]$rule.MoveStoreId }
            if ($rule.CopyStoreId -and $rule.CopyFolderPath) { [string]$rule.CopyStoreId }
        }
    ) | Select-Object -Unique
    if ($TargetStoreId -and @($Stores | Where-Object { $_.StoreId -eq $TargetStoreId }).Count -ne 1) { throw 'TargetStoreId không có hoặc không duy nhất trong Outlook hiện tại.' }
    foreach ($oldStore in $sourceIds) {
        if ($TargetStoreId) { $map[$oldStore] = $TargetStoreId; continue }
        $examples = @($SourceRules | Where-Object { $_.MoveStoreId -eq $oldStore -or $_.CopyStoreId -eq $oldStore } | Select-Object -First 3)
        Write-Host "Chọn kho trên máy mới thay cho kho thư mục cũ: $oldStore"
        foreach ($example in $examples) { Write-Host ("  {0} | {1}" -f $example.MoveFolderPath, $example.CopyFolderPath) }
        for ($i = 0; $i -lt $Stores.Count; $i++) { Write-Host ("[{0}] {1} | {2}" -f ($i + 1), $Stores[$i].DisplayName, $Stores[$i].FilePath) }
        $answer = Read-Host 'Index kho thư mục đích (0 để ánh xạ sau trong CSV; HUY để dừng)'
        $index = -1
        if (-not [int]::TryParse($answer, [ref]$index) -or $index -lt 0 -or $index -gt $Stores.Count) { throw 'Đã hủy hoặc kho đích không hợp lệ.' }
        if ($index -gt 0) { $map[$oldStore] = [string]$Stores[$index - 1].StoreId }
    }
    $map
}

function Invoke-RuleRepairPlan {
    param($Session, [string]$SourceReport, [string]$DataRoot, [int]$AccountIndex = 0, [string]$SourceStoreId, [string]$TargetStoreId)
    $source = Import-RepairSourceReport $SourceReport
    $inventory = Get-OutlookInventory -Session $Session
    Show-MigrationInventory $inventory
    $account = Select-MigrationAccount -Accounts $inventory.Accounts -AccountIndex $AccountIndex
    $oldStore = Select-RepairSourceStore -Source $source -SourceStoreId $SourceStoreId -SmtpAddress $account.SmtpAddress
    $sourceRules = @($source.Rules | Where-Object { $_.StoreId -eq $oldStore })
    $current = Get-StoreRuleReport -Session $Session -StoreId $account.DeliveryStoreId
    if (-not $current.ReadSucceeded) { throw 'Không đọc được bộ Rules đích; kiểm tra Outlook trước khi lập bảng sửa.' }
    $targetMap = Select-RepairTargetStores -SourceRules $sourceRules -Stores $inventory.Stores -TargetStoreId $TargetStoreId
    $plan = New-RuleRepairPlan -CurrentRules $current.Rules -SourceRules $sourceRules -Folders $inventory.Folders -TargetStoreMap $targetMap
    $directory = New-MigrationReportDirectory $DataRoot 'RepairPlan'
    Write-RepairSnapshot -Directory $directory -Inventory $inventory -RuleReport $current
    $mappingFile = Save-RuleRepairPlan -Plan $plan -Directory $directory -RuleStoreId $account.DeliveryStoreId
    $sourceHash = (Get-FileHash -LiteralPath $source.Path -Algorithm SHA256).Hash
    @(
        "OutlookMigration - RepairPlan - $((Get-Date).ToString('o'))"
        "Nguồn: $($source.Path); SHA256: $sourceHash; kho Rules nguồn: $oldStore"
        "Đích Rules: $($account.DisplayName) <$($account.SmtpAddress)> / $($account.DeliveryStoreName)"
        "Số action thiếu/chưa xác minh liên kết có thể lập bảng: $($plan.Rows.Count). Gợi ý khớp: $(@($plan.Rows | Where-Object { $_.Apply -eq 'Yes' }).Count)."
        'Plan chỉ đọc Outlook. Bảng chỉ gồm action Move/Copy đang bật có tham chiếu thiếu/chưa xác minh; liên kết đang đọc được giữ nguyên.'
        'Khớp bằng Name+RuleType duy nhất và toàn bộ đường dẫn con trong kho bạn chọn. Kiểm tra lại các gợi ý trước Apply.'
        'Sửa Apply (Yes/No), TargetStoreId và TargetFolderPath; sao chép đích từ folders.csv. Dòng mơ hồ/thiếu nguồn mặc định No.'
        'Giữ nguyên các cột khác và sidecar .plan.json. Nếu Rules thay đổi, tạo Plan mới.'
        'Trước Apply: Export Rules hiện có ra .rwz riêng. Thử -WhatIf trước khi lưu.'
        'Apply giữ nguyên trạng thái bật/tắt mặc định. -EnableRepairedRules chỉ bật lại rule đã bật ở nguồn và đủ điều kiện kiểm tra.'
        'OnOtherMachine và các điều kiện/hành động khác cần kiểm tra thủ công; không tự sửa.'
        "Bảng ánh xạ: $mappingFile"
    ) | Set-Content -LiteralPath (Join-Path $directory 'Summary.txt') -Encoding UTF8
    Write-Host ($plan.Rows | Select-Object RuleName, Action, SourceFolderPath, TargetFolderPath, Apply, MatchStatus | Format-Table -Wrap -AutoSize | Out-String)
    Write-Host "Bảng sửa: $mappingFile"
    Write-Host 'Xem và sửa bảng trước khi Apply. Chưa có thay đổi Rules.'
    [pscustomobject]@{ MappingFile = $mappingFile; ReportDirectory = $directory; Plan = $plan }
}
