# Shared helpers for Windows PowerShell 5.1. Dot-source from the entry scripts.
Set-StrictMode -Version 2.0

function Resolve-MigrationDataRoot {
    param([Parameter(Mandatory = $true)][string]$DataRoot)
    $provider = $null
    $drive = $null
    $path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DataRoot, [ref]$provider, [ref]$drive)
    if ($provider.Name -ne 'FileSystem') { throw 'DataRoot phải là thư mục trên hệ thống tập tin.' }
    $path
}

function Initialize-MigrationLayout {
    param([Parameter(Mandatory = $true)][string]$DataRoot)
    $DataRoot = Resolve-MigrationDataRoot $DataRoot
    foreach ($name in @('', 'Rules', 'Signatures', 'Templates', 'PST', 'Report')) {
        $path = $DataRoot
        if ($name) { $path = Join-Path $DataRoot $name }
        if (-not (Test-Path -LiteralPath $path -PathType Container)) {
            [IO.Directory]::CreateDirectory($path) | Out-Null
        }
    }
}

function Get-RulesFileValidation {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Không tìm thấy file Rules: $Path"
    }
    $file = Get-Item -LiteralPath $Path -ErrorAction Stop
    if ($file.PSProvider.Name -ne 'FileSystem' -or $file.Extension -ine '.rwz') {
        throw 'Nguồn Rules phải là file .rwz trên hệ thống tập tin.'
    }
    if ($file.Length -eq 0) { throw 'File .rwz rỗng; hãy xuất lại từ Outlook Classic.' }
    $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256 -ErrorAction Stop).Hash
    $manifestPath = $file.FullName + '.manifest.json'
    $status = 'Unverified'
    if (Test-Path -LiteralPath $manifestPath) {
        try {
            $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($manifest.SchemaVersion -ne 1 -or $manifest.FileName -cne $file.Name -or
                [long]$manifest.Length -ne $file.Length -or
                [string]$manifest.SHA256 -notmatch '^[0-9a-fA-F]{64}$' -or $manifest.SHA256 -ine $hash) {
                throw 'Tên file, kích thước hoặc SHA-256 không khớp manifest.'
            }
            $status = 'Verified'
        }
        catch { throw "Backup không hợp lệ ($manifestPath): $($_.Exception.Message)" }
    }
    [pscustomobject]@{
        Path = $file.FullName; Length = $file.Length; SHA256 = $hash
        IntegrityStatus = $status; ManifestPath = $manifestPath
    }
}

function New-RulesBackup {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$DataRoot
    )
    $source = Get-RulesFileValidation -Path $Path
    $DataRoot = Resolve-MigrationDataRoot $DataRoot
    Initialize-MigrationLayout -DataRoot $DataRoot
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
    $name = 'Rules-{0}-{1}.rwz' -f $stamp, ([guid]::NewGuid().ToString('N').Substring(0, 8))
    $target = Join-Path (Join-Path $DataRoot 'Rules') $name
    # Copy without overwrite even if a filename collision occurs.
    [IO.File]::Copy($source.Path, $target, $false)
    $copy = Get-RulesFileValidation -Path $target
    if ($copy.Length -ne $source.Length -or $copy.SHA256 -ne $source.SHA256) {
        throw "File nguồn thay đổi trong khi copy; không dùng bản sao chưa có manifest: $target"
    }
    $manifest = [ordered]@{
        SchemaVersion = 1; FileName = $name; Length = $copy.Length; SHA256 = $copy.SHA256
        BackedUpAt = (Get-Date).ToString('o'); SourcePath = $source.Path
    }
    $manifest | ConvertTo-Json | Set-Content -LiteralPath ($target + '.manifest.json') -Encoding UTF8 -ErrorAction Stop
    Get-RulesFileValidation -Path $target
}

function Release-OutlookReference {
    param($Reference)
    if ($null -ne $Reference -and [Runtime.InteropServices.Marshal]::IsComObject($Reference)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Reference)
    }
}

function Connect-RunningOutlook {
    if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        throw 'Outlook COM cần STA. Chạy bằng powershell.exe -NoProfile -STA -File ...'
    }
    $application = $null
    $session = $null
    try {
        $application = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
        $session = $application.GetNamespace('MAPI')
        if ($session.Accounts.Count -eq 0) {
            throw 'Profile chưa có tài khoản; hãy cấu hình và mở Outlook Classic trước.'
        }
        [pscustomobject]@{ Application = $application; Session = $session }
    }
    catch {
        Release-OutlookReference $session
        Release-OutlookReference $application
        throw "Không kết nối được Outlook Classic đang mở. Mở Outlook và profile cần dùng, cùng tài khoản/quyền Windows với PowerShell, rồi thử lại. Chi tiết: $($_.Exception.Message)"
    }
}

function Add-MigrationIssue {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()]$Issues,
        [string]$Area, [string]$Identity, [string]$Message, [string]$Status = 'ManualReview'
    )
    [void]$Issues.Add([pscustomobject]@{ Area = $Area; Identity = $Identity; Status = $Status; Message = $Message })
}

function Read-OutlookProperty {
    param($InputObject, [string]$Name, $Issues, [string]$Area, [string]$Identity)
    try {
        if ($null -eq $InputObject) { throw 'COM trả về đối tượng rỗng.' }
        $InputObject.$Name
    }
    catch { Add-MigrationIssue $Issues $Area $Identity ("Không đọc được ${Name}: " + $_.Exception.Message) }
}

function Add-OutlookFolderRows {
    param($Folder, [string]$StoreId, $Rows, $Issues, $Visited, [int]$Depth = 0)
    if ($Depth -gt 128) {
        Add-MigrationIssue $Issues 'Folders' $StoreId 'Cây thư mục vượt quá 128 cấp; kiểm kê chưa đầy đủ.'
        return
    }
    $entryId = Read-OutlookProperty $Folder 'EntryID' $Issues 'Folders' $StoreId
    $path = Read-OutlookProperty $Folder 'FolderPath' $Issues 'Folders' $StoreId
    if ($entryId) {
        $key = $StoreId + '|' + $entryId
        if ($Visited.ContainsKey($key)) { return }
        $Visited[$key] = $true
    }
    [void]$Rows.Add([pscustomobject]@{
        StoreId = $StoreId; EntryId = $entryId; FolderPath = $path
        Name = (Read-OutlookProperty $Folder 'Name' $Issues 'Folders' $path); Depth = $Depth
    })
    $children = $null
    try {
        $children = $Folder.Folders
        for ($i = 1; $i -le $children.Count; $i++) {
            $child = $null
            try {
                $child = $children.Item($i)
                Add-OutlookFolderRows $child $StoreId $Rows $Issues $Visited ($Depth + 1)
            }
            catch { Add-MigrationIssue $Issues 'Folders' $path ("Không đọc được thư mục con ${i}: " + $_.Exception.Message) }
            finally { Release-OutlookReference $child }
        }
    }
    catch { Add-MigrationIssue $Issues 'Folders' $path $_.Exception.Message }
    finally { Release-OutlookReference $children }
}

function Get-OutlookInventory {
    param([Parameter(Mandatory = $true)]$Session)
    $accounts = New-Object 'System.Collections.Generic.List[object]'
    $stores = New-Object 'System.Collections.Generic.List[object]'
    $folders = New-Object 'System.Collections.Generic.List[object]'
    $issues = New-Object 'System.Collections.Generic.List[object]'
    $accountCollection = $null
    $storeCollection = $null
    try {
        $accountCollection = $Session.Accounts
        for ($i = 1; $i -le $accountCollection.Count; $i++) {
            $account = $null
            $delivery = $null
            try {
                $account = $accountCollection.Item($i)
                $name = Read-OutlookProperty $account 'DisplayName' $issues 'Accounts' "$i"
                $smtp = Read-OutlookProperty $account 'SmtpAddress' $issues 'Accounts' "$i"
                $type = Read-OutlookProperty $account 'AccountType' $issues 'Accounts' "$i"
                $deliveryId = ''
                $deliveryName = ''
                try {
                    $delivery = $account.DeliveryStore
                    if ($null -eq $delivery) { throw 'Tài khoản không có DeliveryStore.' }
                    $deliveryId = [string]$delivery.StoreID
                    $deliveryName = [string]$delivery.DisplayName
                }
                catch { Add-MigrationIssue $issues 'Accounts' "$i" $_.Exception.Message }
                $types = @{ 0 = 'Exchange'; 1 = 'IMAP'; 2 = 'POP3'; 3 = 'HTTP'; 4 = 'EAS'; 5 = 'Other' }
                $typeName = 'Unknown'
                if ($null -ne $type -and $types.ContainsKey([int]$type)) { $typeName = $types[[int]$type] }
                [void]$accounts.Add([pscustomobject]@{
                    Index = $i; DisplayName = $name; SmtpAddress = $smtp; AccountType = $typeName
                    DeliveryStoreId = $deliveryId; DeliveryStoreName = $deliveryName
                })
            }
            catch { Add-MigrationIssue $issues 'Accounts' "$i" $_.Exception.Message }
            finally { Release-OutlookReference $delivery; Release-OutlookReference $account }
        }
    }
    catch { Add-MigrationIssue $issues 'Accounts' 'Session' $_.Exception.Message }
    finally { Release-OutlookReference $accountCollection }
    try {
        $storeCollection = $Session.Stores
        $visited = @{}
        for ($i = 1; $i -le $storeCollection.Count; $i++) {
            $store = $null
            $root = $null
            try {
                $store = $storeCollection.Item($i)
                $id = Read-OutlookProperty $store 'StoreID' $issues 'Stores' "$i"
                $name = Read-OutlookProperty $store 'DisplayName' $issues 'Stores' "$i"
                [void]$stores.Add([pscustomobject]@{
                    StoreId = $id; DisplayName = $name
                    FilePath = (Read-OutlookProperty $store 'FilePath' $issues 'Stores' "$i")
                })
                $root = $store.GetRootFolder()
                Add-OutlookFolderRows $root $id $folders $issues $visited
            }
            catch { Add-MigrationIssue $issues 'Stores' "$i" $_.Exception.Message }
            finally { Release-OutlookReference $root; Release-OutlookReference $store }
        }
    }
    catch { Add-MigrationIssue $issues 'Stores' 'Session' $_.Exception.Message }
    finally { Release-OutlookReference $storeCollection }
    [pscustomobject]@{ Accounts = @($accounts.ToArray()); Stores = @($stores.ToArray()); Folders = @($folders.ToArray()); Issues = @($issues.ToArray()) }
}

function Select-MigrationAccount {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Accounts, [int]$AccountIndex = 0)
    if ($Accounts.Count -eq 0) { throw 'Không có tài khoản đọc được. Xem issues.csv, mở đúng profile rồi chạy lại.' }
    if ($AccountIndex -eq 0) {
        $answer = Read-Host 'Nhập Index tài khoản đích ở bảng trên (hoặc HUY)'
        $parsed = 0
        if (-not [int]::TryParse($answer, [ref]$parsed) -or $parsed -le 0) {
            throw 'Đã hủy hoặc Index không hợp lệ; chưa nhập Rules.'
        }
        $AccountIndex = $parsed
    }
    $matches = @($Accounts | Where-Object { $_.Index -eq $AccountIndex })
    if ($matches.Count -ne 1) { throw "AccountIndex $AccountIndex không có trong danh sách tài khoản hiện tại." }
    if ([string]::IsNullOrWhiteSpace($matches[0].DeliveryStoreId)) {
        throw 'Tài khoản này không có DeliveryStore đọc được; không tự chọn PST thay thế.'
    }
    $matches[0]
}

function Get-RuleMemberSummary {
    param($Collection, [ValidateSet('Action', 'Condition')][string]$Kind, $Issues, [string]$Identity)
    $enabledTypes = New-Object 'System.Collections.Generic.List[string]'
    $readSucceeded = $true
    try {
        if ($null -eq $Collection) { throw 'Không có collection.' }
        for ($i = 1; $i -le $Collection.Count; $i++) {
            $member = $null
            try {
                $member = $Collection.Item($i)
                if ($member.Enabled) {
                    if ($Kind -eq 'Action') { $type = $member.ActionType }
                    else { $type = $member.ConditionType }
                    [void]$enabledTypes.Add([string]$type)
                }
            }
            catch {
                $readSucceeded = $false
                Add-MigrationIssue $Issues $Kind $Identity ("Không đọc được mục ${i}: " + $_.Exception.Message)
            }
            finally { Release-OutlookReference $member }
        }
    }
    catch { $readSucceeded = $false; Add-MigrationIssue $Issues $Kind $Identity $_.Exception.Message }
    [pscustomobject]@{ Types = ($enabledTypes -join ';'); OtherMachine = $enabledTypes.Contains('28'); ReadSucceeded = $readSucceeded }
}

function Get-RuleFolderCheck {
    param($Actions, [string]$ActionName, $Session, $Issues, [string]$Identity)
    $action = $null
    $folder = $null
    $resolved = $null
    $result = [ordered]@{ Status = 'Unverifiable'; FolderPath = ''; EntryId = ''; StoreId = ''; Detail = '' }
    try {
        $action = $Actions.$ActionName
        if ($null -eq $action) { throw "Không đọc được action $ActionName." }
        if (-not $action.Enabled) {
            $result.Status = 'NotEnabled'
        }
        else {
            $folder = $action.Folder
            if ($null -eq $folder) {
                $result.Status = 'MissingTarget'
                $result.Detail = 'Action đang bật nhưng Folder rỗng; chọn lại thư mục đích trong Outlook.'
            }
            else {
                $result.EntryId = [string]$folder.EntryID
                $result.StoreId = [string]$folder.StoreID
                $result.FolderPath = [string]$folder.FolderPath
                if (-not $result.EntryId -or -not $result.StoreId) {
                    $result.Status = 'MissingTarget'
                    $result.Detail = 'Thiếu EntryID/StoreID của thư mục đích.'
                }
                else {
                    $resolved = $Session.GetFolderFromID($result.EntryId, $result.StoreId)
                    if ($null -eq $resolved) { throw 'Không resolve được thư mục đích.' }
                    $result.FolderPath = [string]$resolved.FolderPath
                    $result.Status = 'Resolved'
                }
            }
        }
    }
    catch { $result.Detail = $_.Exception.Message }
    finally { Release-OutlookReference $resolved; Release-OutlookReference $folder; Release-OutlookReference $action }
    if ($result.Status -eq 'MissingTarget') { Add-MigrationIssue $Issues $ActionName $Identity $result.Detail 'NeedsRepair' }
    elseif ($result.Status -eq 'Unverifiable') { Add-MigrationIssue $Issues $ActionName $Identity $result.Detail }
    [pscustomobject]$result
}

function Get-StoreRuleReport {
    param([Parameter(Mandatory = $true)]$Session, [Parameter(Mandatory = $true)][string]$StoreId)
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $issues = New-Object 'System.Collections.Generic.List[object]'
    $storeCollection = $null
    $store = $null
    $rules = $null
    $success = $false
    try {
        $storeCollection = $Session.Stores
        for ($s = 1; $s -le $storeCollection.Count; $s++) {
            $candidate = $storeCollection.Item($s)
            if ($candidate.StoreID -eq $StoreId) { $store = $candidate; break }
            Release-OutlookReference $candidate
        }
        if ($null -eq $store) { throw "Kho dữ liệu không còn mở: $StoreId" }
        $rules = $store.GetRules()
        if ($null -eq $rules) { throw 'GetRules trả về rỗng; không thể xác nhận số lượng rule.' }
        $count = $rules.Count
        $success = $true
        for ($i = 1; $i -le $count; $i++) {
            $rule = $null
            $actions = $null
            $conditions = $null
            $exceptions = $null
            $row = [ordered]@{
                StoreId = $StoreId; RuleIndex = $i; Name = ''; Enabled = $null; ExecutionOrder = $null
                RuleType = $null; IsLocalRule = $null; EnabledActionTypes = ''; EnabledConditionTypes = ''; EnabledExceptionTypes = ''
                MoveStatus = 'Unverifiable'; MoveFolderPath = ''; MoveStoreId = ''; MoveEntryId = ''
                CopyStatus = 'Unverifiable'; CopyFolderPath = ''; CopyStoreId = ''; CopyEntryId = ''
                TechnicalStatus = 'ManualReview'; ManualReviewRequired = $true
                Notes = 'Luôn đối chiếu điều kiện, ngoại lệ, account, hành động khác và đúng thư mục mong muốn trong Outlook; cần email thử.'
            }
            $before = $issues.Count
            try {
                $rule = $rules.Item($i)
                foreach ($property in @('Name', 'Enabled', 'ExecutionOrder', 'RuleType', 'IsLocalRule')) {
                    $row[$property] = Read-OutlookProperty $rule $property $issues 'Rules' "$i"
                }
                $actions = $rule.Actions
                $conditions = $rule.Conditions
                $exceptions = $rule.Exceptions
                $actionSummary = Get-RuleMemberSummary $actions 'Action' $issues "$i"
                $conditionSummary = Get-RuleMemberSummary $conditions 'Condition' $issues "$i"
                $exceptionSummary = Get-RuleMemberSummary $exceptions 'Condition' $issues "${i}:Exceptions"
                $row.EnabledActionTypes = $actionSummary.Types
                $row.EnabledConditionTypes = $conditionSummary.Types
                $row.EnabledExceptionTypes = $exceptionSummary.Types
                $move = Get-RuleFolderCheck $actions 'MoveToFolder' $Session $issues "$i"
                $copy = Get-RuleFolderCheck $actions 'CopyToFolder' $Session $issues "$i"
                $row.MoveStatus = $move.Status; $row.MoveFolderPath = $move.FolderPath; $row.MoveStoreId = $move.StoreId; $row.MoveEntryId = $move.EntryId
                $row.CopyStatus = $copy.Status; $row.CopyFolderPath = $copy.FolderPath; $row.CopyStoreId = $copy.StoreId; $row.CopyEntryId = $copy.EntryId
                if ($move.Status -eq 'MissingTarget' -or $copy.Status -eq 'MissingTarget' -or $conditionSummary.OtherMachine) {
                    $row.TechnicalStatus = 'NeedsRepair'
                    if ($conditionSummary.OtherMachine) {
                        Add-MigrationIssue $issues 'Conditions' "$i" 'Rule có điều kiện OnOtherMachine (28); kiểm tra/sửa ràng buộc máy trong Outlook.' 'NeedsRepair'
                    }
                }
                elseif ($issues.Count -eq $before -and ($move.Status -eq 'Resolved' -or $copy.Status -eq 'Resolved')) {
                    $row.TechnicalStatus = 'Checked'
                }
                if ($row.Enabled -eq $false) { $row.Notes += ' Rule đang tắt; script không tự bật.' }
            }
            catch { Add-MigrationIssue $issues 'Rules' "$i" $_.Exception.Message }
            finally {
                Release-OutlookReference $exceptions; Release-OutlookReference $conditions
                Release-OutlookReference $actions; Release-OutlookReference $rule
            }
            [void]$rows.Add([pscustomobject]$row)
        }
    }
    catch { Add-MigrationIssue $issues 'Rules' $StoreId $_.Exception.Message }
    finally { Release-OutlookReference $rules; Release-OutlookReference $store; Release-OutlookReference $storeCollection }
    [pscustomobject]@{ Rules = @($rows.ToArray()); Issues = @($issues.ToArray()); ReadSucceeded = $success }
}

function New-MigrationReportDirectory {
    param([string]$DataRoot, [string]$Stage)
    $DataRoot = Resolve-MigrationDataRoot $DataRoot
    Initialize-MigrationLayout $DataRoot
    $name = '{0}-{1}-{2}' -f (Get-Date -Format 'yyyyMMdd-HHmmss-fff'), $Stage, ([guid]::NewGuid().ToString('N').Substring(0, 8))
    $path = Join-Path (Join-Path $DataRoot 'Report') $name
    [IO.Directory]::CreateDirectory($path) | Out-Null
    $path
}

function Export-MigrationCsv {
    param([object[]]$Rows, [string[]]$Columns, [string]$Path)
    if ($Rows.Count -gt 0) {
        $Rows | Select-Object -Property $Columns | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    }
    else {
        # Keep empty reports useful to CSV readers by retaining the header.
        $header = ($Columns | ForEach-Object { '"' + ($_ -replace '"', '""') + '"' }) -join ','
        Set-Content -LiteralPath $Path -Value $header -Encoding UTF8 -ErrorAction Stop
    }
}

function Write-MigrationReport {
    param(
        [Parameter(Mandatory = $true)][string]$Directory,
        $FileValidation, $Inventory, $RuleReport, $SelectedAccount,
        [string]$Stage, [string[]]$AdditionalNotes = @()
    )
    $accounts = @(); $stores = @(); $folders = @(); $rules = @(); $issues = @()
    if ($null -ne $Inventory) {
        $accounts = @($Inventory.Accounts); $stores = @($Inventory.Stores); $folders = @($Inventory.Folders); $issues += @($Inventory.Issues)
    }
    if ($null -ne $RuleReport) { $rules = @($RuleReport.Rules); $issues += @($RuleReport.Issues) }
    Export-MigrationCsv $accounts @('Index', 'DisplayName', 'SmtpAddress', 'AccountType', 'DeliveryStoreId', 'DeliveryStoreName') (Join-Path $Directory 'accounts.csv')
    Export-MigrationCsv $stores @('StoreId', 'DisplayName', 'FilePath') (Join-Path $Directory 'stores.csv')
    Export-MigrationCsv $folders @('StoreId', 'EntryId', 'FolderPath', 'Name', 'Depth') (Join-Path $Directory 'folders.csv')
    Export-MigrationCsv $rules @('StoreId', 'RuleIndex', 'Name', 'Enabled', 'ExecutionOrder', 'RuleType', 'IsLocalRule', 'EnabledActionTypes', 'EnabledConditionTypes', 'EnabledExceptionTypes', 'MoveStatus', 'MoveFolderPath', 'MoveStoreId', 'MoveEntryId', 'CopyStatus', 'CopyFolderPath', 'CopyStoreId', 'CopyEntryId', 'TechnicalStatus', 'ManualReviewRequired', 'Notes') (Join-Path $Directory 'rules.csv')
    Export-MigrationCsv $issues @('Area', 'Identity', 'Status', 'Message') (Join-Path $Directory 'issues.csv')
    $summary = New-Object 'System.Collections.Generic.List[string]'
    [void]$summary.Add("OutlookMigration - $Stage - $((Get-Date).ToString('o'))")
    if ($null -ne $FileValidation) {
        [void]$summary.Add("RWZ: $($FileValidation.Path)")
        [void]$summary.Add("Bytes: $($FileValidation.Length); SHA-256: $($FileValidation.SHA256)")
        [void]$summary.Add("IntegrityStatus: $($FileValidation.IntegrityStatus)")
        if ($FileValidation.IntegrityStatus -eq 'Unverified') { [void]$summary.Add('File ngoài chưa có manifest: không có checksum gốc để đối chiếu. Đuôi .rwz và SHA-256 không chứng minh Outlook có thể nhập file.') }
    }
    [void]$summary.Add("Tài khoản: $($accounts.Count); kho dữ liệu: $($stores.Count); thư mục: $($folders.Count); vấn đề đọc/kiểm tra: $($issues.Count).")
    if ($null -ne $SelectedAccount) {
        [void]$summary.Add("Tài khoản chọn: [$($SelectedAccount.Index)] $($SelectedAccount.DisplayName) <$($SelectedAccount.SmtpAddress)>")
        [void]$summary.Add("DeliveryStore: $($SelectedAccount.DeliveryStoreName); StoreID: $($SelectedAccount.DeliveryStoreId)")
        [void]$summary.Add('Rules thuộc DeliveryStore. Nhiều tài khoản POP dùng chung kho có thể nhìn thấy cùng bộ Rules; đối chiếu điều kiện through the specified account.')
    }
    if ($null -ne $RuleReport) {
        [void]$summary.Add("Đọc danh sách Rules: $($RuleReport.ReadSucceeded); số rule đọc được: $($rules.Count).")
        foreach ($status in @('Checked', 'NeedsRepair', 'ManualReview')) {
            [void]$summary.Add("${status}: $(@($rules | Where-Object { $_.TechnicalStatus -eq $status }).Count)")
        }
    }
    [void]$summary.Add('Checked = đã kiểm tra liên kết thư mục đọc được; không xác nhận toàn bộ rule hoạt động hay đúng thư mục mong muốn.')
    [void]$summary.Add('NeedsRepair = cần sửa (thiếu tham chiếu thư mục hoặc ràng buộc máy cũ). ManualReview = không kiểm tra tự động được hoặc không có action thư mục để kiểm tra.')
    [void]$summary.Add('Unverifiable = COM không xác minh được; có thể do kho chưa mở/quyền/provider. Xem issues.csv rồi kiểm tra trong Outlook.')
    [void]$summary.Add('Điều kiện/ngoại lệ chỉ được liệt kê mã loại đang bật; không kiểm tra nội dung, người nhận, giá trị hoặc hành động khác. Luôn cần kiểm tra thủ công và email thử.')
    [void]$summary.Add('Script không Save/Execute/sửa Rules, không import/sửa PST và không tự bật rule.')
    foreach ($note in $AdditionalNotes) { [void]$summary.Add($note) }
    $summary.ToArray() | Set-Content -LiteralPath (Join-Path $Directory 'Summary.txt') -Encoding UTF8 -ErrorAction Stop
    Write-Host "Báo cáo: $Directory"
}

function Show-MigrationInventory {
    param($Inventory)
    Write-Host 'Tài khoản trong profile hiện tại:'
    Write-Host ($Inventory.Accounts | Select-Object Index, DisplayName, SmtpAddress, AccountType, DeliveryStoreName | Format-Table -AutoSize | Out-String)
    Write-Host 'Kho dữ liệu đang mở (đối chiếu FilePath để nhận diện PST):'
    Write-Host ($Inventory.Stores | Select-Object DisplayName, FilePath, StoreId | Format-Table -Wrap -AutoSize | Out-String)
    Write-Host 'Cây thư mục đầy đủ: folders.csv. Script không tự chọn PST hay ánh xạ thư mục.'
}

function Invoke-RulesBackup {
    param([string]$RulesFile, [string]$DataRoot, [switch]$SkipOutlook)
    $validation = New-RulesBackup -Path $RulesFile -DataRoot $DataRoot
    $directory = New-MigrationReportDirectory $DataRoot 'Backup'
    $inventory = $null
    $ruleReport = $null
    $connection = $null
    $notes = @('Bản sao .rwz và manifest được tạo; luôn mang theo cả hai file khi chuyển máy.')
    try {
        if ($SkipOutlook) {
            $notes += 'Đã bỏ qua kiểm kê Outlook theo -SkipOutlook. Không có snapshot Rules/thư mục.'
        }
        else {
            try {
                $connection = Connect-RunningOutlook
                $inventory = Get-OutlookInventory -Session $connection.Session
                $allRules = @(); $allIssues = @()
                # Read each store once; POP accounts can share a delivery store.
                $storeIds = @($inventory.Stores | Where-Object { $_.StoreId } | Select-Object -ExpandProperty StoreId -Unique)
                $allRead = $storeIds.Count -gt 0
                foreach ($id in $storeIds) {
                    $report = Get-StoreRuleReport -Session $connection.Session -StoreId $id
                    $allRules += @($report.Rules); $allIssues += @($report.Issues)
                    $allRead = $allRead -and $report.ReadSucceeded
                }
                $ruleReport = [pscustomobject]@{ Rules = $allRules; Issues = $allIssues; ReadSucceeded = $allRead }
                $notes += 'Snapshot phản ánh Outlook đang mở lúc backup, không phải nội dung được giải mã từ file .rwz.'
            }
            catch {
                $notes += 'Đã backup file Rules; chưa kiểm kê được Outlook: ' + $_.Exception.Message
                $issue = [pscustomobject]@{ Area = 'Outlook'; Identity = 'Connection'; Status = 'ManualReview'; Message = $_.Exception.Message }
                $inventory = [pscustomobject]@{ Accounts = @(); Stores = @(); Folders = @(); Issues = @($issue) }
                Write-Warning $notes[-1]
            }
        }
        Write-MigrationReport -Directory $directory -FileValidation $validation -Inventory $inventory -RuleReport $ruleReport -Stage 'Backup' -AdditionalNotes $notes
        Write-Host "Backup Rules: $($validation.Path)"
        [pscustomobject]@{ FileValidation = $validation; ReportDirectory = $directory }
    }
    finally {
        if ($null -ne $connection) {
            Release-OutlookReference $connection.Session
            Release-OutlookReference $connection.Application
        }
    }
}

function Invoke-RulesRestore {
    param(
        [string]$RulesFile, [string]$DataRoot,
        [ValidateSet('Inspect', 'GuideImport', 'Verify')][string]$Stage = 'Inspect',
        [int]$AccountIndex = 0
    )
    # Validation precedes COM attachment and all import guidance.
    $validation = Get-RulesFileValidation -Path $RulesFile
    if ($validation.IntegrityStatus -eq 'Unverified') {
        Write-Warning 'File .rwz chưa có manifest để đối chiếu checksum gốc; Outlook sẽ kiểm tra khả năng nhập. Script không giải mã nội dung .rwz.'
    }
    $directory = New-MigrationReportDirectory $DataRoot $Stage
    $inventory = $null; $ruleReport = $null; $selected = $null; $connection = $null
    $notes = @()
    try {
        $connection = Connect-RunningOutlook
        $inventory = Get-OutlookInventory -Session $connection.Session
        Show-MigrationInventory $inventory
        # Retain inventory even when the user cancels or selection fails.
        Write-MigrationReport -Directory $directory -FileValidation $validation -Inventory $inventory -Stage $Stage -AdditionalNotes @('Chưa chọn tài khoản; chưa kiểm tra Rules của kho đích.')
        $selected = Select-MigrationAccount -Accounts $inventory.Accounts -AccountIndex $AccountIndex
        Write-Host "Đích đã chọn: $($selected.DisplayName) <$($selected.SmtpAddress)> / $($selected.DeliveryStoreName)"
        $ruleReport = Get-StoreRuleReport -Session $connection.Session -StoreId $selected.DeliveryStoreId

        if ($Stage -eq 'GuideImport') {
            $beforeDirectory = Join-Path $directory 'PreImport'
            [IO.Directory]::CreateDirectory($beforeDirectory) | Out-Null
            Write-MigrationReport -Directory $beforeDirectory -FileValidation $validation -Inventory $inventory -RuleReport $ruleReport -SelectedAccount $selected -Stage 'PreImport'
            $beforeCount = $ruleReport.Rules.Count
            $allowEmpty = $ruleReport.ReadSucceeded -and $beforeCount -eq 0
            $notes += 'Snapshot trước nhập được lưu riêng trong PreImport. Chưa xác minh thao tác nhập.'
            Write-Host '1. Trong Outlook Classic: File > Manage Rules & Alerts.'
            Write-Host "2. Ở Apply changes to this folder, chọn Inbox thuộc tài khoản/kho: $($selected.DisplayName) / $($selected.DeliveryStoreName)."
            Write-Host '3. Options > Export Rules: lưu bộ Rules hiện có ra một file .rwz KHÁC file nguồn, để có thể khôi phục thủ công.'
            Write-Host 'Lưu ý: import thêm vào bộ Rules hiện có; nhập lặp có thể tạo rule trùng. Kiểm tra Rules hiện có trước khi nhập.'
            $prompt = 'Gõ SAO-LUU khi đã xuất Rules hiện có, hoặc HUY để dừng'
            if ($allowEmpty) { $prompt += '; nếu đã kiểm tra đúng kho và không có Rules, có thể gõ TRONG' }
            $confirmation = Read-Host $prompt
            if ($confirmation -ine 'SAO-LUU' -and -not ($allowEmpty -and $confirmation -ieq 'TRONG')) {
                throw 'Đã hủy: chưa xác nhận backup Rules hiện có; chưa hướng dẫn nhập.'
            }
            $current = Get-RulesFileValidation -Path $validation.Path
            if ($current.SHA256 -ne $validation.SHA256 -or $current.Length -ne $validation.Length) {
                throw 'File .rwz đã thay đổi kể từ lúc kiểm tra; dừng và chạy lại.'
            }
            # From this point the user may change Outlook. Do not label the old
            # snapshot as current if import is cancelled or post-checks fail.
            $ruleReport = $null
            Write-Host '4. Tại cùng tài khoản/kho: Options > Import Rules, chọn chính xác file:'
            Write-Host $validation.Path
            Write-Host '5. Xem cảnh báo và liên kết thư mục; xác nhận thao tác trong Outlook rồi đóng hộp thoại bằng Apply/OK.'
            Write-Host 'Không chọn Run Rules Now trên thư cũ trong quá trình chuyển.'
            $confirmation = Read-Host 'Gõ DA-NHAP khi đã nhập và đóng hộp thoại Rules trong Outlook, hoặc HUY'
            if ($confirmation -ine 'DA-NHAP') { throw 'Đã hủy bước kiểm tra sau import. Nếu đã thao tác trong Outlook, script không tự hoàn tác; chạy -Stage Verify sau.' }
            $current = Get-RulesFileValidation -Path $validation.Path
            if ($current.SHA256 -ne $validation.SHA256 -or $current.Length -ne $validation.Length) {
                throw 'File .rwz thay đổi trong lúc nhập; không thể đối chiếu cùng nguồn. Kiểm tra lại thủ công.'
            }
            # Re-read after the manual operation; never reuse the cached Rules collection.
            $inventory = Get-OutlookInventory -Session $connection.Session
            $matches = @($inventory.Accounts | Where-Object {
                $_.DeliveryStoreId -eq $selected.DeliveryStoreId -and $_.SmtpAddress -eq $selected.SmtpAddress -and $_.DisplayName -eq $selected.DisplayName
            })
            if ($matches.Count -ne 1) { throw 'Tài khoản/profile/kho đích đã thay đổi; chạy lại Inspect để chọn đích hiện tại.' }
            $selected = $matches[0]
            $ruleReport = Get-StoreRuleReport -Session $connection.Session -StoreId $selected.DeliveryStoreId
            $notes += "Người dùng xác nhận đã nhập thủ công. Số rule trước: $beforeCount; sau: $($ruleReport.Rules.Count). Số lượng không chứng minh import thành công/đầy đủ. Snapshot trước: PreImport."
        }
        elseif ($Stage -eq 'Inspect') {
            $notes += 'Chỉ kiểm tra trước; chưa nhập Rules. Dùng -Stage GuideImport để được hướng dẫn nhập, hoặc Verify nếu bạn đã tự nhập.'
        }
        else { $notes += 'Kiểm tra trạng thái hiện tại; script không xác nhận file .rwz đã được nhập hay mọi rule đã chạy đúng.' }
        Write-MigrationReport -Directory $directory -FileValidation $validation -Inventory $inventory -RuleReport $ruleReport -SelectedAccount $selected -Stage $Stage -AdditionalNotes $notes
        if (-not $ruleReport.ReadSucceeded) { throw 'Không đọc được danh sách Rules của kho đích. Xem issues.csv rồi kiểm tra lại trong Outlook.' }
        Write-Host 'Đối chiếu rules.csv và issues.csv; sửa liên kết trong Outlook, rồi gửi email thử cho các nhóm Rules quan trọng.'
        [pscustomobject]@{ Stage = $Stage; SelectedAccount = $selected; RuleReport = $ruleReport; ReportDirectory = $directory }
    }
    catch {
        $failure = $_
        if ($null -eq $inventory) {
            $issue = [pscustomobject]@{ Area = 'Outlook'; Identity = 'Connection'; Status = 'ManualReview'; Message = $failure.Exception.Message }
            $inventory = [pscustomobject]@{ Accounts = @(); Stores = @(); Folders = @(); Issues = @($issue) }
        }
        # Do not mask the original error if the report destination itself fails.
        try {
            Write-MigrationReport -Directory $directory -FileValidation $validation -Inventory $inventory -RuleReport $ruleReport -SelectedAccount $selected -Stage $Stage -AdditionalNotes ($notes + @('DỪNG / CHƯA HOÀN TẤT: ' + $failure.Exception.Message))
        }
        catch { Write-Warning ("Không ghi được báo cáo lỗi: " + $_.Exception.Message) }
        throw $failure
    }
    finally {
        if ($null -ne $connection) {
            Release-OutlookReference $connection.Session
            Release-OutlookReference $connection.Application
        }
    }
}
