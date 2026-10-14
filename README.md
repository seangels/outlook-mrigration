# OutlookMigration — chuyển Rules trong Outlook Classic

Bản này hỗ trợ sao lưu file Rules `.rwz`, kiểm kê Outlook, hướng dẫn nhập Rules và sửa hàng loạt liên kết thư mục Move/Copy sau khi chuyển máy. PST của lần chuyển này đã được nhập thủ công. Script không nhập hay sao chép PST; các thư mục `Signatures`, `Templates`, `PST` được chuẩn bị cho các phiên bản tiếp theo. AutoComplete và cấu hình Outlook khác chưa được chuyển.

`Backup.ps1` và `Restore.ps1` chỉ đọc Outlook để lập báo cáo; việc xuất/nhập `.rwz` vẫn thực hiện trong giao diện Outlook. `Repair-Rules.ps1` bổ sung hai bước riêng: `Plan` chỉ lập bảng ánh xạ để xem trước; `Apply` mới lưu các liên kết thư mục được chọn. Script không tạo/xóa Rules hoặc thư mục và không chạy Rules trên thư cũ.

## Chuẩn bị

- Dùng **Outlook Classic cho Windows** và **Windows PowerShell 5.1**, dưới tài khoản Windows thường dùng Outlook. Không cần chạy với quyền quản trị.
- Trên máy mới, cấu hình tài khoản POP/IMAP, mở Outlook và xác nhận các PST/thư mục đích cần dùng đã hiện trong Outlook.
- Để script đọc được Outlook, mở Outlook Classic và hoàn tất mọi hộp thoại chọn profile, đăng nhập trước khi chạy. Script chỉ kết nối phiên Outlook đang chạy, không tự khởi động Outlook hay tạo profile.
- Có file `.rwz` xuất từ máy cũ. Nếu chưa có, làm bước xuất bên dưới. Đổi đuôi một file khác thành `.rwz` không tạo thành Rules hợp lệ.
- Giữ `Common.ps1` cạnh các script chính. Để sửa hàng loạt, chép thêm `Repair-Rules.ps1` và `Repair.Common.ps1`, đồng thời cập nhật `Common.ps1` từ cùng một bộ script.

```text
OutlookMigration\
├── Backup.ps1
├── Restore.ps1
├── Common.ps1
├── Repair-Rules.ps1
├── Repair.Common.ps1
├── README.md
├── Rules\
├── Signatures\
├── Templates\
├── PST\
└── Report\
```

Mặc định dữ liệu nằm trong thư mục chứa script. Có thể dùng `-DataRoot "D:\Sao lưu Outlook"` để chọn nơi chứa các thư mục dữ liệu và báo cáo. Dùng dấu ngoặc kép cho đường dẫn có khoảng trắng hoặc tiếng Việt.

## 1. Sao lưu Rules từ máy cũ

Trong Outlook Classic, vào **File → Manage Rules & Alerts**. Chọn đúng tài khoản trong danh sách **Apply changes to this folder**, rồi chọn **Options → Export Rules** và lưu thành `.rwz`. Đây là thao tác thủ công; `Backup.ps1` nhận file đã xuất, không tự xuất Rules từ Outlook.

Mở Windows PowerShell tại thư mục bộ script và chạy:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File ".\Backup.ps1" -RulesFile "D:\Sao lưu Outlook\Rules máy cũ.rwz"
```

Script sao chép file vào `Rules` với dấu thời gian và mã riêng để không ghi đè bản cũ, rồi tạo file `<tên-file>.rwz.manifest.json` đi kèm, gồm kích thước và SHA-256. Giữ nguyên tên hai file để đối chiếu manifest. Nếu kết nối được Outlook đang mở, script lập thêm báo cáo kiểm kê để đối chiếu sau này. Snapshot phản ánh phiên Outlook lúc chạy, không phải nội dung được giải mã từ `.rwz`. Không đọc được Outlook không làm mất bản sao `.rwz`; xem cảnh báo và báo cáo để biết phần nào còn thiếu.

Nếu chỉ muốn lưu file Rules, không cần kiểm kê Outlook:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File ".\Backup.ps1" -RulesFile "D:\Sao lưu Outlook\Rules máy cũ.rwz" -SkipOutlook
```

Chuyển cả thư mục bộ script, hoặc ít nhất file `.rwz` **cùng manifest đi kèm**, sang máy mới. Dùng đường dẫn bản sao trong `Rules` do script vừa in ra cho các bước tiếp theo.

## 2. Kiểm tra trên máy mới trước khi nhập

Mở Outlook Classic trên máy mới, rồi chạy:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File ".\Restore.ps1" -RulesFile ".\Rules\<tên bản sao>.rwz"
```

Thay `<tên bản sao>` bằng tên file thực tế. Mặc định là `-Stage Inspect`: xác minh file và lập danh sách tài khoản, kho dữ liệu, cây thư mục, cùng Rules của tài khoản được chọn. Chưa có thao tác nhập ở bước này.

Script luôn yêu cầu chọn tài khoản đích, kể cả chỉ có một tài khoản. Chọn theo số thứ tự trong danh sách hiện tại; script không tự suy đoán PST đích. Đối chiếu tên kho dữ liệu và đường dẫn thư mục trong báo cáo với Outlook. Nhiều tài khoản POP có thể dùng chung một kho dữ liệu; Rules được đọc từ `DeliveryStore` của tài khoản đã chọn nên danh sách có thể dùng chung giữa các tài khoản đó.

Nếu cần chạy kiểm kê không tương tác, truyền số tài khoản đã đối chiếu với danh sách của lần chạy hiện tại:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File ".\Restore.ps1" -RulesFile ".\Rules\<tên bản sao>.rwz" -Stage Inspect -AccountIndex 1
```

`-AccountIndex` là số nguyên dương, bắt đầu từ 1. Số thứ tự có thể thay đổi khi cấu hình tài khoản thay đổi; không xem đây là mã tài khoản cố định.

Nếu có manifest đi kèm, kích thước và SHA-256 phải khớp. File `.rwz` ngoài bộ backup, không có manifest, vẫn được tiếp nhận nhưng bị đánh dấu **chưa xác minh tính toàn vẹn**. Hash khớp chỉ cho biết file không thay đổi so với manifest, không chứng minh nội dung `.rwz` hợp lệ hoặc tương thích với Outlook hiện tại.

## 3. Nhập Rules qua hướng dẫn

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File ".\Restore.ps1" -RulesFile ".\Rules\<tên bản sao>.rwz" -Stage GuideImport
```

Thực hiện theo các nhắc trong cửa sổ PowerShell:

1. Chọn và kiểm tra lại tài khoản đích. Trong Outlook, mở **File → Manage Rules & Alerts**, chọn đúng tài khoản ở **Apply changes to this folder**.
2. Chọn **Options → Export Rules** để sao lưu Rules hiện có trên máy mới vào một file riêng. Gõ `SAO-LUU` sau khi đã xuất. Nếu script đọc thành công danh sách trống, bạn có thể gõ `TRONG` sau khi đã kiểm tra đúng kho trong Outlook và xác nhận không có Rules.
3. Chọn **Options → Import Rules**, mở đúng file `.rwz` mà script chỉ ra và hoàn tất các hộp thoại Outlook. Kiểm tra danh sách Rules trước khi lưu các thay đổi trong Outlook.
4. Quay lại PowerShell, gõ `DA-NHAP` khi đã hoàn tất. Script đọc lại trạng thái Outlook và lập báo cáo kiểm tra sau nhập. Có thể gõ `HUY` ở lời nhắc hỗ trợ hủy để kết thúc hướng dẫn; việc này không hoàn tác thao tác đã thực hiện trong Outlook.

Outlook thêm Rules được nhập vào danh sách hiện có. Vì vậy, **không nhập lặp lại cùng file** để thử sửa lỗi: có thể tạo Rules trùng và làm thay đổi thứ tự xử lý. Script không tự loại Rules trùng hoặc hoàn tác import. Cách xuất/nhập và hành vi thêm vào danh sách được Microsoft mô tả tại [hướng dẫn nhập hoặc xuất Rules trong Outlook Classic](https://support.microsoft.com/en-us/outlook/mail/import-or-export-a-set-of-rules-in-classic-outlook).

Trong lúc hướng dẫn, script vẫn không điều khiển nút trong Outlook và không tự áp dụng thay đổi. Các token xác nhận chỉ ghi nhận thao tác người dùng, không kiểm chứng rằng import thành công.

## 4. Đọc báo cáo và sửa liên kết thư mục

Mỗi lần chạy có thư mục báo cáo riêng trong `Report`, mang dấu thời gian, tên giai đoạn và mã riêng. Đường dẫn được in ra ở cửa sổ PowerShell. Các file CSV/TXT dùng UTF-8 có BOM để giữ tiếng Việt khi mở bằng công cụ Windows. `GuideImport` lưu snapshot trước nhập vào thư mục con `PreImport`; báo cáo ở thư mục cha được cập nhật sau khi xác nhận nhập. Nếu hủy hoặc kiểm tra sau nhập thất bại, đọc dòng **DỪNG / CHƯA HOÀN TẤT** trong `Summary.txt`.

| File | Nội dung |
| --- | --- |
| `accounts.csv` | Tài khoản và kho nhận thư tương ứng; dùng để chọn đúng tài khoản đích. |
| `stores.csv` | Các mailbox/PST đang mở mà Outlook cho phép đọc. |
| `folders.csv` | Cây thư mục và định danh đọc được để đối chiếu đích chuyển/copy. |
| `rules.csv` | Tên Rules, trạng thái bật/tắt, thứ tự và kết quả kiểm tra các hành động chuyển/copy. |
| `issues.csv` | Lỗi đọc COM, mục cần sửa hoặc cần kiểm tra thủ công. |
| `Summary.txt` | Kết quả tổng hợp, giới hạn kiểm tra và việc cần làm tiếp. |

Các nhãn kiểm tra có ý nghĩa:

- **`Checked` — đã kiểm tra:** thuộc tính hoặc liên kết thư mục đọc được, và kiểm tra kỹ thuật liên quan đã đạt. Vẫn cần xác nhận đó là đúng thư mục mong muốn và thử email thực tế.
- **`NeedsRepair` — cần sửa:** action chuyển/copy đang bật nhưng thiếu tham chiếu thư mục, hoặc có điều kiện `OnOtherMachine` ràng buộc máy cũ; mở rule trong Outlook để sửa.
- **`ManualReview` — cần kiểm tra thủ công:** thiếu thông tin để kết luận tự động, chẳng hạn COM không cho đọc một phần rule. Đọc chi tiết trong báo cáo và đối chiếu Outlook.

Điều kiện, ngoại lệ, các hành động khác chuyển/copy và ý nghĩa nghiệp vụ của thư mục đích luôn cần kiểm tra thủ công. Nếu Outlook/COM không cung cấp thông tin, báo cáo ghi rõ; ô trống không đồng nghĩa với giá trị hợp lệ. Một thư mục đọc được vẫn có thể thuộc PST hoặc tài khoản khác với mục đích của rule. Số lượng Rules khớp không đủ để kết luận migration thành công.

`MoveStatus`/`CopyStatus` có thể là `Resolved` (resolve được thư mục), `NotEnabled` (action tắt), `MissingTarget` (thiếu tham chiếu), hoặc `Unverifiable` (COM không xác minh được, có thể do quyền/provider/kho chưa mở). Các cột `EnabledActionTypes`, `EnabledConditionTypes`, `EnabledExceptionTypes` chỉ liệt kê mã loại đang bật, không có giá trị chi tiết của điều kiện. Tra mã tại [OlRuleActionType](https://learn.microsoft.com/en-us/office/vba/api/outlook.olruleactiontype) và [OlRuleConditionType](https://learn.microsoft.com/en-us/office/vba/api/outlook.olruleconditiontype).

Để sửa một rule, vào **Manage Rules & Alerts**, chọn rule và mở phần chỉnh sửa. Trong mô tả rule, bấm liên kết thư mục của hành động chuyển/copy, rồi chọn lại đúng thư mục trong cây mailbox/PST trên máy mới. Kiểm tra cả điều kiện tài khoản, ngoại lệ, thứ tự, tùy chọn dừng xử lý và trạng thái bật/tắt trước khi lưu. Nếu Outlook báo lỗi không thể sửa bằng chọn lại thư mục, tạo lại rule tương ứng bằng giao diện Outlook, dựa trên cấu hình máy cũ.

Sau khi sửa thủ công hoặc sửa hàng loạt ở mục 5, chạy kiểm tra mới; `Verify` chỉ đọc, không nhập lại file:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File ".\Restore.ps1" -RulesFile ".\Rules\<tên bản sao>.rwz" -Stage Verify
```

Script dùng Outlook COM để đọc Rules từ kho dữ liệu. Khả năng kiểm tra phụ thuộc thông tin Outlook cung cấp, theo giao diện [Store.GetRules của Microsoft](https://learn.microsoft.com/en-us/office/vba/api/outlook.store.getrules).

## 5. Sửa hàng loạt liên kết thư mục trên máy mới

Áp dụng khi đã nhập `.rwz`, nhiều Rules hiện `(Error)` và chọn lại thư mục bằng tay giải quyết được lỗi. Cần báo cáo `Backup` từ máy cũ có `rules.csv`: đây là nguồn đường dẫn thư mục cũ để đối chiếu; script không giải mã các đường dẫn này từ file `.rwz`.

### Chuẩn bị trên máy mới

1. Chép bản mới của `Repair-Rules.ps1`, `Repair.Common.ps1` và `Common.ps1` vào `D:\OutlookMigration`, cạnh `Backup.ps1` và `Restore.ps1`.
2. Chép nguyên thư mục báo cáo `Backup` từ máy cũ vào `D:\OutlookMigration\Report`. Dùng thư mục chứa `rules.csv` của máy cũ, không dùng nhầm báo cáo `Restore` vừa tạo trên máy mới.
3. Mở Outlook Classic trên máy mới, xác nhận mailbox/PST và các thư mục cần dùng đang hiện. Đóng cửa sổ **Rules and Alerts** trước khi chạy `Apply` để tránh chỉnh Rules đồng thời với script.

### Lập bảng ánh xạ, chưa sửa Outlook

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "D:\OutlookMigration\Repair-Rules.ps1" -Stage Plan -SourceReport "D:\OutlookMigration\Report\<old-Backup-folder>" -DataRoot "D:\OutlookMigration"
```

Thay `<old-Backup-folder>` bằng tên thư mục báo cáo cũ thực tế. `-SourceReport` nhận thư mục chứa `rules.csv` hoặc đường dẫn đầy đủ tới file `rules.csv` đó.

Script yêu cầu chọn tài khoản đang chứa các Rules đã import. Có thể dùng `-AccountIndex 1` nếu đã đối chiếu số tài khoản trong lần chạy hiện tại. Nếu báo cáo cũ chứa Rules từ nhiều kho, script chỉ tự chọn nguồn khi xác định duy nhất hoặc khớp tài khoản SMTP; nếu chưa rõ, chọn nguồn khi được nhắc hoặc truyền `-SourceStoreId` lấy từ cột `StoreId` của `rules.csv` máy cũ.

Với mỗi kho chứa thư mục đích cũ, bạn phải chọn kho đích tương ứng đang mở trên máy mới. Ví dụ: thư mục cũ thuộc PST lưu trữ thì chọn đúng PST mới đang chứa cây thư mục đó. Script ghép Rules theo **tên + loại rule** khi kết quả là duy nhất, rồi tìm **đường dẫn tương đối chính xác** bên trong kho đích bạn đã chọn. Script không đoán thư mục từ tên rule. Trùng tên rule, thiếu đường dẫn nguồn hoặc không tìm thấy đúng thư mục đều được để lại để kiểm tra thủ công. Liên kết thư mục hiện tại đã xác minh được sẽ được bỏ qua; bạn vẫn cần kiểm tra nó đúng ý định sử dụng.

Kết quả nằm trong thư mục mới `Report\<dấu-thời-gian>-RepairPlan-...` được in ra:

| File | Cách dùng |
| --- | --- |
| `folder-mapping.csv` | Mỗi dòng là một đích Move/Copy cần đối chiếu; chỉ những dòng `Apply=Yes` được sửa. |
| File `.plan.json` đi kèm | Thông tin nhận diện và kiểm tra kế hoạch trước khi lưu; giữ nguyên cạnh CSV. |
| `folders.csv` | Danh sách thư mục hiện có trên máy mới để lấy đúng `StoreId` và `FolderPath`. |
| Các báo cáo kiểm kê và `Summary.txt` | Thông tin tài khoản, Rules và lý do các dòng cần xử lý thủ công. |

Mở `folder-mapping.csv` để duyệt. Chỉ sửa ba cột **`Apply`** (`Yes` hoặc `No`), **`TargetStoreId`** và **`TargetFolderPath`**. Khi sửa đích, chép `StoreId` và **toàn bộ** `FolderPath` từ cùng một dòng trong `folders.csv`; không chỉ gõ tên thư mục cuối. Các dòng không được đề xuất tự động có thể ánh xạ thủ công theo cách này rồi đổi `Apply` thành `Yes`. Giữ nguyên các cột mô tả, `RowId` và file `.plan.json`; lưu CSV bằng UTF-8, không đổi tên cột. Nếu đã sửa Rules trong Outlook sau khi tạo bảng, chạy lại `Plan` thay vì dùng bảng cũ.

Kiểm tra độ mới của bảng dựa trên metadata và định danh thư mục đọc được; không phát hiện mọi thay đổi trong giá trị điều kiện, người gửi hoặc action khác. Giữ cửa sổ Rules and Alerts đóng và không chỉnh Rules trong lúc áp dụng.

### Sao lưu Rules hiện tại và xem trước

Trong Outlook, mở **File → Manage Rules & Alerts**, chọn đúng tài khoản, dùng **Options → Export Rules** để lưu Rules hiện tại thành `D:\OutlookMigration\Rules\BeforeRepair.rwz`. Script không tự xuất được bản sao này. Sau đó đóng cửa sổ **Rules and Alerts**.

Chạy thử với `-WhatIf`, thay đường dẫn bảng bằng đường dẫn thực tế script đã in:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "D:\OutlookMigration\Repair-Rules.ps1" -Stage Apply -MappingFile "D:\OutlookMigration\Report\<RepairPlan-folder>\folder-mapping.csv" -DataRoot "D:\OutlookMigration" -WhatIf
```

`-WhatIf` kiểm tra bảng và các đích trong Outlook, hiển thị thao tác dự kiến trong PowerShell, không tạo báo cáo mới hoặc lưu thay đổi. File backup chỉ bắt buộc khi áp dụng thật; vẫn nên xuất trước để sẵn sàng phục hồi thủ công.

### Áp dụng những dòng đã duyệt

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "D:\OutlookMigration\Repair-Rules.ps1" -Stage Apply -MappingFile "D:\OutlookMigration\Report\<RepairPlan-folder>\folder-mapping.csv" -BackupRulesFile "D:\OutlookMigration\Rules\BeforeRepair.rwz" -DataRoot "D:\OutlookMigration"
```

Script kiểm tra toàn bộ các dòng được chọn trước khi gán các đích Move/Copy, giữ trạng thái bật/tắt của action và gọi lưu Rules một lần. Những rule đã import ở trạng thái tắt sẽ **vẫn tắt theo mặc định**. Sau khi kiểm tra kết quả, bật chúng trong Outlook khi sẵn sàng thử email.

Có thể thêm `-EnableRepairedRules` nếu muốn bật lại các rule đã sửa mà báo cáo cũ xác nhận vốn đang bật. Tùy chọn này chỉ bật rule đối chiếu được duy nhất, mọi đích thư mục đã xác minh được và không có điều kiện ràng buộc máy cũ `OnOtherMachine`. Rule vốn tắt trên máy cũ hoặc chưa đủ điều kiện sẽ không được bật tự động. Script không sửa điều kiện, ngoại lệ, thứ tự hoặc các action khác.

Báo cáo áp dụng gồm snapshot `BeforeApply`, `AfterApply`, `repair-results.csv` và `Summary.txt`. Đối chiếu các dòng thành công, chưa sửa hoặc cần kiểm tra thủ công. Nếu lưu thất bại, trạng thái có thể chưa xác định: **không tự chạy lại Apply**, kiểm tra Rules trong Outlook và báo cáo trước. Script không tự hoàn tác về các liên kết thư mục đã hỏng. File `BeforeRepair.rwz` dùng để phục hồi thủ công; vì Outlook nhập thêm Rules vào danh sách hiện có, kiểm tra kỹ danh sách trước khi nhập lại để tránh trùng. Khi cần tiếp tục sau khi thay đổi Rules, tạo một `Plan` mới.

Thao tác sửa dùng thuộc tính ghi được [MoveOrCopyRuleAction.Folder](https://learn.microsoft.com/en-us/office/vba/api/outlook.moveorcopyruleaction.folder). Outlook có thể từ chối thư mục không phù hợp, thư mục tìm kiếm, chỉ đọc hoặc thiếu quyền. Lệnh [Rules.Save](https://learn.microsoft.com/en-us/office/vba/api/outlook.rules.save) lưu toàn bộ bộ Rules của kho được chọn và có thể lỗi nếu còn rule/điều kiện không hợp lệ; không có đảm bảo lưu nguyên tử. Các báo cáo kỹ thuật không thay thế nghiệm thu email thực tế.

## 6. Nghiệm thu bằng email thử

- Chọn từng nhóm rule quan trọng: chuyển thư vào thư mục PST, copy thư, điều kiện người gửi/chủ đề và rule có ngoại lệ hoặc dừng xử lý.
- Trong Outlook, chủ động kiểm tra và bật các Rules đã sửa khi sẵn sàng. Gửi email mới thỏa điều kiện tới đúng tài khoản, giữ Outlook mở, rồi kiểm tra kết quả ở đúng thư mục đích.
- Thử thêm một email không thỏa điều kiện hoặc thuộc ngoại lệ để phát hiện rule chạy quá rộng. Đối chiếu thứ tự khi nhiều rule có thể cùng áp dụng.
- Ghi lại kết quả kiểm thử và chạy `Verify` lần cuối để lưu trạng thái sau sửa. Không dùng chức năng chạy Rules trên toàn bộ thư cũ cho bước thử này.

## Xử lý tình huống thường gặp

**Không kết nối được Outlook:** kiểm tra đang dùng Outlook Classic, đã mở profile và hoàn tất các hộp thoại. Chạy Outlook và PowerShell bằng cùng tài khoản Windows, cùng mức quyền thông thường. Đóng các hộp thoại đang chặn Outlook rồi thử lại. `Backup` vẫn có thể lưu file với `-SkipOutlook`; `Restore` cần kết nối Outlook để kiểm kê tài khoản đích. Không sửa được lỗi đọc COM thì dùng `issues.csv` và giao diện Outlook để kiểm tra thủ công.

**Không có tài khoản hoặc chưa có profile:** mở Outlook và cấu hình tài khoản bằng giao diện trước. Script không tự tạo profile hoặc gắn thêm PST.

**DataRoot bị rỗng ở bản script cũ:** cập nhật cả `Backup.ps1` và `Restore.ps1`. Có thể chạy ngay bằng cách thêm `-DataRoot "D:\Sao lưu Outlook\OutlookMigration"` với đường dẫn thư mục bộ script thực tế. Bản hiện tại xác định giá trị mặc định trong thân script, tương thích Windows PowerShell 5.1.

**Hash/kích thước không khớp hoặc manifest hỏng:** dừng dùng bản đó, lấy lại cặp `.rwz` và manifest nguyên vẹn từ bản backup. Không chỉnh manifest chỉ để bỏ qua lỗi. Nếu chỉ còn file `.rwz` độc lập, giữ nguyên bản gốc và hiểu rằng tính toàn vẹn chưa được xác minh.

**PowerShell chặn thực thi:** kiểm tra nguồn bộ script trước. Nếu chính sách máy cho phép, có thể dùng `-ExecutionPolicy Bypass` chỉ cho tiến trình chạy lệnh bên dưới; không đổi chính sách toàn máy. Chính sách do tổ chức quản lý có thể vẫn chặn lệnh, khi đó liên hệ người quản trị máy.

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File ".\Restore.ps1" -RulesFile ".\Rules\<tên bản sao>.rwz" -Stage Inspect
```

**Tên thư mục giống nhau:** đối chiếu toàn bộ đường dẫn và kho chứa trong `folders.csv`, sau đó chọn lại trong Outlook hoặc trong bảng sửa hàng loạt. Kế hoạch tự động chỉ dùng đường dẫn tương đối chính xác trong kho đích bạn chọn; không ghép theo tên thư mục cuối vì tên giống nhau không đảm bảo cùng đích.

**Bảo quản bộ backup:** `.rwz` và báo cáo có thể chứa tên tài khoản, đường dẫn PST và thông tin tổ chức thư. Lưu cùng những bản sao dữ liệu cá nhân khác của bạn và chỉ chia sẻ các phần cần thiết khi nhờ hỗ trợ.

## Kiểm thử bộ script

Bộ kiểm thử cần Pester 3.4 trở lên, chạy bằng Windows PowerShell 5.1:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File ".\tests\Run-Tests.ps1"
```

Kiểm thử sử dụng file tạm và đối tượng mô phỏng COM, không kết nối mailbox thật. Có kiểm tra entry script backup với `-SkipOutlook`. Kiểm thử tự động không thay thế việc nghiệm thu bằng Outlook Classic và email thực tế trên máy đích.
