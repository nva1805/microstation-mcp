# Bộ máy kết nối nhanh (VBA in-process): dựng thế nào, có an toàn không

## Tóm tắt
- Toàn bộ mã chạy bên trong MicroStation là **một file chữ**: [`vba/McpFast.bas`](../vba/McpFast.bas)
  (~560 dòng, chỉ ký tự ASCII). Không có DLL, MDL, `.ma`, `.exe` hay `.mvba` dựng sẵn trong repo.
- File project `.mvba` được **dựng trên chính máy của bạn** từ file `.bas` đó, bằng các key-in
  chuẩn của MicroStation. Repo cố ý **không** chứa `.mvba` vì đó là file nhị phân không đọc được.
- Bạn có thể tự kiểm chứng mã đang chạy trong MicroStation trùng với file nguồn:
  `python scripts/vba_engine.py verify`.

## Vì sao cần VBA thay vì chỉ COM
COM từ Python là gọi **liên tiến trình**: mỗi lần đọc một thuộc tính (màu, level, đỉnh...) mất
vài mili-giây. File địa chính ~27.000 đối tượng → hơn 3 phút. VBA chạy **bên trong** `ustation.exe`
nên cùng vòng lặp chỉ mất ~1 giây. COM vẫn được dùng, nhưng chỉ để gửi vài key-in điều khiển.

## Các bước dựng (tự động khi MCP chạy lần đầu)
Python (`src/core/engine.py`, hàm `Engine.ensure`) gửi đúng các key-in sau qua COM:

```
vba create  %LOCALAPPDATA%\msmcp\msmcp_<sha8>.mvba
vba execute VBE.VBProjects("msmcp_<sha8>").VBComponents.Import "%LOCALAPPDATA%\msmcp\McpFast.bas"
vba save    msmcp_<sha8>
```

- `<sha8>` = 8 ký tự đầu SHA-1 của `vba/McpFast.bas`. Mã nguồn đổi → tên project đổi → dựng lại.
- `McpFast.bas` trong `%LOCALAPPDATA%\msmcp\` là bản sao nguyên văn của file trong repo.
- Những lần sau: `vba load <file .mvba>`.
- Xem các lệnh và mã băm trên máy bạn: `python scripts/vba_engine.py info`.

## Giao thức chạy
1. Python ghi `%LOCALAPPDATA%\msmcp\req.txt` (ASCII, dạng `khoa=gia_tri`; chuỗi tiếng Việt mã hóa hex).
2. Python gửi key-in `vba run [msmcp_<sha8>]McpFast.Main`.
3. VBA đọc `req.txt`, làm việc, ghi `resp.txt` (JSON từng dòng) rồi `done.txt` (mã phiên).

## Mã VBA làm được gì / không làm gì
**Có**:
- Đọc đối tượng của file đang mở, hoặc mở một file khác ở chế độ **chỉ đọc**
  (`OpenDesignFileForProgram(path, True)`).
- Ghi vào file đang mở các lệnh do bạn duyệt qua `ms_edit`: tạo, sửa level/màu/nét, di chuyển, sửa chữ, xóa, nối đường.
- Đọc/ghi đúng 3 file trong `%LOCALAPPDATA%\msmcp\` (`req.txt`, `resp.txt`, `done.txt`).

**Không** (kiểm tra được bằng cách đọc file `.bas`):
- Không `Shell`, không `CreateObject`, không gọi API Windows (`Declare`), không truy cập mạng.
- Không ghi file nào ngoài 3 file giao tiếp trên; không đụng registry.

Lệnh tìm nhanh trong mã nguồn (kết quả mong đợi: không có dòng nào):

```powershell
Select-String -Path vba\McpFast.bas -Pattern 'Shell|CreateObject|Declare |Kill |URLDownload|WScript|Environ\$\("(?!LOCALAPPDATA)'
```

## Kiểm chứng mã đang chạy
```powershell
.venv\Scripts\python scripts\vba_engine.py verify
```
Script xuất module `McpFast` từ project đang nạp trong MicroStation và so từng dòng với
`vba/McpFast.bas`. VBA tự đổi hoa/thường tên biến trong toàn project (không ảnh hưởng chạy), nên
phần mã được so không phân biệt hoa/thường; chuỗi và chú thích so chính xác.
Kết quả mong đợi: `KHỚP: module McpFast trong project msmcp_<sha8> trùng N dòng với vba/McpFast.bas`.

## Lưu ý
- Lỗi lúc chạy trong VBA có thể làm trình soạn VBA dừng ở chế độ `[break]`; server tự gửi
  `vba execute VBE.CommandBars.FindControl(, 228).Execute` (Run › Reset) rồi báo lỗi.
- Tắt hẳn bộ máy VBA: đặt biến môi trường `MSMCP_DISABLE_VBA=1` → chỉ dùng COM (chậm, chỉ đọc).
- Xóa: đóng MicroStation rồi xóa thư mục `%LOCALAPPDATA%\msmcp\`.
