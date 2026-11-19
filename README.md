# MicroStation V8i MCP — đọc/ghi và KCS bản đồ địa chính (TT 26/2024)

MCP server cho phép trợ lý AI đọc, ghi và kiểm tra (KCS) bản vẽ MicroStation V8i theo
Thông tư 26/2024/TT-BTNMT (Phụ lục 21 phân lớp, Phụ lục 22 ký hiệu), đã cập nhật TT 23/2025.

## Kiến trúc

```
Claude ──MCP stdio──> Python server ──COM (điều khiển)──> MicroStation V8i
                           │                                   │
                           └── req.txt / resp.txt ◄── VBA McpFast.bas chạy TRONG ustation.exe
```

- **Bộ máy đọc/ghi nhanh = VBA in-process** (`vba/McpFast.bas`, mã nguồn chữ, ASCII).
  Vòng lặp quét chạy bên trong MicroStation, không qua ranh giới tiến trình:
  đọc ~27.000 đối tượng kèm hình học trong ~1 giây (COM thuần: > 3 phút).
  Lần đầu chạy, server dựng project `%LOCALAPPDATA%\msmcp\msmcp_<hash>.mvba` bằng key-in chuẩn
  (`vba create` / `VBComponents.Import` / `vba save`). Đổi mã → hash đổi → project mới.
- **Kiểm chứng an toàn**: repo không chứa file nhị phân nào; `.mvba` được dựng trên máy bạn từ
  `vba/McpFast.bas`. Chi tiết các bước dựng, giao thức và phạm vi quyền: [docs/vba-engine.md](docs/vba-engine.md).
  Tự đối chiếu mã đang chạy trong MicroStation với nguồn: `python scripts/vba_engine.py verify`.
- **COM dự phòng**: nếu VBA không dùng được (hoặc đặt `MSMCP_DISABLE_VBA=1`) thì đọc qua COM
  (cùng định dạng dữ liệu, chậm). Ghi bắt buộc dùng VBA.
- Tất cả lời gọi COM chạy trên **một luồng STA** duy nhất → gọi song song an toàn.
- **Không tự khởi động MicroStation**, không ghi registry, không tự cập nhật, không tải gì từ mạng.
- Không dùng DLL/MDL/file nhị phân dựng sẵn nào.

## Công cụ

| Tool | Việc |
|---|---|
| `ms_status` | Kết nối, file đang mở, bộ máy (vba/com) |
| `ms_summary(file?)` | Thống kê toàn file theo level: loại, màu (chỉ số=RGB), nét, mẫu chữ, cell, độ dài; đối chiếu PL21 |
| `ms_query(...)` | Lọc theo level, loại, vùng bbox, chữ/regex, cell, màu; phân trang; trả đỉnh, độ dài, diện tích |
| `ms_get(ids)` | Chi tiết đầy đủ theo id |
| `ms_edit(commands, dry_run=True)` | Tạo line/linestring/shape/text/cell; set level/màu/nét; move; text_set; delete; join (nối thành vùng/chuỗi) |
| `ms_undo` | Hoàn tác cả lô ms_edit gần nhất |
| `kcs_check(file?, checks?)` | KCS: phân lớp, màu/nét, topology, dựng thửa + nhãn + diện tích; xuất CSV |
| `ms_view(action)` | fit / zoom_window / zoom_to / zoom_ids + ảnh vùng đó (dựng từ dữ liệu, tô đậm đối tượng) |
| `ms_render(file?, bbox?, levels?, highlight?)` | Dựng ảnh từ dữ liệu, không đổi khung nhìn; dùng được cho file mở ngầm |
| `ms_copy_from(file, levels?...)` | Chép đối tượng từ file khác vào file đang mở, xử lý trong server (302 đường trong 0,4 s), bỏ trùng |
| `ms_screenshot` | Chụp cửa sổ thật (chỉ đúng khi MicroStation đang hiện trên màn hình) |
| `ms_open`, `ms_save` | Mở file lên cửa sổ / lưu |
| `ms_keyin(cmd)` | Key-in (chặn vba, mdl, delete, save, exit, ...) |

`file=` trong các tool đọc: mở ngầm **chỉ đọc** file khác, không đổi file đang hiển thị
(giới hạn V8i: file mở ngầm không đọc được tên/thuộc tính level).

### An toàn khi ghi
- Chỉ ghi file đang mở; `dry_run=True` mặc định (xem trước, kiểm tra id và level tồn tại).
- Lần ghi đầu vào mỗi file trong phiên: sao lưu sang `%LOCALAPPDATA%\msmcp\backups\`.
- Mỗi lô đặt điểm undo (`mark`) → `ms_undo` hoàn tác cả lô.
- Lệnh tạo không chỉ định ký hiệu: màu theo nhóm PL22 của level (đen 0 / xanh 2 / đỏ 3), nét 0, kiểu nét 0.
- Nếu mã VBA gặp lỗi và dừng ở chế độ [break], server tự gửi Reset (control 228) và báo lỗi.

## KCS (`kcs_check`)

| Nhóm | Kiểm tra | Căn cứ |
|---|---|---|
| levels | Level ngoài PL21, level đã bãi bỏ (44, 45), dạng đối tượng sai lớp | PL21, TT 23/2025 |
| symbology | Màu sai nhóm đen/xanh/nâu/đỏ; tường nhà nét liền | PL22 I.3, I.4b |
| topology | Điểm treo, hở < 0,3 m, trùng đường, cắt nhau không có nút, tự cắt, đoạn < 1 cm | 16.1b |
| parcels | Dựng thửa từ ranh (10, 21, 23, 31, 32, 38 + địa giới); thiếu/trùng số thửa; thiếu nhãn loại đất/diện tích; mã loại đất lạ; nhãn diện tích ≠ diện tích hình học làm tròn 0,1 m²; dãy số thửa thiếu | Điều 13, 13.5b |

Ngưỡng, bảng level, màu, mã loại đất nằm trong `rules/tt26.json` (sửa được).
Danh mục mã loại đất là tổng hợp TT 27/2018 + TT 08/2024 — **cần đối chiếu lại TT 08/2024**.

## Cài đặt

```powershell
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
claude mcp add microstation -- "<repo>\.venv\Scripts\python.exe" -X utf8 "<repo>\src\main.py"
```

Yêu cầu: Windows, MicroStation V8i đang chạy và đã mở file .dgn, VBA của MicroStation bật (mặc định).
