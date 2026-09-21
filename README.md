# CheckList Dự án đầu tư — bản Python

Quét toàn bộ cây thư mục của một dự án và xuất **báo cáo Excel** với đúng bố cục
`Info / ListFolder / ListFile` như file `CheckList_Dự_án_đầu_tư_v4.xlsm`.

Phần **chạy** là Python. Phần **kết quả** vẫn là Excel.

---

## 1. Cài đặt (làm một lần)

1. Cài Python 3.9 trở lên: <https://www.python.org/downloads/> — khi cài nhớ
   tích **“Add Python to PATH”**.
2. Mở thư mục chứa dự án này, gõ vào thanh địa chỉ `cmd` rồi Enter, sau đó chạy:

```bat
pip install -r requirements.txt
```

## 2. Cấu hình

Mở `config.ini` bằng Notepad, sửa dòng:

```ini
folder_goc = L:\05-Du an dau tu\215DT-Yen Binh
```

Các mục khác: đuôi file cần thống kê, thư mục loại trừ, độ sâu tối đa, có tô màu
đường dẫn hay không.

## 3. Chạy

Nháy đúp **`run.bat`**, hoặc:

```bat
python quet.py
python quet.py --root "L:\05-Du an dau tu\215DT-Yen Binh"
python quet.py --root "L:\..." --out "D:\BaoCao.xlsx"
python quet.py --duoi pdf,docx,xlsx --khong-to-mau
```

Kết quả nằm trong thư mục `ketqua\`, tên dạng
`CheckList_<tên dự án>_<ngày giờ>.xlsx`. Nhấn **Ctrl+C** để dừng giữa chừng —
dữ liệu đã quét vẫn được ghi ra file.

---

## 4. File kết quả

| Sheet | Nội dung |
|---|---|
| **Info** | Folder gốc, số folder/file, thời gian quét, chú giải màu, **log lỗi ở cột N:R** |
| **ListFolder** | STT · đường dẫn đầy đủ · mã phân cấp (`I`, `I.1`, `I.1.3`…) |
| **ListFile** | Dòng tiêu đề folder + các file bên trong |

Cột của **ListFile**: `A` STT/mã (hyperlink) · `B` Nội dung · `C` Ghi chú ·
`D` Mở File · `E` Ngày tạo · `F` Ngày sửa · `H` đường dẫn đầy đủ (ẩn) ·
`I` cờ F1/F0 (ẩn).

Quy ước màu giữ nguyên bản cũ:

- 🟦 `9BC2E6` — folder **có** file hợp lệ
- 🟩 `C6E0BE` — folder **không** có file hợp lệ
- ⬜ `C8C8C8` — *File được tạo mới* (so với lần quét trước)
- 🟨 `FFFF00` — *File đã thay đổi nội dung*

Lần quét trước được ghi vào `.trang_thai_quet.json`; lần quét đầu tiên của một
thư mục gốc mới sẽ không đánh dấu mới/sửa.

---

## 5. Những lỗi của bản VBA đã được xử lý

| Lỗi VBA | Số lần (lần quét 21/09/2026) | Nguyên nhân | Cách xử lý trong Python |
|---|---|---|---|
| **91** Object variable not set | 580 | `For Each ... In folder.SubFolders` (FSO) ném lỗi khi gặp path > 259 ký tự hoặc folder không quyền đọc; `Resume Next` nhảy lại vào chính vòng lặp đã hỏng → lỗi nhân bản | Mỗi folder quét trong một khối `try/except` riêng; lỗi ghi log **một dòng** rồi bỏ qua đúng folder đó |
| **92** For loop not initialized | 195 | hệ quả trực tiếp của lỗi trên | như trên |
| **70** Permission denied | 3 | folder bị chặn quyền đọc | bắt `PermissionError`, ghi log, quét tiếp |
| **28** Out of stack space | 1 | đệ quy không giới hạn | duyệt bằng **stack tường minh**, không đệ quy |
| Đường dẫn > 260 ký tự | — | giới hạn MAX_PATH của FSO | tự thêm tiền tố `\\?\` (kể cả UNC `\\server\share`) |
| Junction/lối tắt vòng | — | duyệt vô tận | bỏ qua reparse point + nhớ `(dev, inode)` đã duyệt + giới hạn độ sâu |
| Vượt 1.048.576 dòng Excel | — | ghi tràn, hỏng file | dừng đúng lúc và ghi cảnh báo vào sheet Info |

Ngoài ra: bỏ file khóa tạm của Office (`~$…`), sắp xếp folder/file theo A→Z
(bản VBA phụ thuộc thứ tự NTFS), hyperlink tự chuyển sang cột `H` khi đường dẫn
> 255 ký tự (Excel không nhận HYPERLINK dài hơn).

> **Lưu ý về “ngày tạo”:** Windows lưu ngày tạo thật nên nhãn *File được tạo mới*
> chính xác. Trên Linux/macOS không có ngày tạo → file mới sẽ bị gán nhãn
> *File đã thay đổi nội dung*. Không ảnh hưởng khi chạy trên Windows.

---

## 6. Cấu trúc mã nguồn

```
quet.py                  điểm chạy, xử lý tham số dòng lệnh
config.ini               cấu hình
run.bat                  nháy đúp để chạy trên Windows
checklist/scanner.py     engine quét (thay LayTenFolder / DuyetThuMuc)
checklist/excel.py       xuất .xlsx (thay GhiRaSheet / ApDinhDang / ToMauDuongDan)
checklist/config.py      đọc config.ini, nhớ mốc thời gian lần quét trước
tests/test_scanner.py    kiểm thử: path dài, không quyền, symlink vòng, độ sâu
```

Chạy kiểm thử:

```bat
python tests\test_scanner.py
```

---

## 7. Còn dùng file .xlsm cũ được không?

Được. File Excel kết quả là file độc lập, không đụng tới `CheckList_..._v4.xlsm`.
Nếu vẫn muốn giữ các sheet `Công việc` / `FORM` của file cũ, chỉ cần copy hai
sheet `ListFolder` và `ListFile` từ file kết quả dán sang.
