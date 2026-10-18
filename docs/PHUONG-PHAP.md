# Phương pháp và kiến trúc xử lý

Tài liệu này mô tả chi tiết quy trình trong `src/xlsx2csv.py` và `src/normalize.py`, tương ứng với mục 2.3 của bài báo ERSD 2026.

## 1. Sơ đồ tổng thể

```text
Thư mục .xls/.xlsx
        │
        ▼
[1] Quét thư mục (discover_workbooks)
        │
        ▼
[2] Đọc workbook
    ├─ .xlsx → openpyxl (data_only=True)
    └─ .xls  → xlrd 1.2.0 (formatting_info=True)
        │
        ▼
[3] Duyệt từng sheet
    ├─ bỏ sheet rỗng
    ├─ nhận diện TCVN3 (font + marker)
    └─ chuyển Unicode (2 pha)
        │
        ▼
[4] Chuẩn hóa cột góc (nếu có)
        │
        ▼
[5] Ghi CSV UTF-8 có BOM
    tên_file_ten_sheet.csv
        │
        ▼
[6] (tùy chọn) normalize.py → AllData.csv
```

## 2. Quét dữ liệu đầu vào

`discover_workbooks(in_dir)`:

- Lấy `*.xlsx` và `*.xls`.
- Hỗ trợ cả tên dạng `ten_file.xls.xlsx` (vì đuôi cuối là `.xlsx`).
- Sắp xếp theo tên, không phân biệt hoa thường.
- Bỏ qua PDF, ảnh quét và định dạng không phải bảng tính.

Tham số CLI:

| Tham số | Ý nghĩa | Mặc định |
|---|---|---|
| `--input` | Thư mục chứa `.xls/.xlsx` | `<repo>/dataXLSX` |
| `--output` | Thư mục ghi CSV | `<repo>/dataCSV` |

Đường dẫn được chuẩn hóa bằng `pathlib.Path`. Nếu thư mục input không tồn tại, chương trình trả mã lỗi `2` và in thông báo.

## 3. Đọc workbook

### 3.1. `.xlsx` — openpyxl

```python
wb = openpyxl.load_workbook(path, data_only=True)
```

- `data_only=True`: lấy **giá trị đã tính** của ô, không tự đánh giá công thức.
- Nếu công thức chưa được Excel tính và lưu cache, giá trị trả về có thể là `None`.
- Duyệt `ws.max_row`, `ws.max_column`; bỏ sheet không có dữ liệu.

### 3.2. `.xls` — xlrd 1.2.0

```python
book = xlrd.open_workbook(path, formatting_info=True)
```

- `formatting_info=True` cần thiết để truy ra **font name** của từng ô.
- xlrd ≥ 2.0 đã bỏ hỗ trợ `.xls`; dự án pin `xlrd==1.2.0`.
- Với mỗi ô, lấy `xf_list -> font_list -> name` để kiểm tra font TCVN3.

## 4. Nhận diện TCVN3

Hàm `looks_like_tcvn3(text)` kết hợp hai tín hiệu:

### 4.1. Theo font

`is_tcvn3_font(font_name)` trả `True` nếu:

- tên font bắt đầu bằng `.vn` (`.VnTime`, `.VnArial`, `.VnCourier`…), hoặc
- tên font chứa `tcvn3`.

### 4.2. Theo nội dung

TCVN3 dùng các ký tự Latin/ký hiệu không giống Unicode. Ví dụ:

```text
TCVN3 : b¸o c¸o kÕt qu¶ thÝ nghiÖm ®Êt
Unicode: báo cáo kết quả thí nghiệm đất
```

Dự án phân biệt:

- **Marker mạnh** — gần như không xuất hiện trong văn bản Unicode tiếng Việt:  
  `¸ © · ¡ § ¨ ® ¬ ­ ¹ ¶ » ¼ ½ ¾ Æ ¢ £ ¤ ¥ ¦ « ª`
- **Marker yếu** — có thể trùng với ký tự Latin/Unicode:  
  `Ì Ð Ý Ó Ê …`

Quy tắc:

1. Có ít nhất một marker mạnh → coi là TCVN3.
2. Không có marker mạnh, nhưng có **≥ 2 marker yếu** → coi là TCVN3.
3. Ngược lại → giữ nguyên.

Nhờ vậy, các chuỗi Unicode hợp lệ như `núi`, `báo cáo`, `hạt 10 µm` không bị chuyển nhầm.

## 5. Chuyển mã hai pha

Bảng ánh xạ có cả ký tự đơn và cụm hai ký tự:

```python
("¢", "Â")
("¢Ê", "Ấ")
("Aµ", "À")
("µ", "à")
```

Nếu thay trực tiếp theo thứ tự sai, `¢` có thể bị thay trước và làm mất cụm `¢Ê`. Thuật toán dùng placeholder:

```python
# Pha 1
for i, (src, dst) in enumerate(pairs):
    text = text.replace(src, f"::{i}::")

# Pha 2
for i, (src, dst) in enumerate(pairs):
    text = text.replace(f"::{i}::", dst)
```

Danh sách `pairs` được sắp xếp **giảm dần theo độ dài chuỗi nguồn** để ưu tiên cụm dài. Thứ tự cụ thể được kiểm chứng bằng thực nghiệm, không theo alphabet.

> Lưu ý: nếu dữ liệu gốc chứa đúng chuỗi dạng `::i::`, placeholder có thể va chạm. Với dữ liệu đất hiện tại, rủi ro này rất thấp; có thể đổi sang ký tự private-use nếu cần.

## 6. Chuẩn hóa cột góc

Một số báo cáo ghi góc ma sát trong dạng số nguyên `DDMM`:

- `708` → 7°08′ → `7 + 8/60 = 7.13333333`
- `1642` → 16°42′ → `16.7`

Hàm `apply_phi_degrees_decimal(rows)`:

- Nhận diện cột có tiêu đề chứa `góc ma sát trong`, `j`, `góc nghiêng`, `ac`…
- Bắt đầu từ hàng dữ liệu đầu tiên (có STT).
- Chuyển số nguyên DDMM, chuỗi `19°17′`, `19:17`, phân số ngày của Excel sang độ thập phân.

Áp dụng cho cả `.xlsx` và `.xls`.

## 7. Ghi CSV

- Mỗi sheet có dữ liệu → một tệp CSV.
- Tên: `ten_file_goc_ten_sheet.csv`.
- Tên sheet được làm sạch ký tự Windows (`\ / : * ? " < > |`), bỏ ký tự điều khiển, cắt khoảng trắng/dấu chấm cuối.
- Nếu hai sheet trùng tên sau khi làm sạch, thêm hậu tố `_2`, `_3`…
- Mã hóa: `utf-8-sig` (UTF-8 có BOM).
- Ô `None` ghi thành ô trống.
- Giữ nguyên số hàng/cột và cấu trúc lưới của sheet gốc.

## 8. Xử lý lỗi

- Lỗi ở một sheet: ghi `WARN`, bỏ qua sheet đó, tiếp tục các sheet còn lại.
- Lỗi ở một workbook: ghi `FAIL`, bỏ qua workbook, tiếp tục các tệp còn lại.
- Tệp hỏng ZIP (`.xlsx`) hoặc có mật khẩu (`.xls`) được ghi log để xử lý thủ công.
- Kết thúc: in tổng số CSV và danh sách tệp lỗi.

## 9. Gộp AllData.csv

`src/normalize.py` đọc các CSV đã tạo và gộp thành bảng dọc:

- Chỉ giữ các file có dạng bảng ngang báo cáo thí nghiệm (có anchor như `Thành phần hạt`, `Độ ẩm`, `Chiều sâu`…).
- Mỗi dòng mẫu phải có STT và số hiệu mẫu/hố khoan.
- Tên cột = ghép các dòng tiêu đề, làm sạch thành slug.
- Thêm metadata: `ten_file_nguon`, `workbook_key`, `sheet`, `stt`, `so_hieu_mau`, `do_sau_tu`, `do_sau_den`, `ma_ho_khoan`.
- Chuẩn hóa cột góc về độ thập phân.
- Xuất `AllData.csv` UTF-8 có BOM.

Tham số:

| Tham số | Ý nghĩa | Mặc định |
|---|---|---|
| `--input` | Thư mục CSV | `<repo>/dataCSV_UTF8` |
| `--output` | Tệp AllData.csv | `<repo>/dataCSV_UTF8/AllData.csv` |
| `--max-files` | Giới hạn số CSV | `0` (toàn bộ) |

## 10. Quan hệ với bài báo

| Bài báo | Mã nguồn |
|---|---|
| `Xlsx2CSV.py` | `src/xlsx2csv.py` |
| `Normalize.Py` | `src/normalize.py` |
| Giao diện dòng lệnh | `main()` trong `src/xlsx2csv.py` |
| GUI Qt6 (bổ sung) | `src/gui.py` |
| Kiểm chứng | `tests/` |

## 11. Hạn chế đã biết

- Chưa có schema cột thống nhất; CSV vẫn là dữ liệu lưới.
- Chưa kiểm tra đơn vị, missing, outlier, trùng lặp.
- Chưa hỗ trợ VNI/VIQR.
- Chưa xử lý ảnh, biểu đồ, cell style trong workbook.
- `.xls` phụ thuộc `xlrd==1.2.0`; cần kiểm tra khi nâng Python.
- Hiệu năng tuần tự; chưa song song hóa khi số tệp rất lớn.
