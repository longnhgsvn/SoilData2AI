# Kết quả kiểm chứng

## 1. Môi trường

| Thành phần | Phiên bản |
|---|---|
| Python | 3.12.0 |
| openpyxl | 3.1.5 |
| xlrd | 1.2.0 |
| PySide6 | 6.11.2 |
| PyInstaller | 6.12.0 |
| OS | Windows 10/11 64-bit |

## 2. Kiểm thử tự động

```bash
python -m unittest discover -s tests -v
```

Kết quả: **12/12 test pass**.

Bao phủ:

- Chuyển TCVN3 → Unicode.
- Không nhầm Unicode có dấu với TCVN3 (`núi`, `µm`).
- Nhận diện font `.Vn…` / `TCVN3`.
- Làm sạch tên tệp Windows.
- Đổi DDMM → độ thập phân.
- Xuất CSV UTF-8 có BOM.
- Bỏ sheet rỗng.
- Không ghi đè khi trùng tên.
- Đọc `.xls` thật.
- Chuẩn hóa `Normalize`.
- Luồng QThread của GUI.

## 3. Chạy toàn bộ dữ liệu

Lệnh:

```bash
python src/xlsx2csv.py --input "D:\...\dataXLSX" --output "D:\...\dataCSV_run"
python src/normalize.py --input "D:\...\dataCSV_run" --output "D:\...\AllData.csv"
```

Kết quả trên snapshot hiện hành:

| Chỉ số | Giá trị |
|---|---:|
| Tệp đầu vào | 166 (158 `.xlsx` + 8 `.xls`) |
| Tệp xử lý thành công | 166 |
| Tệp lỗi | 0 |
| Tệp CSV tạo ra | 962 |
| Dòng `AllData.csv` | 20.352 |
| Cột chỉ tiêu trong AllData | 400 |
| Thời gian chạy | khoảng 1,5–2,5 phút (tùy máy) |

Kiểm tra thêm:

- Tất cả CSV đọc được bằng `utf-8-sig`.
- Không có `WARN`/`FAIL` trong log chạy đầy đủ.
- Chạy lại 2 lần cho cùng kết quả (962 CSV, 20.352 dòng).
- So sánh SHA-256 giữa 2 lần chạy: **0 tệp thay đổi**.

## 4. Đối chiếu với bài báo

Bài báo ghi:

- 166 tệp
- 963 sheet
- 20.532 mẫu

Đây là số liệu của **lần chạy trước**. Trên snapshot dữ liệu hiện hành:

- 962 CSV sheet
- 20.352 dòng

Nguyên nhân có thể:

1. Bộ dữ liệu đã được bổ sung/loại bỏ một sheet.
2. Một sheet trước đây có dữ liệu, nay rỗng (hoặc ngược lại).
3. Thay đổi heuristic nhận diện/normalize giữa các phiên bản.
4. Khác biệt cách đếm: file `AllData.csv` từng bị tính vào số CSV.

**Khuyến nghị:** khi công bố, ghi rõ:

- ngày chạy,
- phiên bản mã nguồn (git commit),
- hash của thư mục dữ liệu,
- số file/sheet/dòng,
- danh sách file lỗi (nếu có).

## 5. Kiểm tra EXE

```powershell
.\dist\Xlsx2CSV_Converter.exe --selftest
.\dist\Xlsx2CSV_Converter.exe --smoke-gui
.\dist\Xlsx2CSV_Converter.exe --cli --input examples --output examples\out --merge
```

Kết quả:

- `--selftest`: exit 0.
- `--smoke-gui`: exit 0, không để lại process.
- `--cli --merge`: exit 0, tạo CSV và `AllData.csv`.
- Dung lượng EXE: **54,3 MB**.
- Không đóng gói `torch`, `pandas`, `scipy`, `cv2`, `numba`…

## 6. Kiểm chứng định tính

Các chuỗi tiếng Việt sau chuyển mã hiển thị đúng:

| Trước | Sau |
|---|---|
| `b¸o c¸o kÕt qu¶ thÝ nghiÖm ®Êt` | `báo cáo kết quả thí nghiệm đất` |
| `C«ng tr×nh` | `Công trình` |
| `§é Èm tù nhiªn` | `Độ ẩm tự nhiên` |
| `Gãc ma s¸t trong` | `Góc ma sát trong` |

## 7. Hạn chế của kiểm chứng

- Chưa có ground truth do chuyên gia gán cho toàn bộ 20 nghìn dòng.
- Chưa đánh giá precision/recall của nhận diện TCVN3 trên tập độc lập.
- Chưa đo chất lượng mô hình AI (ngoài phạm vi bài).
- Chưa kiểm thử trên Linux/macOS.
- Chưa đo hiệu năng với > 10.000 tệp.
