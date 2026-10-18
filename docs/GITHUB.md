# Gợi ý đưa repo lên GitHub

## Tên repo

`SoilData2AI`

Tên thay thế nếu muốn nhấn mạnh TCVN3:

- `soildata2ai`
- `soil-tcvn3-to-csv`
- `geotech-data2ai`

## Mô tả ngắn (GitHub Description)

**Tiếng Việt:**

> Python + Qt6 chuyển dữ liệu chỉ tiêu cơ lý đất từ Excel/TCVN3 sang CSV UTF-8, chuẩn bị dữ liệu huấn luyện AI; kèm GUI, EXE, test và slide ERSD 2026.

**English:**

> Python/Qt6 pipeline that converts geotechnical soil-property spreadsheets (TCVN3) into UTF-8 CSV for AI training, with GUI, Windows EXE, tests and ERSD 2026 slides.

## Topics

```text
python
excel
xlsx
xls
csv
tcvn3
vietnamese
unicode
data-normalization
geotechnical-engineering
soil-mechanics
machine-learning
data-preparation
pyside6
qt6
pyinstaller
```

## Lệnh đẩy lên GitHub

```bash
cd "D:\bo mon\Nghien cuu khoa hoc\bai bao Python\SoilData2AI"
git init
git add .
git commit -m "Initial commit: SoilData2AI - Excel/TCVN3 to CSV UTF-8 for AI"
git branch -M main
git remote add origin https://github.com/longnhgsvn/SoilData2AI.git
git push -u origin main
```

## Gợi ý cấu trúc Release

Tạo release `v2.0.0` và đính kèm:

- `Xlsx2CSV_Converter.exe` (build từ `scripts/build_exe.ps1`)
- `ERSD2026_NgoThiPhuongThao_Python_DataAI.pdf`
- Source code ZIP do GitHub tự tạo

**Release notes mẫu:**

```markdown
## SoilData2AI v2.0.0 — ERSD 2026

- Batch Excel (.xls/.xlsx) → CSV UTF-8, mỗi sheet một CSV.
- Nhận diện TCVN3 theo font + mẫu ký tự; chuyển Unicode bằng thuật toán 2 pha.
- Gộp AllData.csv bằng normalize.py.
- GUI Qt6 + Windows EXE (~54 MB).
- 12 test tự động; kiểm chứng trên 166 tệp, 962 CSV, 20.352 dòng.
- Kèm 19 slide báo cáo và bộ 20 câu hỏi phản biện.

Lưu ý: repo không chứa dữ liệu thật. Dùng `examples/` để chạy thử.
```

## Checklist trước khi public

- [x] URL repo đã cập nhật: `https://github.com/longnhgsvn/SoilData2AI`.
- [ ] Kiểm tra không có `dataXLSX/`, `dataCSV*`, `AllData.csv`, log hoặc file dữ liệu thật.
- [ ] Chạy `python -m unittest discover -s tests -v`.
- [ ] Chạy `python src/gui.py` và `python src/xlsx2csv.py --input examples --output examples/out`.
- [ ] Kiểm tra LICENSE, CITATION.cff, README.
- [ ] Tạo GitHub Release kèm EXE/PDF (không commit EXE vào repo).


