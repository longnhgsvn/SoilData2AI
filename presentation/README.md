# Slide báo cáo ERSD 2026

Bộ slide 19 trang được dựng tự động bằng Python từ template hội nghị.

## Tệp chính

- `ERSD2026_NgoThiPhuongThao_Python_DataAI.pptx` — bản trình chiếu.
- `ERSD2026_NgoThiPhuongThao_Python_DataAI.pdf` — bản preview.
- `contact_sheet.png` — ảnh xem nhanh toàn bộ slide.
- `build_presentation.py` — mã dựng slide bằng `python-pptx`.
- `extract_template_assets.py` — trích nền/logo từ template.
- `templates/Side 60 năm_ERSD 3.pptx` — template gốc.
- `assets/` — nền full-slide và logo ERSD.
- `images/` — ảnh minh họa từ bài báo.

## Dựng lại

Từ thư mục repo:

```bash
cd presentation
python extract_template_assets.py
python build_presentation.py
```

Hoặc chỉ định template/đầu ra:

```bash
python build_presentation.py `
  --template "templates\Side 60 năm_ERSD 3.pptx" `
  --output "ERSD2026.pptx"
```

Yêu cầu: `python-pptx`, `Pillow`.

## Nội dung

1. Trang bìa
2. Nội dung trình bày
3. Đặt vấn đề
4. Mục tiêu
5. Dữ liệu đầu vào (166 tệp)
6. Cơ sở lý thuyết TCVN3 → Unicode
7. Phương pháp 4 bước
8. Pipeline 6 bước
9. Phân tích workbook
10. Kiến trúc mã nguồn
11. Thuật toán chuyển mã 2 pha
12. Xuất CSV & đặt tên
13. Kiểm chứng
14. Kết quả định lượng
15. Kết quả định tính
16. Thảo luận
17. Kết luận
18. Hướng phát triển
19. Cảm ơn / Q&A

Mỗi slide có speaker notes tiếng Việt.

## Lưu ý số liệu

Slide dùng số liệu theo bài báo: **166 tệp · 963 sheet · 20.532 mẫu**.

Khi chạy lại bộ dữ liệu hiện hành: **962 CSV · 20.352 dòng AllData.csv**.

Khi trình bày, nên nói rõ đây là số liệu của lần chạy nào; xem `../docs/KET-QUA-KIEM-CHUNG.md`.
