# TCVN3 → Unicode trong SoilData2AI

## 1. Vì sao cần chuyển mã?

TCVN3 là bảng mã 8-bit dùng cho tiếng Việt trước Unicode. Văn bản TCVN3 thường đi kèm font `.VnTime`, `.VnArial`, `.VnCourier`… Nếu mở file bằng công cụ hiện đại mà không chuyển mã, các byte TCVN3 bị hiểu là ký tự Latin, dẫn đến sai dấu hoặc mất dấu.

Ví dụ:

| TCVN3 | Unicode |
|---|---|
| `b¸o c¸o kÕt qu¶ thÝ nghiÖm ®Êt` | `báo cáo kết quả thí nghiệm đất` |
| `C«ng tr×nh` | `Công trình` |
| `§é Èm tù nhiªn` | `Độ ẩm tự nhiên` |
| `Gãc ma s¸t trong` | `Góc ma sát trong` |

## 2. Bảng ánh xạ

Bảng đầy đủ nằm trong `src/xlsx2csv.py`, biến `_TCVN3_UNICODE_PAIRS`. Một số cặp tiêu biểu:

| Nguồn | Đích | Ghi chú |
|---|---|---|
| `Aµ` | `À` | cụm |
| `A¸` | `Á` | cụm |
| `¢` | `Â` | ký tự đơn |
| `¢Ê` | `Ấ` | cụm, phải thay trước `¢` |
| `EÌ` | `È` | cụm |
| `£` | `Ê` | ký tự đơn |
| `µ` | `à` | ký tự đơn |
| `¸` | `á` | ký tự đơn |
| `¡` | `Ă` | ký tự đơn |
| `§` | `Đ` | ký tự đơn |
| `¥` | `Ơ` | ký tự đơn |
| `¦` | `Ư` | ký tự đơn |
| `¨` | `ă` | ký tự đơn |
| `®` | `đ` | ký tự đơn |

## 3. Thuật toán hai pha

### Vì sao không thay trực tiếp?

Giả sử cần chuyển `¢Ê` → `Ấ`. Nếu thay `¢` trước:

```text
¢Ê  →  ÂÊ  →  ÂÊ   (sai, mất cụm)
```

Nếu thay `¢Ê` trước rồi `¢`:

```text
¢Ê  →  Ấ   (đúng)
```

Nhưng với dữ liệu lớn, việc kiểm soát thứ tự thủ công dễ sai. Hai pha dùng placeholder để cô lập:

```python
# pairs đã sắp xếp giảm dần theo len(src)
for i, (src, dst) in enumerate(pairs):
    text = text.replace(src, f"::{i}::")

for i, (src, dst) in enumerate(pairs):
    text = text.replace(f"::{i}::", dst)
```

Ví dụ:

```text
b¸o c¸o
→ b::0::o c::0::o
→ báo cáo
```

## 4. Nhận diện khi nào áp dụng

Chỉ chuyển khi:

1. Font của ô là `.Vn…` hoặc chứa `TCVN3`, **hoặc**
2. Chuỗi có marker mạnh, **hoặc**
3. Chuỗi có ≥ 2 marker yếu.

Marker mạnh:

```text
¸ © · ¡ § ¨ ® ¬ ­ ¹ ¶ » ¼ ½ ¾ Æ ¢ £ ¤ ¥ ¦ « ª
```

Marker yếu:

```text
Ì Ð Ý Ó Ê …
```

Mục tiêu là tránh chuyển nhầm:

- `núi đất` (Unicode, `ú` là chữ hợp lệ)
- `báo cáo` (Unicode)
- `hạt 10 µm` (đơn vị micro-mét)

## 5. Kiểm thử

```bash
python -m unittest tests.test_core -v
```

Các test liên quan:

- `test_convert_common_sentence`
- `test_unicode_text_is_not_mistaken`
- `test_font_detection`
- `test_convert_and_angle`
- `test_real_xls`

## 6. Giới hạn

- Ánh xạ một chiều, không thể đảo ngược hoàn toàn.
- Chuỗi trộn Unicode + TCVN3 có thể cần kiểm tra thủ công.
- Chưa hỗ trợ VNI, VIQR, Unicode tổ hợp (NFD) hoặc các bảng mã khác.
- Nếu ô chứa `::i::` đúng nghĩa, placeholder có thể va chạm (rất hiếm).
- Cần đối chiếu chuyên gia với các cụm hiếm hoặc dữ liệu ngoài miền đất.

## 7. Mở rộng

- Thay placeholder bằng ký tự private-use (`\uE000…`) để tránh va chạm.
- Thêm bảng VNI/VIQR.
- Chuẩn hóa Unicode NFC sau chuyển mã.
- Ghi log các ô nghi ngờ để review.
- So sánh với `vietnamese-conversion` trên bộ test chuẩn.
