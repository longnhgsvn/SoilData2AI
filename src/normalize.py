# -*- coding: utf-8 -*-
"""
Chuẩn hóa CSV (đã xuất từ Xlsx2CSV.py) thành một bảng dọc AllData.csv.

- Đọc toàn bộ file (mặc định) từ thư mục dataCSV_UTF8; có thể giới hạn bằng --max-files.
- Chỉ giữ file dạng bảng ngang: có khối tiêu đề kiểu báo cáo cơ lý với dòng chứa
  "Thành phần hạt" / "Độ ẩm" / "Chiều"; mỗi mẫu là một dòng (cột 1 = STT).
- Bỏ qua file dạng dọc / quá hẹp / rác (metadata Book1...).
- Tên cột = ghép nhiều dòng tiêu đề (làm sạch).
- Cột đầu: ten_file_nguon, workbook_key (tên workbook bỏ phần sheet), sheet,
  sau đó stt, so_hieu_mau, ma_ho_khoan, do_sau_tu, do_sau_den, rồi các chỉ tiêu.
"""
from __future__ import annotations

__version__ = "2.0.0"

import argparse
import csv
import re
import sys
from pathlib import Path
from typing import Any, Optional

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# --- Nhận diện bảng ngang (báo cáo thí nghiệm đất tiêu chuẩn) ---
_HORIZ_PAT = re.compile(
    r"thành\s*phần\s*hạt|độ\s*ẩm|chiều\s*sâu|dung\s*trọng|"
    r"lỗ\s*rỗng|góc\s*ma\s*sát|lực\s*dính|cường\s*độ\s*kháng\s*cắt",
    re.IGNORECASE,
)

# Cột góc (φ, góc ma sát trong, góc nghiêng cát…) — giá trị thường bị Excel/ghi CSV
# thành số nguyên "DDMM" (7°08′ → 708) hoặc dạng "19°17′", "19o17"
_ANGLE_COL_PAT = re.compile(
    r"góc|goc|angle|friction|ma_sát|ma_sat|nghiêng|nghieng|"
    r"rest_angle|_\bj\b|\bphi\b",
    re.IGNORECASE,
)
_DMS_TEXT_PAT = re.compile(
    r"^\s*(\d{1,2})\s*[°ºoO]\s*(\d{1,2})\s*(?:['′′′]|phút|phut)?\s*$",
    re.UNICODE | re.IGNORECASE,
)
_TIME_LIKE_PAT = re.compile(r"^\s*(\d{1,2})\s*:\s*(\d{1,2})\s*$")


def split_workbook_sheet(stem: str) -> tuple[str, str]:
    if "_" not in stem:
        return stem, ""
    base, sheet = stem.rsplit("_", 1)
    return base, sheet


def load_csv_rows(path: Path, max_rows: int = 500) -> list[list[str]]:
    rows: list[list[str]] = []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f)
        for i, row in enumerate(r):
            rows.append([c.strip() for c in row])
            if i >= max_rows - 1:
                break
    return rows


def is_junk(rows: list[list[str]]) -> bool:
    if not rows:
        return True
    head = " | ".join(rows[0][:5]).lower()
    if "book1" in head and ".xls" in head:
        return True
    w = max((len(r) for r in rows[:20]), default=0)
    if w < 12:
        return True
    return False


def find_horizontal_anchor(rows: list[list[str]], scan: int = 45) -> Optional[int]:
    for i, row in enumerate(rows[:scan]):
        line = " ".join(row).lower()
        hit = 0
        if "thành phần hạt" in line or "thanh phan hat" in line:
            hit += 1
        if "độ ẩm" in line or "do am" in line:
            hit += 1
        if "chiều" in line and "sâu" in line:
            hit += 1
        if hit >= 2 or ("thành phần hạt" in line and "độ ẩm" in line):
            return i
        if _HORIZ_PAT.search(line) and len([c for c in row if c]) >= 8:
            return i
    return None


def pad_row(row: list[str], n: int) -> list[str]:
    if len(row) >= n:
        return row[:n]
    return row + [""] * (n - len(row))


def merge_header_strings(rows: list[list[str]], start: int, depth: int, ncols: int) -> list[str]:
    names: list[str] = []
    block = rows[start : start + depth]
    for j in range(ncols):
        parts: list[str] = []
        for r in block:
            if j < len(r) and r[j]:
                t = re.sub(r"\s+", " ", r[j]).strip()
                if t and t not in parts:
                    parts.append(t)
        raw = " ".join(parts).strip()
        if not raw:
            raw = f"col_{j+1}"
        names.append(raw)
    return names


def looks_like_data_row(row: list[str]) -> bool:
    if len(row) < 6:
        return False
    if row[1].isdigit() and (row[2].isdigit() or re.match(r"^[A-Za-z0-9]{2,}$", row[2])):
        return True
    return False


def normalize_one_file(path: Path) -> Optional[tuple[list[str], list[list[str]]]]:
    rows = load_csv_rows(path)
    if is_junk(rows):
        return None
    anchor = find_horizontal_anchor(rows)
    if anchor is None:
        return None
    w = max((len(r) for r in rows[anchor : anchor + 12]), default=0)
    if w < 20:
        return None

    header_depth = min(5, len(rows) - anchor)
    for d in (5, 4, 3):
        if anchor + d <= len(rows):
            header_depth = d
            break

    names = merge_header_strings(rows, anchor, header_depth, w)
    data_start = anchor + header_depth
    # Bỏ các dòng trống / đơn vị rời sau khối tiêu đề
    while data_start < len(rows) and not looks_like_data_row(pad_row(rows[data_start], w)):
        data_start += 1
        if data_start - anchor > header_depth + 5:
            break

    data_rows: list[list[str]] = []
    for r in rows[data_start:]:
        r = pad_row(r, w)
        if not looks_like_data_row(r):
            continue
        if not any(c for c in r[6:]):  # không có chỉ số phía sau metadata
            continue
        data_rows.append(r)

    if not data_rows:
        return None

    return names, data_rows


def is_angle_metric_column(col_key: str) -> bool:
    return bool(_ANGLE_COL_PAT.search(col_key))


def _fmt_decimal_deg(x: float) -> str:
    if x != x:  # NaN
        return ""
    ax = abs(x)
    r = round(x, 8)
    if abs(r - round(r)) < 1e-6:
        ir = int(round(r))
        if abs(ir) < 1_000_000:
            return str(ir)
    s = f"{r:.8f}".rstrip("0").rstrip(".")
    return s if s else "0"


def _compact_int_to_decimal_degrees(n: int) -> Optional[float]:
    """7°08′ → 708; 19°17′ → 1917; 16°42′ → 1642."""
    if n < 100:
        return None
    if n > 9999:
        return None
    deg, minutes = divmod(n, 100)
    if minutes >= 60:
        return None
    if deg > 90:
        return None
    if deg == 0 and minutes == 0:
        return None
    return float(deg) + float(minutes) / 60.0


def normalize_angle_cell_value(val: Any) -> str:
    """Đưa giá trị ô góc về độ thập phân (chuỗi), nếu nhận dạng được."""
    if val is None:
        return ""
    if isinstance(val, str):
        s = val.strip()
        if not s:
            return ""
        m = _DMS_TEXT_PAT.match(s)
        if m:
            d, mn = int(m.group(1)), int(m.group(2))
            if mn < 60:
                return _fmt_decimal_deg(d + mn / 60.0)
        m = _TIME_LIKE_PAT.match(s)
        if m:
            d, mn = int(m.group(1)), int(m.group(2))
            if mn < 60 and d < 24:
                return _fmt_decimal_deg(d + mn / 60.0)
        s2 = s.replace(",", ".")
        low = s2.lower()
        if "o" in low and re.match(r"^\d{1,2}\s*o\s*\d{1,2}$", low):
            parts = re.split(r"\s*o\s*", low, maxsplit=1)
            d, mn = int(parts[0]), int(parts[1])
            if mn < 60:
                return _fmt_decimal_deg(d + mn / 60.0)
        try:
            x = float(s2)
        except ValueError:
            return val.strip()
        return _normalize_angle_numeric(x)
    if isinstance(val, (int, float)):
        return _normalize_angle_numeric(float(val))
    return str(val)


def _normalize_angle_numeric(x: float) -> str:
    if x != x:
        return ""
    # Excel: thời gian 7:08 lưu dạng phần ngày 0..1 → độ (giờ) = x*24 = 7 + 8/60
    if 0 < x < 1:
        deg_from_time = x * 24.0
        if deg_from_time <= 90:
            return _fmt_decimal_deg(deg_from_time)
    if abs(x - round(x)) < 1e-9:
        n = int(round(x))
        conv = _compact_int_to_decimal_degrees(n)
        if conv is not None:
            return _fmt_decimal_deg(conv)
    return _fmt_decimal_deg(x)


def slug_header(h: str, used: dict[str, int]) -> str:
    s = re.sub(r"\s+", " ", h).strip()
    s = s.replace("/", "_")
    key = re.sub(r"[^\w\d]+", "_", s, flags=re.UNICODE).strip("_").lower()
    if not key:
        key = "cot"
    if key in used:
        used[key] += 1
        key = f"{key}_{used[key]}"
    else:
        used[key] = 0
    return key


def main(argv: list[str]) -> int:
    here = Path(__file__).resolve().parent
    root = here.parent

    ap = argparse.ArgumentParser(description="Gộp CSV ngang -> AllData.csv")
    ap.add_argument(
        "--input",
        default=str(root / "dataCSV_UTF8"),
        help="Thư mục CSV đầu vào",
    )
    ap.add_argument(
        "--output",
        default=str(root / "dataCSV_UTF8" / "AllData.csv"),
        help="File CSV đầu ra",
    )
    ap.add_argument(
        "--max-files",
        type=int,
        default=0,
        help="Tối đa số file xử lý (mặc định 0 = không giới hạn)",
    )
    args = ap.parse_args(argv)

    in_dir = Path(args.input).resolve()
    out_path = Path(args.output).resolve()

    if not in_dir.is_dir():
        print(f"Không có thư mục: {in_dir}", file=sys.stderr)
        return 2

    files = sorted(in_dir.glob("*.csv"))
    files = [f for f in files if f.name.lower() != "alldata.csv"]
    if args.max_files > 0:
        files = files[: args.max_files]

    all_keys: list[str] = []
    key_set: set[str] = set()
    out_rows: list[dict[str, Any]] = []

    # Thứ tự: nguồn → STT → số hiệu mẫu / TN (cùng cột nguồn nếu file chỉ có một mã)
    # → độ sâu → mã hố khoan.
    meta_fixed = [
        "ten_file_nguon",
        "workbook_key",
        "sheet",
        "stt",
        "so_hieu_mau",
        "so_hieu_thi_nghiem",
        "do_sau_tu",
        "do_sau_den",
        "ma_ho_khoan",
    ]

    for fp in files:
        parsed = normalize_one_file(fp)
        if not parsed:
            continue
        header, drows = parsed
        stem = fp.stem
        wb, sh = split_workbook_sheet(stem)

        used_slug: dict[str, int] = {}
        col_keys = [slug_header(h, used_slug) for h in header]
        for key in col_keys[6:]:
            if key not in key_set:
                key_set.add(key)
                all_keys.append(key)

        for raw in drows:
            if len(raw) < 6:
                continue
            stt = raw[1]
            tn = raw[2]
            ma_ho = raw[3]
            tu = raw[4]
            den = raw[5]
            rec: dict[str, Any] = {
                "ten_file_nguon": fp.name,
                "workbook_key": wb,
                "sheet": sh,
                "stt": stt,
                "so_hieu_mau": tn,
                "so_hieu_thi_nghiem": tn,
                "do_sau_tu": tu,
                "do_sau_den": den,
                "ma_ho_khoan": ma_ho,
            }
            for j, key in enumerate(col_keys):
                if j < 6:
                    continue
                val = raw[j] if j < len(raw) else ""
                if is_angle_metric_column(key):
                    val = normalize_angle_cell_value(val)
                rec[key] = val
            out_rows.append(rec)

    fieldnames = meta_fixed + all_keys
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        w.writeheader()
        for rec in out_rows:
            row = {k: rec.get(k, "") for k in fieldnames}
            w.writerow(row)

    print(
        f"OK: {len(out_rows)} dòng mẫu, {len(files)} file đã xét, cột chỉ tiêu: {len(all_keys)}"
        + (" (max-files=0: toàn bộ .csv)" if args.max_files == 0 else "")
    )
    print(f"Ghi: {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
