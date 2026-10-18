# -*- coding: utf-8 -*-
"""
SoilData2AI — chuyển bảng tính chỉ tiêu cơ lý đất sang CSV UTF-8.

Quét hàng loạt .xls/.xlsx, tách từng sheet có dữ liệu thành một CSV,
nhận diện TCVN3 (font .Vn... hoặc mẫu ký tự) và chuyển sang Unicode.
Đây là lõi tương ứng với Xlsx2CSV.py trong bài báo ERSD 2026.

CLI:
    python src/xlsx2csv.py --input dataXLSX --output dataCSV_UTF8
"""
from __future__ import annotations

__version__ = "2.0.0"

import argparse
import csv
import re
import sys
import unicodedata
from datetime import datetime, time
from pathlib import Path
from typing import Any, Iterable, Optional

# Góc trong Excel định dạng tùy chỉnh thường đọc ra số nguyên: độ × 100 + phút (2 chữ số phút).
# Ví dụ 7°08′ → 708 (7 độ, 08 phút), 16°42′ → 1642 (16 độ, 42 phút). Công thức: độ = n // 100,
# phút = n % 100 → thập phân = độ + phút/60. Một số ô là kiểu giờ (time) hoặc phân số ngày.
# Script nhận diện cột theo tiêu đề (φ / j, góc nghiêng ac) rồi ghi ra CSV dạng độ thập phân.

# Ensure Vietnamese output works on Windows consoles with legacy codepages.
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


_TCVN3_UNICODE_PAIRS: list[tuple[str, str]] = [
    # Nguồn mapping: vietnamese-conversion (TCVN3 chuẩn VietKey).
    ("Aµ", "À"),
    ("A¸", "Á"),
    ("¢", "Â"),
    ("A·", "Ã"),
    ("EÌ", "È"),
    ("EÐ", "É"),
    ("£", "Ê"),
    ("I×", "Ì"),
    ("IÝ", "Í"),
    ("Oß", "Ò"),
    ("Oã", "Ó"),
    ("¤", "Ô"),
    ("Oâ", "Õ"),
    ("Uï", "Ù"),
    ("Uó", "Ú"),
    ("Yý", "Ý"),
    ("µ", "à"),
    ("¸", "á"),
    ("©", "â"),
    ("·", "ã"),
    ("Ì", "è"),
    ("Ð", "é"),
    ("ª", "ê"),
    ("×", "ì"),
    ("Ý", "í"),
    ("ß", "ò"),
    ("ã", "ó"),
    ("«", "ô"),
    ("â", "õ"),
    ("ï", "ù"),
    ("ó", "ú"),
    ("ý", "ý"),
    ("¡", "Ă"),
    ("§", "Đ"),
    ("IÜ", "Ĩ"),
    ("Uò", "Ũ"),
    ("¥", "Ơ"),
    ("¦", "Ư"),
    ("¨", "ă"),
    ("®", "đ"),
    ("Ü", "ĩ"),
    ("ò", "ũ"),
    ("¬", "ơ"),
    ("­", "ư"),
    ("A¹", "Ạ"),
    ("A¶", "Ả"),
    ("¢Ê", "Ấ"),
    ("¢Ç", "Ầ"),
    ("¢È", "Ẩ"),
    ("¢É", "Ẫ"),
    ("¢Ë", "Ậ"),
    ("¡¾", "Ắ"),
    ("¡»", "Ằ"),
    ("¡¼", "Ẳ"),
    ("¡½", "Ẵ"),
    ("¡Æ", "Ặ"),
    ("EÑ", "Ẹ"),
    ("EÎ", "Ẻ"),
    ("EÏ", "Ẽ"),
    ("£Õ", "Ế"),
    ("£Ò", "Ề"),
    ("£Ó", "Ể"),
    ("£Ô", "Ễ"),
    ("£Ö", "Ệ"),
    ("IØ", "Ỉ"),
    ("IÞ", "Ị"),
    ("Oä", "Ọ"),
    ("Oá", "Ỏ"),
    ("¤è", "Ố"),
    ("¤å", "Ồ"),
    ("¤æ", "Ổ"),
    ("¤ç", "Ỗ"),
    ("¤é", "Ộ"),
    ("¥í", "Ớ"),
    ("¥ê", "Ờ"),
    ("¥ë", "Ở"),
    ("¥ì", "Ỡ"),
    ("¥î", "Ợ"),
    ("Uô", "Ụ"),
    ("Uñ", "Ủ"),
    ("¦ø", "Ứ"),
    ("¦õ", "Ừ"),
    ("¦ö", "Ử"),
    ("¦÷", "Ữ"),
    ("¦ù", "Ự"),
    ("Yú", "Ỳ"),
    ("Yþ", "Ỵ"),
    ("Yû", "Ỷ"),
    ("Yü", "Ỹ"),
    ("¹", "ạ"),
    ("¶", "ả"),
    ("Ê", "ấ"),
    ("Ç", "ầ"),
    ("È", "ẩ"),
    ("É", "ẫ"),
    ("Ë", "ậ"),
    ("¾", "ắ"),
    ("»", "ằ"),
    ("¼", "ẳ"),
    ("½", "ẵ"),
    ("Æ", "ặ"),
    ("Ñ", "ẹ"),
    ("Î", "ẻ"),
    ("Ï", "ẽ"),
    ("Õ", "ế"),
    ("Ò", "ề"),
    ("Ó", "ể"),
    ("Ô", "ễ"),
    ("Ö", "ệ"),
    ("Ø", "ỉ"),
    ("Þ", "ị"),
    ("ä", "ọ"),
    ("á", "ỏ"),
    ("è", "ố"),
    ("å", "ồ"),
    ("æ", "ổ"),
    ("ç", "ỗ"),
    ("é", "ộ"),
    ("í", "ớ"),
    ("ê", "ờ"),
    ("ë", "ở"),
    ("ì", "ỡ"),
    ("î", "ợ"),
    ("ô", "ụ"),
    ("ñ", "ủ"),
    ("ø", "ứ"),
    ("õ", "ừ"),
    ("ö", "ử"),
    ("÷", "ữ"),
    ("ù", "ự"),
    ("ú", "ỳ"),
    ("þ", "ỵ"),
    ("û", "ỷ"),
    ("ü", "ỹ"),
]

_TCVN3_TO_UNICODE = sorted(_TCVN3_UNICODE_PAIRS, key=lambda p: len(p[0]), reverse=True)
# Marker "mạnh": gần như không xuất hiện trong văn bản Unicode tiếng Việt.
# Dùng để tránh nhầm các chữ có dấu Unicode (é, è, í, ù, ú, ...) với TCVN3.
_TCVN3_STRONG_MARKERS_RE = re.compile(r"[¸©·¡§¨®¬­¹¶»¼½¾Æ¢£¤¥¦«ª]")
# Marker "yếu": có thể trùng với ký tự Latin/Unicode; chỉ kết luận TCVN3 khi
# xuất hiện từ 2 marker trở lên (hoặc khi font là .Vn.../TCVN3).
_TCVN3_MARKERS_RE = re.compile(r"[µ¸©·¡§ÌÐÝÓ¨®Ü¬­¹¶Ê»¼½Æ]")
_UNICODE_VN_RE = re.compile(
    r"[áàảãạăắằẳẵặâấầẩẫậéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵđĐ]",
    re.IGNORECASE,
)


def looks_like_tcvn3(text: str) -> bool:
    # TCVN3 thường chứa các ký tự đặc thù như µ¸©...; Unicode "chuẩn" sẽ có áàả...
    if not text:
        return False
    # Một marker mạnh là đủ (ví dụ dấu cedilla trong "b¸o").
    if _TCVN3_STRONG_MARKERS_RE.search(text) is not None:
        return True
    # Với marker yếu (Ì, Ð, Ý, Ó, Ê...), yêu cầu ít nhất 2 lần xuất hiện để
    # giảm false positive với văn bản Latin/Unicode.
    return len(_TCVN3_MARKERS_RE.findall(text)) >= 2


def tcvn3_to_unicode(text: str) -> str:
    if not text:
        return text
    # 2-phase replacement để tránh đè lẫn nhau (ví dụ: '¢' và '¢Ê')
    out = text
    for i, (src, _dst) in enumerate(_TCVN3_TO_UNICODE):
        out = out.replace(src, f"::{i}::")
    for i, (_src, dst) in enumerate(_TCVN3_TO_UNICODE):
        out = out.replace(f"::{i}::", dst)
    return out


def is_tcvn3_font(font_name: Optional[str]) -> bool:
    if not font_name:
        return False
    f = font_name.strip().lower()
    # Common TCVN3 fonts in Excel: ".VnTime", ".VnArial", ".VnCourier", ...
    return ("tcvn3" in f) or f.startswith(".vn")


def safe_filename_piece(name: str) -> str:
    """Làm sạch tên sheet để dùng trong tên tệp trên Windows."""
    name = name.strip()
    name = re.sub(r"[\x00-\x1f]+", "", name)
    name = re.sub(r"[\\/:*?\"<>|]+", "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    # Windows không cho tên kết thúc bằng dấu chấm hoặc khoảng trắng.
    name = name.rstrip(" .")
    return name[:180].rstrip(" .") if name else "sheet"


def _sheet_output_path(
    out_dir: Path,
    file_path: Path,
    sheet_title: str,
    used_names: Optional[set[str]] = None,
) -> Path:
    """Tên CSV duy nhất trong một lần chạy: ten_file_ten_sheet.csv."""
    base = f"{file_path.stem}_{safe_filename_piece(sheet_title)}"
    # Giới hạn độ dài và tránh tên thiết bị dành riêng của Windows.
    base = base[:200].rstrip(" .")
    if base.upper() in {"CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "LPT1", "LPT2", "LPT3"}:
        base = "_" + base
    if used_names is None:
        used_names = set()
    name = base
    n = 2
    while name.casefold() in used_names:
        name = f"{base}_{n}"
        n += 1
    used_names.add(name.casefold())
    return out_dir / f"{name}.csv"


def has_any_data(rows: Iterable[Iterable[Any]]) -> bool:
    for r in rows:
        for v in r:
            if v is None:
                continue
            if isinstance(v, str):
                if v.strip() != "":
                    return True
            else:
                return True
    return False


def _fold_accents(s: str) -> str:
    s = unicodedata.normalize("NFD", s)
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower()


_dms_txt_pat = re.compile(
    r"^\s*(\d{1,2})\s*[°ºoO]\s*(\d{1,2})\s*(?:['′′]|phút|phut)?\s*$",
    re.UNICODE | re.IGNORECASE,
)
_time_like_pat = re.compile(r"^\s*(\d{1,2})\s*:\s*(\d{1,2})\s*$")


def _col_header_blob(rows: list[list[Any]], col: int, max_row: int = 16) -> str:
    parts: list[str] = []
    for r in range(min(max_row, len(rows))):
        if col < len(rows[r]):
            v = rows[r][col]
            if v is None:
                continue
            t = str(v).strip()
            if t:
                parts.append(t)
    return " ".join(parts)


def _is_phi_ma_sat_trong_column(label: str) -> bool:
    """Cột góc ma sát trong φ (ô mã j), không gồm góc nghiêng cát (ac)."""
    f = _fold_accents(label)
    if "trong" not in f:
        return False
    # Tránh nhầm cột chỉ có "trong" kiểu khác
    if "ma" in f and "sat" in f:
        return True
    if re.search(r"\bj\b", label) or re.search(r"\bj\b", f):
        return True
    return False


def _is_goc_nghieng_ac_column(label: str) -> bool:
    """Cột góc nghiêng cát khi khô (thường mã ac) — cùng kiểu số DDMM như φ."""
    f = _fold_accents(label)
    if "nghieng" not in f:
        return False
    if re.search(r"\bac\b", f) or "kho" in f:
        return True
    return False


def _excel_integer_dms_to_decimal_degrees(n: int) -> Optional[float]:
    """
    Số nguyên Excel kiểu «độ nối hai chữ số phút»: n = độ*100 + phút.
    708 → 7°08′ → 7 + 8/60; 1642 → 16°42′ → 16 + 42/60.
    """
    if n < 100 or n > 9999:
        return None
    deg, minutes = divmod(n, 100)
    if minutes >= 60 or deg > 90:
        return None
    if deg == 0 and minutes == 0:
        return None
    return float(deg) + float(minutes) / 60.0


def _convert_angle_cell_for_csv(v: Any) -> Any:
    """Một ô góc: số DDMM / chuỗi DMS / time / phân số ngày → độ thập phân (float)."""
    if v is None:
        return None
    if isinstance(v, datetime):
        return round(
            v.hour + v.minute / 60.0 + v.second / 3600.0,
            8,
        )
    if isinstance(v, time):
        return round(
            v.hour + v.minute / 60.0 + v.second / 3600.0,
            8,
        )
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return v
        m = _dms_txt_pat.match(s)
        if m:
            d, mn = int(m.group(1)), int(m.group(2))
            if mn < 60:
                return round(d + mn / 60.0, 8)
        m = _time_like_pat.match(s)
        if m:
            d, mn = int(m.group(1)), int(m.group(2))
            if mn < 60 and d < 24:
                return round(d + mn / 60.0, 8)
        low = s.lower().replace(",", ".")
        if re.match(r"^\d{1,2}\s*o\s*\d{1,2}$", low):
            a, b = low.split("o", 1)
            d, mn = int(a.strip()), int(b.strip())
            if mn < 60:
                return round(d + mn / 60.0, 8)
        try:
            x = float(low)
        except ValueError:
            return v
        return _convert_angle_numeric(x)
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return _convert_angle_numeric(float(v))
    return v


def _convert_angle_numeric(x: float) -> Any:
    if x != x:  # NaN
        return x
    if 0 < x < 1:
        return round(x * 24.0, 8)
    if abs(x - round(x)) < 1e-9:
        n = int(round(x))
        conv = _excel_integer_dms_to_decimal_degrees(n)
        if conv is not None:
            return round(conv, 8)
    return x


def _sheet_ncols(rows: list[list[Any]], scan: int = 25) -> int:
    return max((len(rows[i]) for i in range(min(scan, len(rows)))), default=0)


def _detect_excel_dms_angle_columns(rows: list[list[Any]]) -> set[int]:
    """Cột φ (j) và cột góc nghiêng (ac) cần đổi DDMM → độ thập phân."""
    ncols = _sheet_ncols(rows)
    out: set[int] = set()
    for c in range(ncols):
        blob = _col_header_blob(rows, c)
        if _is_phi_ma_sat_trong_column(blob) or _is_goc_nghieng_ac_column(blob):
            out.add(c)
    return out


def _first_data_row_index(rows: list[list[Any]], ncols: int) -> int:
    """Hàng đầu có STT ở cột 1 (giống Normalize.Py)."""
    for i, row in enumerate(rows):
        if len(row) < 6:
            continue
        a1 = str(row[1]).strip() if len(row) > 1 else ""
        a2 = str(row[2]).strip() if len(row) > 2 else ""
        if a1.isdigit() and (a2.isdigit() or re.match(r"^[A-Za-z0-9]{2,}$", a2)):
            return i
    return 25


def apply_phi_degrees_decimal(rows: list[list[Any]]) -> None:
    """Sửa tại chỗ: các cột góc (φ/j, góc nghiêng ac) từ hàng dữ liệu trở đi."""
    angle_cols = _detect_excel_dms_angle_columns(rows)
    if not angle_cols:
        return
    ncols = max(_sheet_ncols(rows), max(angle_cols) + 1)
    start = _first_data_row_index(rows, ncols)
    for r in range(start, len(rows)):
        row = rows[r]
        for c in angle_cols:
            if c < len(row):
                row[c] = _convert_angle_cell_for_csv(row[c])


def write_csv(path: Path, rows: list[list[Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        for r in rows:
            w.writerow(["" if v is None else v for v in r])


def convert_xlsx(
    file_path: Path,
    out_dir: Path,
    used_names: Optional[set[str]] = None,
) -> int:
    try:
        import openpyxl  # type: ignore
    except Exception as e:  # noqa: BLE001
        raise RuntimeError(
            "Thiếu thư viện openpyxl. Cài bằng: pip install openpyxl"
        ) from e

    wb = openpyxl.load_workbook(file_path, data_only=True)
    written = 0
    try:
        for ws in wb.worksheets:
            try:
                max_row = ws.max_row or 0
                max_col = ws.max_column or 0
                if max_row == 0 or max_col == 0:
                    continue

                rows: list[list[Any]] = []
                for r in range(1, max_row + 1):
                    row: list[Any] = []
                    for c in range(1, max_col + 1):
                        cell = ws.cell(row=r, column=c)
                        v = cell.value
                        if isinstance(v, str):
                            font_name = getattr(cell.font, "name", None)
                            if is_tcvn3_font(font_name) or looks_like_tcvn3(v):
                                v = tcvn3_to_unicode(v)
                        row.append(v)
                    rows.append(row)

                if not has_any_data(rows):
                    continue

                # Đồng bộ với .xls: đổi cột góc dạng DDMM sang độ thập phân.
                apply_phi_degrees_decimal(rows)
                out_path = _sheet_output_path(out_dir, file_path, ws.title, used_names)
                write_csv(out_path, rows)
                written += 1
            except Exception as e:  # noqa: BLE001
                print(
                    f"WARN: {file_path.name} / sheet {ws.title!r} -> {e}",
                    file=sys.stderr,
                )
                continue
    finally:
        try:
            wb.close()
        except Exception:
            pass

    return written


def convert_xls(
    file_path: Path,
    out_dir: Path,
    used_names: Optional[set[str]] = None,
) -> int:
    try:
        import xlrd  # type: ignore
    except Exception as e:  # noqa: BLE001
        raise RuntimeError(
            "Thiếu thư viện xlrd để đọc .xls. Cài bằng: pip install xlrd==1.2.0"
        ) from e

    try:
        book = xlrd.open_workbook(file_path.as_posix(), formatting_info=True)
    except TypeError:
        # xlrd>=2.0.0 removes xls support; user must install 1.2.0
        raise RuntimeError(
            "Phiên bản xlrd hiện tại không hỗ trợ .xls. Hãy cài: pip install xlrd==1.2.0"
        )

    written = 0
    for sheet in book.sheets():
        try:
            if sheet.nrows == 0 or sheet.ncols == 0:
                continue

            rows: list[list[Any]] = []
            for r in range(sheet.nrows):
                row: list[Any] = []
                for c in range(sheet.ncols):
                    v = sheet.cell_value(r, c)
                    if v == "":
                        v = None
                    font_name: Optional[str] = None
                    try:
                        xf = book.xf_list[sheet.cell_xf_index(r, c)]
                        font = book.font_list[xf.font_index]
                        font_name = getattr(font, "name", None)
                    except Exception:
                        font_name = None

                    if isinstance(v, str):
                        if is_tcvn3_font(font_name) or looks_like_tcvn3(v):
                            v = tcvn3_to_unicode(v)
                    row.append(v)
                rows.append(row)

            if not has_any_data(rows):
                continue

            apply_phi_degrees_decimal(rows)
            out_path = _sheet_output_path(out_dir, file_path, sheet.name, used_names)
            write_csv(out_path, rows)
            written += 1
        except Exception as e:  # noqa: BLE001
            print(
                f"WARN: {file_path.name} / sheet {sheet.name!r} -> {e}",
                file=sys.stderr,
            )
            continue

    return written


def discover_workbooks(in_dir: Path) -> list[Path]:
    """Danh sách .xls/.xlsx (gồm cả tên dạng .xls.xlsx), đã sắp xếp."""
    files = [*in_dir.glob("*.xlsx"), *in_dir.glob("*.xls")]
    return sorted({p.resolve() for p in files if p.is_file()}, key=lambda p: p.name.casefold())


def main(argv: list[str]) -> int:
    script_dir = Path(__file__).resolve().parent
    root = script_dir.parent

    p = argparse.ArgumentParser(
        description="Chuyển toàn bộ .xls/.xlsx trong dataXLSX sang CSV; đổi TCVN3 -> UTF-8 theo font."
    )
    p.add_argument(
        "--input",
        default=str(root / "dataXLSX"),
        help="Thư mục chứa file .xls/.xlsx (mặc định: <root>/dataXLSX)",
    )
    p.add_argument(
        "--output",
        default=str(root / "dataCSV"),
        help="Thư mục xuất CSV (mặc định: <root>/dataCSV)",
    )
    args = p.parse_args(argv)

    in_dir = Path(args.input).resolve()
    out_dir = Path(args.output).resolve()

    if not in_dir.exists() or not in_dir.is_dir():
        print(f"Không tìm thấy thư mục input: {in_dir}", file=sys.stderr)
        return 2

    files = discover_workbooks(in_dir)
    if not files:
        print(f"Không có file .xls/.xlsx trong: {in_dir}")
        return 0

    out_dir.mkdir(parents=True, exist_ok=True)
    total_sheets = 0
    failed: list[tuple[Path, str]] = []
    used_names: set[str] = set()
    for f in files:
        try:
            if f.suffix.lower() == ".xlsx":
                total_sheets += convert_xlsx(f, out_dir, used_names)
            else:
                total_sheets += convert_xls(f, out_dir, used_names)
            print(f"OK: {f.name}")
        except Exception as e:  # noqa: BLE001
            failed.append((f, str(e)))
            print(f"FAIL: {f.name} -> {e}", file=sys.stderr)

    print(f"Đã xuất {total_sheets} file CSV vào: {out_dir}")
    if failed:
        print("Các file lỗi:", file=sys.stderr)
        for f, msg in failed:
            print(f"- {f.name}: {msg}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
