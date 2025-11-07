"""
Focused Donuk (DONDURMALAR) writer.
- Parses CSV and writes only the ice cream section of the Donuk workbook.
- Detects size columns (3,5 KG | 350 GR | 150 GR) and flavor rows.
- Clears numeric values in that block and writes aggregated quantities.

This module intentionally keeps only what's needed now.
"""
from __future__ import annotations

from typing import Dict, Iterable, Optional, Tuple
import os
import re
import unicodedata
import pandas as pd  # pyright: ignore[reportMissingImports]
import openpyxl   # pyright: ignore[reportMissingModuleSource]
from openpyxl import load_workbook, Workbook # pyright: ignore[reportMissingModuleSource]
from openpyxl.styles import Alignment # pyright: ignore[reportMissingModuleSource]
from openpyxl.utils import get_column_letter # pyright: ignore[reportMissingModuleSource]

DATA_START_ROW = 3

# ----------------------------- Normalization -----------------------------

def normalize_text(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s)
    tr = str.maketrans({
        "ı": "i", "ğ": "g", "ş": "s", "ö": "o", "ç": "c", "ü": "u",
        "İ": "I", "Ğ": "G", "Ş": "S", "Ö": "O", "Ç": "C", "Ü": "U",
    })
    s = s.translate(tr)
    s = s.upper()
    s = s.replace("GOGUSLU", "GOGSU").replace("GOGSULU", "GOGSU")
    s = s.replace("HARMANDALI", "EFESUS")
    s = s.replace("AMASRA", "DADAYLI")
    s = re.sub(r"[^\w\s]", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def size_from_stock_or_unit(stock_name: str, unit_text: str) -> Optional[str]:
    # Preserve punctuation for size patterns like 1*3,50 and 350 GR
    def up_keep_punct(s: str) -> str:
        if s is None:
            return ""
        try:
            s = unicodedata.normalize("NFKD", str(s))
        except Exception:
            s = str(s)
        tr = str.maketrans({
            "ı": "i", "ğ": "g", "ş": "s", "ö": "o", "ç": "c", "ü": "u",
            "İ": "I", "Ğ": "G", "Ş": "S", "Ö": "O", "Ç": "C", "Ü": "U",
        })
        s = s.translate(tr).upper()
        s = s.replace(" ", " ")
        return s

    upn_raw = up_keep_punct(stock_name)
    upu_raw = up_keep_punct(unit_text)

    # Helper checks
    def is_35kg(s: str) -> bool:
        # 1*3,5 | 1x3,5 | 1*3.5 | 3,5 KG | 3,50 KG | KL_3,5_KG (do NOT match bare 350)
        if re.search(r"1\s*[\*Xx]\s*3[,\.]?5\b", s):
            return True
        if re.search(r"\b3[,\.]?5\s*KG\b", s):
            return True
        if re.search(r"\b3[,\.]50\s*KG\b", s):  # treat 3,50 KG as 3.5 KG, but not bare 350
            return True
        if re.search(r"KL[\s_\-]*3[,\.]?5", s):
            return True
        return False

    def is_350gr(s: str) -> bool:
        return re.search(r"\b350\s*(GR|G)\b", s) is not None

    def is_150gr(s: str) -> bool:
        return re.search(r"\b150\s*(GR|G)\b", s) is not None

    if is_35kg(upn_raw) or is_35kg(upu_raw):
        return "35KG"
    if is_350gr(upn_raw) or is_350gr(upu_raw):
        return "350GR"
    if is_150gr(upn_raw) or is_150gr(upu_raw):
        return "150GR"
    return None

# --------------------------- CSV Utilities ---------------------------

def read_csv(csv_path: str) -> pd.DataFrame:
    try:
        return pd.read_csv(csv_path, encoding="utf-8", delimiter=",", header=2)
    except Exception:
        return pd.read_csv(csv_path, encoding="utf-8", delimiter=",", header=0)


def find_col(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    cols_up = {normalize_text(c): c for c in df.columns}
    for c in candidates:
        cc = normalize_text(c)
        if cc in cols_up:
            return cols_up[cc]
    for c in candidates:
        cc = normalize_text(c)
        for up, orig in cols_up.items():
            if cc in up:
                return orig
    return None


def read_branch_from_file(csv_path: str) -> Optional[str]:
    try:
        with open(csv_path, encoding="utf-8") as f:
            for line in f:
                up = normalize_text(line)
                if "SUBE" in up and ("KODU" in up or "ADI" in up):
                    part = line.split(":", 1)[-1] if ":" in line else line
                    part = part.strip()
                    if "-" in part:
                        part = part.split("-", 1)[-1]
                    part = part.strip()
                    m = re.search(r"\(([^)]+)\)", part)
                    if m:
                        return m.group(1).strip()
                    if part.upper().endswith(" DEPO"):
                        part = part[:-5].strip()
                    return part
    except Exception:
        pass
    return None

# --------------------------- Excel helpers ---------------------------

def master_cell(ws: openpyxl.worksheet.worksheet.Worksheet, r: int, c: int):
    for mr in ws.merged_cells.ranges:
        min_row, min_col, max_row, max_col = mr.bounds
        if min_row <= r <= max_row and min_col <= c <= max_col:
            return ws.cell(row=min_row, column=min_col)
    return ws.cell(row=r, column=c)


def safe_write(ws: openpyxl.worksheet.worksheet.Worksheet, r: int, c: int, value) -> None:
    ws.cell(row=r, column=c).value = value
    # Check if cell is part of a merged range
    '''for mr in ws.merged_cells.ranges:
        min_row, min_col, max_row, max_col = mr.bounds
        if min_row <= r <= max_row and min_col <= c <= max_col:
            # If this is not the master cell, find the master cell
            if r != min_row or c != min_col:
                master_cell = ws.cell(row=min_row, column=min_col)
                master_cell.value = value
                return
    # If not merged or is master cell, write directly
    cell.value = value'''


def find_branch_span(ws: openpyxl.worksheet.worksheet.Worksheet, branch_name: str) -> Optional[Tuple[int, int, int]]:
    if not branch_name:
        return None
    up = normalize_text(branch_name)
    # Prefer merged ranges whose master cell matches the branch
    for mr in ws.merged_cells.ranges:
        min_row, min_col, max_row, max_col = mr.bounds
        v = ws.cell(row=min_row, column=min_col).value
        if not v:
            continue
        vv = normalize_text(v)
        if vv == up or up in vv or vv in up:
            return (min_col, max_col, min_row)
    # Fallback: scan early rows for any cell matching
    for r in range(1, min(25, ws.max_row) + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if not v:
                continue
            vv = normalize_text(v)
            if vv == up or up in vv or vv in up:
                # If inside a merge, return the full span
                for mr in ws.merged_cells.ranges:
                    min_row, min_col, max_row, max_col = mr.bounds
                    if min_row <= r <= max_row and min_col <= c <= max_col:
                        return (min_col, max_col, min_row)
                return (c, c, r)
    return None

def is_merged_at(ws: openpyxl.worksheet.worksheet.Worksheet, r: int, c: int) -> Optional[Tuple[int, int, int, int]]:
    for mr in ws.merged_cells.ranges:
        min_row, min_col, max_row, max_col = mr.bounds
        if min_row <= r <= max_row and min_col <= c <= max_col:
            return (min_row, min_col, max_row, max_col)
    return None

def resolve_numeric_col(ws: openpyxl.worksheet.worksheet.Worksheet, r: int, c: int, min_c: int, max_c: int) -> int:
    # If target cell is part of a merged region whose master is not this column, shift right until outside merge
    merged = is_merged_at(ws, r, c)
    if merged:
        _, mcol, _, mmax = merged
        # If this column is the master, ok; otherwise move to the right edge of this merge (bounded by branch span)
        if mcol != c:
            c = min(mmax, max_c)
    # Ensure final cell is not part of a label merge; if still merged and master != c, step right
    while True:
        merged = is_merged_at(ws, r, c)
        if not merged:
            break
        _, mcol, _, _ = merged
        if mcol == c:
            break
        if c >= max_c:
            break
        c += 1
    return c


def find_size_columns(ws: openpyxl.worksheet.worksheet.Worksheet, min_c: int, max_c: int, row_hint: int) -> Dict[str, Optional[int]]:
    sizes: Dict[str, Optional[int]] = {"35KG": None, "350GR": None, "150GR": None}
    def scan_rows(r1: int, r2: int):
        # If span width is 1, allow scanning a few columns to the right for unit headers
        c_start = max(1, min_c)
        c_end = min(ws.max_column, max_c if max_c >= min_c else min_c)
        if c_start == c_end and c_end < ws.max_column:
            c_end = min(ws.max_column, c_end + 6)
        for r in range(max(1, r1), min(ws.max_row, r2) + 1):
            for c in range(c_start, c_end + 1):
                v = ws.cell(row=r, column=c).value
                if not v:
                    continue
                s = normalize_text(v)
                def rightmost_if_merged(rr: int, cc: int) -> int:
                    for mr in ws.merged_cells.ranges:
                        min_row, min_col, max_row, max_col = mr.bounds
                        if min_row <= rr <= max_row and min_col <= cc <= max_col:
                            return max_col
                    return cc
                if ("3,5" in s or ("35" in s and "KG" in s)) and sizes["35KG"] is None:
                    sizes["35KG"] = rightmost_if_merged(r, c)
                if ("350" in s and ("GR" in s or "G" in s)) and sizes["350GR"] is None:
                    sizes["350GR"] = rightmost_if_merged(r, c)
                if ("150" in s and ("GR" in s or "G" in s)) and sizes["150GR"] is None:
                    sizes["150GR"] = rightmost_if_merged(r, c)
    if row_hint:
        scan_rows(row_hint, row_hint + 3)
    if not all(sizes.values()):
        scan_rows(1, 25)
    return sizes


def locate_dondurmalar_block(ws: openpyxl.worksheet.worksheet.Worksheet, min_c: int, max_c: int, debug: bool = False) -> Tuple[int, Dict[str, int]]:
    """Find the 'DONDURMALAR' header row and the size columns on that row within (min_c..max_c, plus small right margin).
    Returns: (header_row_index, size_columns dict with keys '35KG','350GR','150GR', and 'MONO','KUCUK','BUYUK').
    """
    header_row = None
    pasta_cols = {"MONO": None, "KUCUK": None, "BUYUK": None}
    all_pasta_rows = []  # Pasta başlıkları için tüm satırları tut

    for r in range(1, min(ws.max_row, 100) + 1):
        v = ws.cell(row=r, column=1).value
        if not v:
            continue
        up = normalize_text(v)
        if debug:
            print(f"[DEBUG] Checking row {r} for DONDURMALAR: '{v}' -> '{up}'")
        if "DONDURMALAR" in up:
            if debug:
                print(f"[DEBUG] Found DONDURMALAR header at row {r}")
            header_row = r
            if ("PASTA" in up or "KROKAN" in up or "ORMAN" in up or "GANAJ" in up or "ANANAS" in up or "FISTIK" in up):
                all_pasta_rows.append(r)
                if debug:
                    print(f"[DEBUG] Found potential pasta row at {r}: '{v}'")

        if all_pasta_rows:
            # En üstteki pasta satırı
            first_pasta_row = min(all_pasta_rows)
            # Bu satırdan önceki 2 satırı kontrol et - MONO/KÜÇÜK/BÜYÜK başlıkları burada olmalı
            for r in range(max(1, first_pasta_row - 2), first_pasta_row + 1):
                for c in range(min_c, max_c + 1):
                    v = ws.cell(row=r, column=c).value
                    if not v:
                        continue
                    up = normalize_text(v)
                    if debug:
                        print(f"[DEBUG] Checking potential pasta header cell r={r} c={c}: '{v}' -> '{up}'")
                    # Daha esnek başlık eşleştirme
                    if "MONO" in up or "TEK" in up or "36" in up:
                        pasta_cols["MONO"] = c
                        if debug:
                            print(f"[DEBUG] Found MONO pasta column at r={r} c={c}")
                    elif ("KUCUK" in up or "KÜÇÜK" in up):
                        pasta_cols["KUCUK"] = c
                        if debug:
                            print(f"[DEBUG] Found KUCUK pasta column at r={r} c={c}")
                    elif ("BUYUK" in up or "BÜYÜK" in up):
                        pasta_cols["BUYUK"] = c
                        if debug:
                            print(f"[DEBUG] Found BUYUK pasta column at r={r} c={c}")
            # Pasta kolonları için özel kontrol: header row ve önceki 3 satır
            pasta_cols = {"MONO": None, "KUCUK": None, "BUYUK": None}
            check_rows = list(range(max(1, header_row - 3), header_row + 1))
            if debug:
                print(f"[DEBUG] Checking rows {check_rows} for pasta columns in cols {min_c}-{max_c}")
            for rr in check_rows:
                for c in range(min_c, max_c + 1):
                    v = ws.cell(row=rr, column=c).value
                    if not v:
                        continue
                    up = normalize_text(v)
                    if debug:
                        print(f"[DEBUG] Checking cell r={rr} c={c}: '{v}' -> '{up}'")
                    if ("MONO" in up or "TEK" in up) and "PASTA" in up:
                        pasta_cols["MONO"] = c
                        if debug:
                            print(f"[DEBUG] Found MONO pasta column at r={rr} c={c}")
                    elif ("KUCUK" in up or "KÜÇÜK" in up) and "PASTA" in up:
                        pasta_cols["KUCUK"] = c
                        if debug:
                            print(f"[DEBUG] Found KUCUK pasta column at r={rr} c={c}")
                    elif ("BUYUK" in up or "BÜYÜK" in up) and "PASTA" in up:
                        pasta_cols["BUYUK"] = c
                        if debug:
                            print(f"[DEBUG] Found BUYUK pasta column at r={rr} c={c}")
            break
    
    if header_row is None:
        # Fallback: use given row_hint area
        header_row = max(1, min(10, ws.max_row))
        if debug:
            print(f"[DEBUG] Using fallback header row {header_row}")
    
    sizes = find_size_columns(ws, min_c, max_c, row_hint=header_row)
    
    if debug:
        print(f"[DEBUG] Found size columns: {sizes}")
        print(f"[DEBUG] Found pasta columns: {pasta_cols}")
    
        # Combine size and pasta columns
        all_cols = {k: v for k, v in sizes.items() if v}
        all_cols.update({k: v for k, v in pasta_cols.items() if v})
        return header_row, all_cols

def find_pasta_rows(ws: openpyxl.worksheet.worksheet.Worksheet, base_col: int, start_row: int, debug: bool = False) -> Dict[str, Optional[int]]:
    """Find pasta product rows in worksheet by scanning the branch's first column (base_col).

    Scans downward from start_row+1 to locate rows that contain the pasta type labels
    (KROKAN, FISTIK, ORMAN/MEYVELI, GANAJ, ANANAS) in the given column.
    Returns a dict mapping pasta keys to the found row index (or None).
    """
    targets = {
        "KROKANLI": None,
        "FISTIKLI": None,
        "ORMAN": None,
        "GANAJ": None,
        "ANANAS": None
    }

    # scan a reasonable range below the header (e.g., next 30 rows) or until worksheet end
    max_scan = min(ws.max_row, start_row + 60)
    for r in range(start_row + 1, max_scan + 1):
        v = ws.cell(row=r, column=base_col).value
        if not v:
            continue
        up = normalize_text(v)
        if debug:
            print(f"[DEBUG] Checking row {r} col {base_col} for pasta: '{v}' -> '{up}'")

        if ("KROKANLI" in up or "KROKAN" in up) and targets["KROKANLI"] is None:
            targets["KROKANLI"] = r
            if debug:
                print(f"[DEBUG] Found KROKANLI pasta at row {r}")
        elif ("FISTIKLI" in up or "FISTIK" in up or "ANTEP" in up) and targets["FISTIKLI"] is None:
            targets["FISTIKLI"] = r
            if debug:
                print(f"[DEBUG] Found FISTIKLI pasta at row {r}")
        elif ("ORMAN" in up or ("MEYVELI" in up and "ORMAN" in up)) and targets["ORMAN"] is None:
            targets["ORMAN"] = r
            if debug:
                print(f"[DEBUG] Found ORMAN MEYVELI pasta at row {r}")
        elif ("GANAJ" in up or "GANAJLI" in up) and targets["GANAJ"] is None:
            targets["GANAJ"] = r
            if debug:
                print(f"[DEBUG] Found GANAJ pasta at row {r}")
        elif ("ANANAS" in up or "ANANASLI" in up) and targets["ANANAS"] is None:
            targets["ANANAS"] = r
            if debug:
                print(f"[DEBUG] Found ANANAS pasta at row {r}")

    if debug:
        print(f"[DEBUG] Found pasta rows: {targets}")
        for k, v in targets.items():
            if v is None:
                print(f"[DEBUG] WARNING: Could not find row for {k} pasta")

    return targets

def pasta_key_from_name(name_up: str) -> str:
    tests = [
        ("KROKANLI", ["KROKAN"]),
        ("FISTIKLI", ["FISTIK"]),
        ("ORMAN", ["ORMAN", "MEYVELI"]),
        ("GANAJ", ["GANAJ"]),
        ("ANANAS", ["ANANAS"])
    ]
    for key, needles in tests:
        if any(n in name_up for n in needles):
            return key
    return ""

# ----------------------- DONUK: GENERIC MATRIX BLOCKS -----------------------

def find_group_header_rows(ws: openpyxl.worksheet.worksheet.Worksheet, groups: Iterable[str]) -> Dict[str, int]:
    res: Dict[str, int] = {}
    wanted = {normalize_text(g): g for g in groups}
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if not v:
            continue
        up = normalize_text(v)
        if up in wanted and wanted[up] not in res:
            res[wanted[up]] = r
    return res


def scan_variant_columns(ws: openpyxl.worksheet.worksheet.Worksheet, header_row: int, min_c: int, max_c: int) -> Tuple[Dict[str, int], int]:
    """Scan for variant headers on the given header row or the immediate next row.
    Returns (variants_map, row_used).
    """
    def scan_on_row(row_idx: int) -> Dict[str, int]:
        variants: Dict[str, int] = {}
        c_start = max(1, min_c)
        c_end = min(ws.max_column, max(max_c, min_c) + 12)
        for c in range(c_start, c_end + 1):
            v = ws.cell(row=row_idx, column=c).value
            if not v:
                continue
            up = normalize_text(v)
            if not up or any(k in up for k in ["3,5", "350", "150", "KG", "GR", "DONDURMALAR"]):
                # Skip size-like headers or section labels
                continue
            # Skip pure numbers or tokens without letters (e.g., '2')
            if not re.search(r"[A-Z]", up):
                continue
            if len(up) <= 2:
                continue
            # rightmost if merged
            def rightmost(rr: int, cc: int) -> int:
                for mr in ws.merged_cells.ranges:
                    min_row, min_col, max_row, max_col = mr.bounds
                    if min_row <= rr <= max_row and min_col <= cc <= max_col:
                        return max_col
                return cc
            col = rightmost(row_idx, c)
            variants[up] = col
        return variants

    # try header row first, then header_row+1, then header_row+2
    v_now = scan_on_row(header_row)
    if v_now:
        return v_now, header_row
    if header_row + 1 <= ws.max_row:
        v_next = scan_on_row(header_row + 1)
        if v_next:
            return v_next, header_row + 1
    if header_row + 2 <= ws.max_row:
        v_next2 = scan_on_row(header_row + 2)
        if v_next2:
            return v_next2, header_row + 2
    return {}, header_row


def scan_product_rows(ws: openpyxl.worksheet.worksheet.Worksheet, start_row: int, stop_row: int) -> Dict[int, str]:
    rows: Dict[int, str] = {}
    for r in range(start_row, min(ws.max_row, stop_row)):
        v = ws.cell(row=r, column=1).value
        if v is None or str(v).strip() == "":
            # allow anonymous block row
            continue
        up = normalize_text(v)
        # Skip section headers (another group name)
        if up in {"DONDURMALAR", "TOST", "EKMEK", "CHEESECAKE", "CATAL BOREK", "ÇATAL BÖREK"}:
            break
        rows[r] = up
        # Heuristic: stop collecting after hitting an empty line following content
    if not rows:
        # Provide one anonymous row to write into if there is no explicit product row
        rows[start_row] = ""
    return rows


def build_blocks(ws: openpyxl.worksheet.worksheet.Worksheet, min_c: int, max_c: int) -> list:
    groups = ["TOST", "EKMEK", "CHEESECAKE", "ÇATAL BÖREK", "CATAL BOREK"]
    hdrs = find_group_header_rows(ws, groups)
    # Determine the next header row to bound product rows
    hdr_positions = sorted(hdrs.values()) + [ws.max_row + 1]
    blocks = []
    for gname, r in hdrs.items():
        # find next header below r
        next_r = None
        for p in hdr_positions:
            if p > r:
                next_r = p
                break
        variants, used_row = scan_variant_columns(ws, r, min_c, max_c)
        # start products below the variant header row (used_row); default r+1
        start_products = (used_row + 1) if variants else (r + 1)
        prows = scan_product_rows(ws, start_products, next_r if next_r is not None else ws.max_row + 1)
        if variants:
            blocks.append({
                "group": gname,
                "header_row": r,
                "variants": variants,  # up -> col
                "rows": prows,         # row_index -> up_name (may be empty)
                "write_row": start_products,
            })
    return blocks


def route_group_for_name(name_up: str) -> Optional[str]:
    """Route CSV row name to a target block group by keywords."""
    if any(k in name_up for k in ["CHEESE", "CHEESECAKE", "TRILECE", "TRILECE"]):
        return "CHEESECAKE"
    if any(k in name_up for k in ["BOREK", "BÖREK", "CATAL", "ÇATAL", "KOL BOREGI", "SU BOREGI", "ISPANAKLI", "PATATESLI", "KIYMALI"]):
        return "CATAL BOREK"
    if "EKMEK" in name_up:
        return "EKMEK"
    if "TOST" in name_up:
        return "TOST"
    return None


def match_block_entry(name_up: str, blocks: list, desired_group: Optional[str] = None) -> Optional[Tuple[dict, int, int]]:
    # Try to find best (block, product_row, variant_col)
    best = None
    best_score = 0
    for b in blocks:
        # Restrict to routed desired group if provided
        if desired_group and normalize_text(b.get("group", "")) != normalize_text(desired_group):
            continue
        # match variant
        for v_up, col in b["variants"].items():
            if not v_up:
                continue
            # tokenized variant matching: any word length>=3 in csv name
            tokens = [t for t in re.split(r"\s+", v_up) if len(t) >= 3]
            if any(t in name_up for t in tokens):
                # match product row if any
                candidate_rows = list(b["rows"].items())
                chosen_row = None
                chosen_row_score = 0
                for r_idx, r_name in candidate_rows:
                    if not r_name:
                        # anonymous row baseline
                        if chosen_row is None:
                            chosen_row = r_idx
                            chosen_row_score = 1
                        continue
                    if r_name and r_name in name_up:
                        # prefer exact product mention
                        if 10 > chosen_row_score:
                            chosen_row = r_idx
                            chosen_row_score = 10
                score = 100 + chosen_row_score + len(v_up)
                if score > best_score and chosen_row is not None:
                    best_score = score
                    best = (b, chosen_row, col)
    return best




def find_dondurma_rows(ws: openpyxl.worksheet.worksheet.Worksheet) -> Dict[str, Optional[int]]:
    targets = {"SUTLU": None, "KAKAOLU": None, "ANTEP": None, "KROKAN": None, "KARADUT": None,
               "LIMON": None, "DAMLA": None, "CILEK": None, "LIGHT": None, "BLUE": None,
               "CARK": None, "DOSIDO": None}
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if not v:
            continue
        up = normalize_text(v)
        if "SUTLU" in up and targets["SUTLU"] is None:
            targets["SUTLU"] = r
        elif "KAKAOLU" in up and targets["KAKAOLU"] is None:
            targets["KAKAOLU"] = r
        elif ("ANTEP" in up or "FISTIK" in up) and targets["ANTEP"] is None:
            targets["ANTEP"] = r
        elif "KROKAN" in up and targets["KROKAN"] is None:
            targets["KROKAN"] = r
        elif "KARADUT" in up and targets["KARADUT"] is None:
            targets["KARADUT"] = r
        elif "LIMON" in up and targets["LIMON"] is None:
            targets["LIMON"] = r
        elif ("DAMLA" in up or "SAKIZ" in up) and targets["DAMLA"] is None:
            targets["DAMLA"] = r
        elif "CILEK" in up and targets["CILEK"] is None:
            targets["CILEK"] = r
        elif ("LIGHT" in up) and targets["LIGHT"] is None:
            targets["LIGHT"] = r
        elif ("BLUE" in up or "SKY" in up) and targets["BLUE"] is None:
            targets["BLUE"] = r
        elif ("CARK" in up or "CARKIFELEK" in up) and targets["CARK"] is None:
            targets["CARK"] = r
        elif ("DOSIDO" in up or "DOSİDO" in up) and targets["DOSIDO"] is None:
            targets["DOSIDO"] = r
    return targets


def flavor_key_from_name(name_up: str) -> str:
    # Precedence matters: BLUE SKY overrides SADE; SADE maps to SUTLU (plain) unless LIGHT explicitly present
    if ("BLUE" in name_up or "SKY" in name_up):
        return "BLUE"
    if "LIGHT" in name_up:
        return "LIGHT"
    if "SADE" in name_up:
        return "SUTLU"
    tests = [
        ("SUTLU", ["SUTLU"]),
        ("KAKAOLU", ["KAKAOLU"]),
        ("ANTEP", ["ANTEP", "FISTIK"]),
        ("KROKAN", ["KROKAN"]),
        ("KARADUT", ["KARADUT"]),
        ("LIMON", ["LIMON"]),
        ("DAMLA", ["DAMLA", "SAKIZ"]),
        ("CILEK", ["CILEK"]),
        ("CARK", ["CARK", "CARKIFELEK"]),
        ("DOSIDO", ["DOSIDO", "DOSİDO"]),
    ]
    for key, needles in tests:
        if any(n in name_up for n in needles):
            return key
    return ""

# ----------------------- DONUK: DONDURMALAR ONLY -----------------------

def process_donuk_csv(csv_path: str, output_path: str = "sevkiyat_donuk.xlsx", sheet_name: Optional[str] = None) -> Tuple[int, int]:
    debug = True  # Force debug mode on
    df = read_csv(csv_path)
    stok_col = find_col(df, ["STOK KODU", "STOKKODU", "KOD"])
    miktar_col = find_col(df, ["MIKTAR", "MİKTAR", "ADET"])
    grup_col = find_col(df, ["GRUP", "KATEGORI", "KATEGORI ADI"])
    if not stok_col or not miktar_col:
        raise ValueError("CSV'de 'Stok Kodu' veya 'Miktar' sütunu bulunamadı.")

    branch_guess = read_branch_from_file(csv_path)
    branch_name = branch_guess  # Use branch_guess as branch_name for text updates

    if os.path.exists(output_path):
        wb = load_workbook(output_path)
    else:
        wb = Workbook()

    # Select target worksheet
    ws = None
    if sheet_name and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        if branch_guess:
            for w in wb.worksheets:
                if find_branch_span(w, branch_guess):
                    ws = w
                    break
        if ws is None:
            ws = wb.worksheets[0]

    span = find_branch_span(ws, branch_guess) if branch_guess else None
    if span:
        min_c, max_c, branch_row = span
        # Genişlet arama aralığını
        max_c = max(max_c + 4, min_c + 12)  # En az 12 kolon tara
    else:
        min_c, max_c, branch_row = 2, 14, 2  # Default olarak daha geniş bir aralık
    # Discover the actual DONDURMALAR header row and size columns near this branch span
    header_row, size_cols = locate_dondurmalar_block(ws, min_c, max_c, debug)

    # Determine pasta columns relative to the branch span (user provided mapping):
    # Mono: branch_col + 1, Küçük: branch_col + 2, Büyük: branch_col + 3
    pasta_cols = {
        "MONO": (min_c + 1) if (min_c + 1) <= ws.max_column else None,
        "KUCUK": (min_c + 2) if (min_c + 2) <= ws.max_column else None,
        "BUYUK": (min_c + 3) if (min_c + 3) <= ws.max_column else None,
    }
    if debug:
        print(f"[DEBUG] Pasta columns assigned from branch span: {pasta_cols} (min_c={min_c})")

    # Find pasta rows by scanning the first column of the branch (min_c)
    pasta_rows = find_pasta_rows(ws, min_c, header_row, debug)

    # Ensure aggregator exists
    aggreg: Dict[Tuple[int, int], float] = {}
    matched = 0

    # Aggregate pasta entries from CSV into aggreg by mapping pasta type -> pasta_rows and size -> pasta_cols
    for _, r in df.iterrows():
        name = str(r[stok_col])
        up = normalize_text(name)

        pasta_key = pasta_key_from_name(up)
        if not pasta_key:
            continue

        unit_text = str(r.get("Birim", r.get("BIRIM", "")))

        if debug:
            print(f"[DEBUG] Found pasta CSV row: {name} -> pasta_key={pasta_key}")

        # choose column by size tokens in CSV name/unit
        col_idx = None
        if "MONO" in up or "TEK" in up or re.search(r"\b36\b", up):
            col_idx = pasta_cols.get("MONO")
            if debug:
                print(f"[DEBUG] MONO pasta detected, column={col_idx}")
        elif "KUCUK" in up or "KÜÇÜK" in up:
            col_idx = pasta_cols.get("KUCUK")
            if debug:
                print(f"[DEBUG] KUCUK pasta detected, column={col_idx}")
        elif "BUYUK" in up or "BÜYÜK" in up:
            col_idx = pasta_cols.get("BUYUK")
            if debug:
                print(f"[DEBUG] BUYUK pasta detected, column={col_idx}")

        if debug:
            print(f"[DEBUG] Pasta rows: {pasta_rows}")

        if not col_idx:
            if debug:
                print(f"[DEBUG] WARNING: No column match for size in: {name}")
            continue

        try:
            qty = float(str(r[miktar_col]).replace(",", "."))
        except Exception:
            continue

        row_idx = pasta_rows.get(pasta_key)
        if not row_idx:
            if debug:
                print(f"[DEBUG] WARNING: No pasta row found for key {pasta_key} (source='{name}')")
            continue

        key = (row_idx, col_idx)
        aggreg[key] = aggreg.get(key, 0.0) + qty
        matched += 1

    if debug:
        print(f"[DEBUG] Branch span min_c={min_c} max_c={max_c} row={branch_row}")
        print(f"[DEBUG] DONDURMALAR header at row={header_row}")
        print(f"[DEBUG] Size columns: {size_cols}")

    flavor_rows = find_dondurma_rows(ws)
    if debug:
        print(f"[DEBUG] Flavor rows: {flavor_rows}")

    # 'aggreg' and 'matched' may already be populated by earlier pasta aggregation.
    # Keep block_aggreg and other counters here.
    block_aggreg: Dict[Tuple[int, int], float] = {}
    rokoko_total: float = 0.0

    blocks = build_blocks(ws, min_c, max_c)
    if debug:
        print(f"[DEBUG] Blocks: {[{'group': b['group'], 'header_row': b['header_row'], 'variants': list(b['variants'].keys())} for b in blocks]}")

    # Explicit 4-subcolumn layout under this branch
    # Determine the 4 numeric subcolumns of this branch
    sub_cols = [min_c + i for i in range(4) if (min_c + i) <= ws.max_column]
    # Helper: find row index by label(s) in column A
    def find_row_by_label(keywords: Iterable[str]) -> Optional[int]:
        for rr in range(1, ws.max_row + 1):
            v = ws.cell(row=rr, column=1).value
            if not v:
                continue
            upv = normalize_text(v)
            if all(k in upv for k in keywords):
                return rr
        return None

    row_tost = find_row_by_label(["TOST"]) if sub_cols else None
    row_ekmek = find_row_by_label(["EKMEK"]) if sub_cols else None
    row_cheese = find_row_by_label(["CHEESECAKE"]) if sub_cols else None
    row_catal = find_row_by_label(["CATAL", "BOREK"]) if sub_cols else None
    
    # Aggregators for explicit layout
    matrix_aggreg: Dict[Tuple[int, int], float] = {}
    kunefe_total_sum: float = 0.0
    trilece_total_sum: float = 0.0
    ekler_total_sum: float = 0.0
    for _, r in df.iterrows():
        name = str(r[stok_col])
        grp = str(r[grup_col]) if grup_col and grup_col in df.columns else ""
        up = normalize_text(name)
        g_up = normalize_text(grp)
        

        if "EKLER" in up:
            try:
                qty_r = float(str(r[miktar_col]).replace(",", "."))
            except Exception:
                qty_r = 0.0
            ekler_total_sum += qty_r
            matched += 1
            if debug:
                print(f"[DEBUG] EKLER +{qty_r} now {ekler_total_sum}")
            continue
        # Collect MEYVELI ROKOKO totals for a dedicated text cell update later
        if ("ROKOKO" in up and "MEYVELI" in up) or ("MEYVELI" in up and "ROKOKO" in up):
            try:
                qty_r = float(str(r[miktar_col]).replace(",", "."))
            except Exception:
                qty_r = 0.0
            rokoko_total += qty_r
            if debug:
                print(f"[DEBUG] ROKOKO +{qty_r} now {rokoko_total}")
            continue  # ROKOKO is handled as a text cell, not in the DONDURMALAR grid
        # New explicit mapping path for TATLI/BOREK items
        if "KUNEFE" in up:
            try:
                qty_r = float(str(r[miktar_col]).replace(",", "."))
            except Exception:
                qty_r = 0.0
            kunefe_total_sum += qty_r
            matched += 1
            if debug:
                print(f"[DEBUG] KUNEFE +{qty_r} now {kunefe_total_sum}")
            continue
        if "TRILECE" in up:
            try:
                qty_r = float(str(r[miktar_col]).replace(",", "."))
            except Exception:
                qty_r = 0.0
            trilece_total_sum += qty_r
            matched += 1
            if debug:
                print(f"[DEBUG] TRILECE +{qty_r} now {trilece_total_sum}")
            continue
        
        if g_up in ("TATLI", "BOREK") and sub_cols:
            # parse quantity
            try:
                qty = float(str(r[miktar_col]).replace(",", "."))
            except Exception:
                qty = None
            if qty is not None and qty != 0:
                def add(rr: Optional[int], cols: Iterable[int], val: float):
                    if not rr:
                        return
                    for c_ in cols:
                        key = (rr, c_)
                        matrix_aggreg[key] = matrix_aggreg.get(key, 0.0) + val
                        if debug:
                            print(f"[DEBUG] MATRIX + r={rr} c={c_} val={val} src='{name}'")

                # Specials (apply before groups)
                

                # TOST
                if row_tost and "TOST" in up:
                    if "KASAR" in up:
                        add(row_tost, [sub_cols[0]], qty)
                        matched += 1
                        continue
                    if "KEPEK" in up and len(sub_cols) >= 2:
                        add(row_tost, [sub_cols[1]], qty)
                        matched += 1
                        continue
                    if "KARISIK" in up:
                        cols = sub_cols[2:4] if len(sub_cols) >= 4 else sub_cols[2:3]
                        if cols:
                            add(row_tost, cols, qty)
                            matched += 1
                            continue

                # EKMEK
                if row_ekmek:
                    if ("EKMEK" in up) and ("BEYAZ" in up):
                        add(row_ekmek, [sub_cols[0]], qty)
                        matched += 1
                        continue
                    if ("EKMEK" in up) and ("ESMER" in up) and len(sub_cols) >= 2:
                        add(row_ekmek, [sub_cols[1]], qty)
                        matched += 1
                        continue
                    if ("KIYMALI" in up or "KIYMA" in up) and ("BORE" in up):
                        cols = sub_cols[2:4] if len(sub_cols) >= 4 else sub_cols[2:3]
                        if cols:
                            add(row_ekmek, cols, qty)
                            matched += 1
                            continue

                # CHEESECAKE
                if row_cheese:
                    if "SEBASTIAN" in up:
                        add(row_cheese, sub_cols[:2], qty)
                        matched += 1
                        continue
                    if "FRAMBUAZ" in up:
                        cols = sub_cols[2:4] if len(sub_cols) >= 3 else []
                        if cols:
                            add(row_cheese, cols, qty)
                            matched += 1
                            continue

                # ÇATAL BÖREK
                if row_catal:
                    if "PATATES" in up:
                        add(row_catal, [sub_cols[0]], qty)
                        matched += 1
                        continue
                    if "ISPANAK" in up and len(sub_cols) >= 2:
                        add(row_catal, [sub_cols[1]], qty)
                        matched += 1
                        continue
                    if ("SU" in up) and ("BORE" in up):
                        cols = sub_cols[2:4] if len(sub_cols) >= 3 else []
                        if cols:
                            add(row_catal, cols, qty)
                            matched += 1
                            continue
        # Else try DONDURMALAR grid
        fkey = flavor_key_from_name(up)
        if not fkey:
            continue
        row_idx = flavor_rows.get(fkey)
        if not row_idx:
            continue
        unit_text = str(r.get("Birim", r.get("BIRIM", "")))
        sz = size_from_stock_or_unit(name, unit_text)
        if not sz:
            # Special-case DOSIDO: no explicit size, write into 3,5 KG by convention
            if fkey == "DOSIDO":
                sz = "35KG"
            else:
                continue
        col_idx = size_cols.get("35KG") if sz == "35KG" else size_cols.get("350GR") if sz == "350GR" else size_cols.get("150GR")
        # If size columns couldn't be detected for this branch, skip to avoid writing into wrong areas
        if not col_idx:
            continue
        try:
            qty = float(str(r[miktar_col]).replace(",", "."))
        except Exception:
            continue
        key = (row_idx, col_idx)
        aggreg[key] = aggreg.get(key, 0.0) + qty
        matched += 1
        if debug:
            print(f"[DEBUG] + {name} -> fkey={fkey} size={sz} row={row_idx} col={col_idx} qty={qty}")

    rows_to_clear = [r for r in flavor_rows.values() if r and r > header_row]
    cols_to_clear_raw = [c for c in [size_cols.get("35KG"), size_cols.get("350GR"), size_cols.get("150GR")] if c]
    for rr in rows_to_clear:
        for cc_raw in cols_to_clear_raw:
            cc = resolve_numeric_col(ws, rr, cc_raw, min_c, max_c)
            cell = ws.cell(row=rr, column=cc)
            val = cell.value
            if isinstance(val, (int, float)):
                cell.value = None
            else:
                try:
                    fv = float(str(val).replace(",", ".")) if (val not in (None, "") and str(val).strip() != "") else None
                    if fv is not None:
                        cell.value = None
                except Exception:
                    pass

    # Clear numeric cells in Dondurmalar block
    for (r_, c_), v in aggreg.items():
        cc = resolve_numeric_col(ws, r_, c_, min_c, max_c)
        try:
            safe_write(ws, r_, cc, v)
            if debug:
                print(f"[DEBUG] WRITE r={r_} c={cc} val={v}")
        except Exception as e:
            if debug:
                print(f"[DEBUG] WRITE ERROR r={r_} c={cc} val={v} error={e}")
            # Try to write to master cell if merged
            try:
                master = master_cell(ws, r_, cc)
                master.value = v
                if debug:
                    print(f"[DEBUG] MASTER WRITE r={r_} c={cc} val={v}")
            except Exception as e2:
                if debug:
                    print(f"[DEBUG] MASTER WRITE ERROR r={r_} c={cc} val={v} error={e2}")

    # Explicit matrix writes: clear touched cells, write aggregated, mirror ŞERBET per column
    if matrix_aggreg:
        touched: Dict[int, set] = {}
        for (rr, cc) in matrix_aggreg.keys():
            touched.setdefault(rr, set()).add(cc)
        # clear
        for rr, cols in touched.items():
            for cc in sorted(cols):
                try:
                    cell = ws.cell(row=rr, column=cc)
                    val = cell.value
                    if isinstance(val, (int, float)) or (isinstance(val, str) and val.strip() and re.match(r"^[0-9,\.]+$", val.strip())):
                        cell.value = None
                except Exception:
                    pass
        # write
        for (rr, cc), v in matrix_aggreg.items():
            try:
                safe_write(ws, rr, cc, v)
                if debug:
                    print(f"[DEBUG] MATRIX WRITE r={rr} c={cc} val={v}")
            except Exception as e:
                if debug:
                    print(f"[DEBUG] MATRIX WRITE ERROR r={rr} c={cc} val={v} error={e}")
                # Try to write to master cell if merged
                try:
                    master = master_cell(ws, rr, cc)
                    master.value = v
                    if debug:
                        print(f"[DEBUG] MATRIX MASTER WRITE r={rr} c={cc} val={v}")
                except Exception as e2:
                    if debug:
                        print(f"[DEBUG] MATRIX MASTER WRITE ERROR r={rr} c={cc} val={v} error={e2}")
        # no direct numeric mirroring for ŞERBET; text update below

    # Update KÜNEFE/ŞERBET/DONUK KAR. TRİLEÇE using MEYVELI ROKOKO text update logic
    # These should update the text in the cell, not write to separate numeric cells

    if kunefe_total_sum and ws is not None:
        fmt_qty = int(kunefe_total_sum) if float(kunefe_total_sum).is_integer() else kunefe_total_sum
        
        # Find branch columns first
        branch_cols = []
        for r in range(1, min(4, ws.max_row + 1)):
            for c in range(1, ws.max_column + 1):
                val = ws.cell(row=r, column=c).value
                if not isinstance(val, str):
                    continue
                upv = normalize_text(val)
                if branch_name in upv:
                    branch_cols.append(c)
                    if debug:
                        print(f"[DEBUG] KUNEFE - Found branch '{branch_name}' at r={r} c={c}")
        
        # Search for KÜNEFE only in branch columns
        if branch_cols:
            for r in range(1, ws.max_row + 1):
                for c in branch_cols:
                    val = ws.cell(row=r, column=c).value
                    if not isinstance(val, str):
                        continue
                    upv = normalize_text(val)
                    if "KUNEFE" in upv or "KÜNEFE" in upv:
                        # Preserve left side, replace or append quantity after '=' if present
                        text = val
                        if "=" in text:
                            left, _sep, _right = text.partition("=")
                            new_text = f"{left.strip()} = {fmt_qty}"
                        else:
                            new_text = f"{text.strip()} {fmt_qty}"
                        ws.cell(row=r, column=c).value = new_text
                        if debug:
                            print(f"[DEBUG] KUNEFE TEXT WRITE r={r} c={c} val='{new_text}' (branch column)")
                        # Update only the first match
                        break
                else:
                    continue
                break
        else:
            if debug:
                print(f"[DEBUG] KUNEFE - No branch columns found for '{branch_name}'")

    if trilece_total_sum and ws is not None:
        fmt_qty = int(trilece_total_sum) if float(trilece_total_sum).is_integer() else trilece_total_sum
        
        # Find branch columns first
        branch_cols = []
        for r in range(1, min(4, ws.max_row + 1)):
            for c in range(1, ws.max_column + 1):
                val = ws.cell(row=r, column=c).value
                if not isinstance(val, str):
                    continue
                upv = normalize_text(val)
                if branch_name in upv:
                    branch_cols.append(c)
                    if debug:
                        print(f"[DEBUG] TRİLEÇE - Found branch '{branch_name}' at r={r} c={c}")
        
        # Search for DONUK KAR. TRİLEÇE only in branch columns
        if branch_cols:
            for r in range(1, ws.max_row + 1):
                for c in branch_cols:
                    val = ws.cell(row=r, column=c).value
                    if not isinstance(val, str):
                        continue
                    upv = normalize_text(val)
                    if "DONUK" in upv and ("TRILECE" in upv or "TRİLEÇE" in upv):
                        # Preserve left side, replace or append quantity after '=' if present
                        text = val
                        if "=" in text:
                            left, _sep, _right = text.partition("=")
                            new_text = f"{left.strip()} = {fmt_qty}"
                        else:
                            new_text = f"{text.strip()} {fmt_qty}"
                        ws.cell(row=r, column=c).value = new_text
                        if debug:
                            print(f"[DEBUG] TRİLEÇE TEXT WRITE r={r} c={c} val='{new_text}' (branch column)")
                        # Update only the first match
                        break
                else:
                    continue
                break
        else:
            if debug:
                print(f"[DEBUG] TRİLEÇE - No branch columns found for '{branch_name}'")

    # ŞERBET should mirror KÜNEFE quantity using text update logic
    if kunefe_total_sum and ws is not None:
        fmt_qty = int(kunefe_total_sum) if float(kunefe_total_sum).is_integer() else kunefe_total_sum
        
        # Find branch columns first
        branch_cols = []
        for r in range(1, min(4, ws.max_row + 1)):
            for c in range(1, ws.max_column + 1):
                val = ws.cell(row=r, column=c).value
                if not isinstance(val, str):
                    continue
                upv = normalize_text(val)
                if branch_name in upv:
                    branch_cols.append(c)
                    if debug:
                        print(f"[DEBUG] ŞERBET - Found branch '{branch_name}' at r={r} c={c}")
        
        # Search for ŞERBET only in branch columns
        if branch_cols:
            for r in range(1, ws.max_row + 1):
                for c in branch_cols:
                    val = ws.cell(row=r, column=c).value
                    if not isinstance(val, str):
                        continue
                    upv = normalize_text(val)
                    if "SERBET" in upv or "ŞERBET" in upv:
                        # Preserve left side, replace or append quantity after '=' if present
                        text = val
                        if "=" in text:
                            left, _sep, _right = text.partition("=")
                            new_text = f"{left.strip()} = {fmt_qty}"
                        else:
                            new_text = f"{text.strip()} {fmt_qty}"
                        ws.cell(row=r, column=c).value = new_text
                        if debug:
                            print(f"[DEBUG] ŞERBET TEXT WRITE r={r} c={c} val='{new_text}' (branch column)")
                        # Update only the first match
                        break
                else:
                    continue
                break
        else:
            if debug:
                print(f"[DEBUG] ŞERBET - No branch columns found for '{branch_name}'")    

    # Update MEYVELI ROKOKO text cell if present and total > 0
    if rokoko_total and ws is not None:
        fmt_qty = int(rokoko_total) if float(rokoko_total).is_integer() else rokoko_total


        branch_cols = []
        for r in range(1, min(4, ws.max_row + 1)):
            for c in range(1, ws.max_column + 1):
                val = ws.cell(row=r, column=c).value
                if not isinstance(val, str):
                    continue
                upv = normalize_text(val)
                if branch_name in upv:
                    branch_cols.append(c)
                    if debug:
                        print(f"[DEBUG] MEYVELI ROKOKO - Found branch '{branch_name}' at r={r} c={c}")
        
        # Search for ŞERBET only in branch columns
        if branch_cols:
            for r in range(1, ws.max_row + 1):
                for c in branch_cols:
                    val = ws.cell(row=r, column=c).value
                    if not isinstance(val, str):
                        continue
                    upv = normalize_text(val)
                    if "ROKOKO" in upv:
                        # Preserve left side, replace or append quantity after '=' if present
                        text = val
                        if "=" in text:
                            left, _sep, _right = text.partition("=")
                            new_text = f"{left.strip()} = {fmt_qty}"
                        else:
                            new_text = f"{text.strip()} {fmt_qty}"
                        ws.cell(row=r, column=c).value = new_text
                        if debug:
                            print(f"[DEBUG] MEYVELI ROKOKO TEXT WRITE r={r} c={c} val='{new_text}' (branch column)")
                        # Update only the first match
                        break
                else:
                    continue
                break
        else:
            if debug:
                print(f"[DEBUG] MEYVELI ROKOKO - No branch columns found for '{branch_name}'")

    if ekler_total_sum and ws is not None:
        fmt_qty = int(ekler_total_sum) if float(ekler_total_sum).is_integer() else ekler_total_sum
        
        # Find branch columns first
        branch_cols = []
        for r in range(1, min(4, ws.max_row + 1)):
            for c in range(1, ws.max_column + 1):
                val = ws.cell(row=r, column=c).value
                if not isinstance(val, str):
                    continue
                upv = normalize_text(val)
                if branch_name in upv:
                    branch_cols.append(c)
                    if debug:
                        print(f"[DEBUG] EKLER - Found branch '{branch_name}' at r={r} c={c}")
        
        # Search for EKLER only in branch columns
        if branch_cols:
            for r in range(1, ws.max_row + 1):
                for c in branch_cols:
                    val = ws.cell(row=r, column=c).value
                    if not isinstance(val, str):
                        continue
                    upv = normalize_text(val)
                    if "EKLER" in upv:
                        # Preserve left side, replace or append quantity after '=' if present
                        text = val
                        if "=" in text:
                            left, _sep, _right = text.partition("=")
                            new_text = f"{left.strip()} = {fmt_qty}"
                        else:
                            new_text = f"{text.strip()} {fmt_qty}"
                        ws.cell(row=r, column=c).value = new_text
                        if debug:
                            print(f"[DEBUG] EKLER TEXT WRITE r={r} c={c} val='{new_text}' (branch column)")
                        # Update only the first match
                        break
                else:
                    continue
                break
        else:
            if debug:
                print(f"[DEBUG] EKLER - No branch columns found for '{branch_name}'")

    wb.save(output_path)
    print(f"[DONUK] Dondurmalar yazıldı: {len(aggreg)} hücre, {matched} kalem -> {output_path}")
    return (matched, 0)

# Minimal placeholder for compatibility

def process_csv(csv_path: str, output_path: str = "sevkiyat_tatlı.xlsx") -> Tuple[int, int]:
    # Restore legacy Tatlı writer: reads CSV 'TATLI' items and writes to sevkiyat_tatlı.xlsx template
    # using the branch columns (TEPSI/ADET) and product rows discovered at runtime.
    # Helpers specific to Tatlı flow
    def normalize_text_strict(s):
        return normalize_text(s).replace(" ", "")

    def normalize_variant(v, product_name=""):
        if not v:
            return ""
        v = str(v).upper()
        v = v.replace("*", "X")
        v = v.replace("İ", "I").replace("Ş", "S").replace("Ğ", "G").replace("Ü", "U").replace("Ö", "O").replace("Ç", "C")
        v = v.replace("BÜYÜK", "BUYUK").replace("EKONOMIK PAKET", "BUYUK").replace("EKONOMIKPAKET", "BUYUK")
        v = v.replace("TEKLIPAKET", "PAKET").replace("TEKLI PAKET", "PAKET")
        if product_name in ("EKMEK KADAYIFI", "SEKERPARE"):
            return "ADET"
        if "TEPSI" in v or "TEPSİ" in v or re.search(r"\b42\b", v) or re.search(r"1X?42", v):
            return "TEPSI"
        if "BUYUK" in v:
            return "BUYUK"
        if "KASE" in v:
            return "KASE"
        if "TEKLI" in v:
            return "TEKLI"
        if "PAKET" in v:
            return "PAKET"
        if re.search(r"\d+X\d+", v) or "ADET" in v or re.search(r"PK", v) or "GR" in v:
            return "ADET"
        return re.sub(r"[^\w\s]", "", v).strip()

    def split_tatli_and_variant(s):
        s = str(s or "")
        s = unicodedata.normalize("NFKD", s).upper()
        s = s.replace("İ", "I").replace("Ş", "S").replace("Ğ", "G").replace("Ü", "U").replace("Ö", "O").replace("Ç", "C")
        s = re.sub(r"\{.*?\}", "", s)
        s = re.sub(r"[^\w\s\*\(\)]", "", s)
        s = re.sub(r"\s+", " ", s)
        s = s.strip()
        m = re.match(r"^(.*?)\s*\(([^)]+)\)", s)
        if m:
            ana_ad = m.group(1).strip()
            varyant = m.group(2).strip()
        else:
            ana_ad = s
            varyant = ""
        varyant_kelimeler = ["TEKLI PAKET", "TEKLI", "PAKET", "KASE", "BUYUK", "ADET", "TEPSI"]
        for kelime in varyant_kelimeler:
            if ana_ad.endswith(" " + kelime):
                ana_ad = ana_ad[:-(len(kelime)+1)].strip()
                varyant = kelime
            elif ana_ad.endswith(kelime):
                ana_ad = ana_ad[:-(len(kelime))].strip()
                varyant = kelime
        if re.search(r"\bTEPS[Iİ]\b", ana_ad) or re.search(r"42 ?L[Iİ]?", ana_ad) or re.search(r"1\*?42", ana_ad):
            ana_ad = re.sub(r"\bTEPS[Iİ]\b", "", ana_ad)
            ana_ad = re.sub(r"42 ?L[Iİ]?", "", ana_ad)
            ana_ad = re.sub(r"1\*?42", "", ana_ad)
            ana_ad = ana_ad.strip()
            varyant = "TEPSI"
        return ana_ad.strip(), varyant.strip()

    def tatli_eslesir(excel_ad, csv_ad):
        if excel_ad == csv_ad:
            return True
        ex_strict = normalize_text_strict(excel_ad)
        csv_strict = normalize_text_strict(csv_ad)
        if "TAVUKGOGSUKAZ" in ex_strict and "TAVUKGOGSUKAZANDIBI" in csv_strict:
            return True
        if (ex_strict in csv_strict or csv_strict in ex_strict) and ("KAZ" not in ex_strict and "KAZANDIBI" not in ex_strict):
            return True
        return False

    def varyant_eslesir(excel_v, csv_v):
        if excel_v == csv_v:
            return True
        if excel_v == "TEPSI":
            return csv_v == "TEPSI"
        if excel_v == "ADET":
            return csv_v in ("ADET", "")
        if not excel_v:
            return csv_v in ("", "ADET")
        return excel_v == csv_v

    # CSV oku (eski davranışla uyumlu header toleransı)
    try:
        df = pd.read_csv(csv_path, encoding="utf-8", delimiter=",", header=2)
    except Exception:
        df = pd.read_csv(csv_path, encoding="utf-8", delimiter=",", header=0)

    stok_col = find_col(df, ["STOK KODU", "STOKKODU", "KOD"]) or "STOK KODU"
    miktar_col = find_col(df, ["MIKTAR", "MİKTAR", "ADET"]) or "MIKTAR"
    grup_col = find_col(df, ["GRUP", "KATEGORI", "KATEGORI ADI"]) or "GRUP"

    # Şube adı (CSV'den)
    sube = read_branch_from_file(csv_path)
    if not sube:
        raise ValueError("CSV'den şube adı (ŞUBE KODU/ADI) tespit edilemedi.")
    sube_norm = normalize_text(sube)

    # Tatlı şablonu/yazım dosyası
    if not os.path.exists(output_path):
        raise FileNotFoundError(f"Tatlı şablonu bulunamadı: {output_path}")
    wb = load_workbook(output_path)

    # Doğru sheet'i ve şube sütunlarını bul (2. satırda şube başlıkları)
    ws = None
    col_tepsi = None
    col_adet = None
    for w in wb.worksheets:
        subeler = {}
        # row=2, columns after first
        for cell in w[2][1:]:
            if cell.value:
                sname = normalize_text(cell.value)
                subeler[sname] = {
                    "tepsi": cell.column,
                    "tepsi_2": cell.column + 1,
                    "adet": cell.column + 2,
                    "adet_2": cell.column + 3,
                }
        for sname, cols in subeler.items():
            if sname == sube_norm:
                ws = w
                col_tepsi = cols["tepsi"]
                col_adet = cols["adet"]
                break
        if ws is not None:
            break

    if ws is None or not col_tepsi or not col_adet:
        raise ValueError(f"Şube '{sube}' için hedef sayfa/sütunlar bulunamadı.")

    # Tarih yaz (B1 benzeri: row=2, col=1)
    from datetime import datetime
    ws.cell(row=2, column=1).value = datetime.today().strftime('%d.%m.%Y')

    # Tüm merge'leri geçici olarak aç
    original_merged = [str(r) for r in ws.merged_cells.ranges]
    for r in original_merged:
        try:
            ws.unmerge_cells(r)
        except Exception:
            pass

    # Excel tarafındaki tatlı ürünleri ve hedef hücreleri indexle
    tatli_cells: Dict[Tuple[str, str], Tuple[int, int]] = {}
    skip_keywords = ["SIPARIS TARIHI", "SIPARIS ALAN", "TESLIM TARIHI", "TEYID EDEN"]
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=1):
        ana_cell = row[0]
        if not ana_cell.value:
            continue
        ana_ad, varyant = split_tatli_and_variant(ana_cell.value)
        ana_ad_norm = normalize_text(ana_ad)
        if not ana_ad_norm:
            continue
        if any(ana_ad_norm.startswith(k) or ana_ad_norm == k for k in skip_keywords):
            continue
        varyant_norm = normalize_variant(varyant)
        if "MIKTAR" in ana_ad_norm or "ADET" in ana_ad_norm or "TEPSI" in ana_ad_norm:
            continue
        if ana_cell.row <= 7:
            tatli_cells[(ana_ad_norm, "TEPSI")] = (ana_cell.row, col_tepsi)
            tatli_cells[(ana_ad_norm, "ADET")] = (ana_cell.row, col_adet)
        else:
            if varyant_norm == "TEPSI":
                tatli_cells[(ana_ad_norm, "TEPSI")] = (ana_cell.row, col_tepsi)
            else:
                tatli_cells[(ana_ad_norm, varyant_norm or "ADET")] = (ana_cell.row, col_tepsi)

    # CSV'den TATLI grubu ürünleri topla: {ana_ad_norm: [(varyant_norm, miktar), ...]}
    csv_index: Dict[str, list] = {}
    for _, r in df.iterrows():
        try:
            g = normalize_text(str(r[grup_col]))
        except Exception:
            g = ""
        if g != "TATLI":
            continue
        stok_name = r[stok_col]
        ana_ad, varyant = split_tatli_and_variant(stok_name)
        ana_ad_norm = normalize_text(ana_ad)
        varyant_norm = normalize_variant(varyant)
        try:
            mikt = float(str(r[miktar_col]).replace(",", "."))
        except Exception:
            mikt = r[miktar_col]
        csv_index.setdefault(ana_ad_norm, []).append((varyant_norm, mikt))

    # Önce hedef hücreleri temizle (eski alışkanlıkla: "-")
    for (a, v), (rr, cc) in tatli_cells.items():
        try:
            ws.cell(row=rr, column=cc).value = "-"
        except Exception:
            # nadir merge-hata güvenliği
            try:
                ws.unmerge_cells(start_row=rr, start_column=cc, end_row=rr, end_column=cc)
                ws.cell(row=rr, column=cc).value = "-"
            except Exception:
                pass

    matched = 0
    # Yazma: önce doğrudan ad, sonra gevşek eşleşme
    for (excel_ad, excel_var), (rr, cc) in tatli_cells.items():
        yazildi = False
        if excel_ad in csv_index:
            for csv_var, csv_miktar in csv_index[excel_ad]:
                if varyant_eslesir(excel_var, csv_var):
                    ws.cell(row=rr, column=cc).value = csv_miktar
                    matched += 1
                    yazildi = True
                    break
                if excel_var == "ADET" and csv_var == "TEPSI" and excel_ad in ("EKMEK KADAYIFI", "SEKERPARE"):
                    ws.cell(row=rr, column=cc).value = csv_miktar
                    matched += 1
                    yazildi = True
                    break
        if not yazildi:
            for csv_name, entries in csv_index.items():
                if tatli_eslesir(excel_ad, csv_name):
                    for csv_var, csv_miktar in entries:
                        if varyant_eslesir(excel_var, csv_var):
                            ws.cell(row=rr, column=cc).value = csv_miktar
                            matched += 1
                            yazildi = True
                            break
                        if excel_var == "ADET" and csv_var == "TEPSI" and excel_ad in ("EKMEK KADAYIFI", "SEKERPARE"):
                            ws.cell(row=rr, column=cc).value = csv_miktar
                            matched += 1
                            yazildi = True
                            break
                        if excel_ad == "KAYMAK" and tatli_eslesir(excel_ad, csv_name):
                            ws.cell(row=rr, column=cc).value = csv_miktar
                            matched += 1
                            yazildi = True
                            break
                    if yazildi:
                        break

    # Merge'leri eski haline getir
    for r in original_merged:
        try:
            ws.merge_cells(r)
        except Exception:
            pass

    wb.save(output_path)
    return matched, 0


def process_all(csv_path: str,
                tatli_output: str = "sevkiyat_tatlı.xlsx",
                donuk_output: str = "sevkiyat_donuk.xlsx",
                lojistik_output: str = "sevkiyat_lojistik.xlsx"):
    m2, u2 = process_donuk_csv(csv_path, output_path=donuk_output)
    return {"donuk": {"matched": m2, "unmatched": u2, "file": donuk_output}}

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Kullanım: python parse_gptfix.py <csv_yolu>")
        raise SystemExit(1)
    summary = process_all(sys.argv[1])
    print(summary)
