import os
import sys
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd  # pyright: ignore[reportMissingImports]
import openpyxl  # pyright: ignore[reportMissingModuleSource]
from openpyxl import Workbook  # pyright: ignore[reportMissingModuleSource]
from datetime import datetime

# ------------------ Normalization ------------------
class TextNormalizer:
    @staticmethod
    def up(s: Optional[str]) -> str:
        if s is None:
            return ""
        s = str(s)
        # First apply Turkish character mapping BEFORE normalization
        tr_map = str.maketrans({
            "ı": "i", "ğ": "g", "ü": "u", "ş": "s", "ö": "o", "ç": "c",
            "İ": "I", "Ğ": "G", "Ü": "U", "Ş": "S", "Ö": "O", "Ç": "C",
        })
        s = s.translate(tr_map)
        s = s.upper()
        # Then normalize unicode
        try:
            import unicodedata
            s = unicodedata.normalize("NFKD", s)
        except Exception:
            pass
        s = s.strip()
        return s

# Constants
DATA_START_ROW = 3

# İzmir Bayi Listesi (kullanıcıdan)
IZMIR_BRANCHES = [
    "Mavibahçe", "Forum", "Point", "Folkart", "Efesus", "Gaziemir", "Balçova", "Hatay", "Folkart Vega", "İstasyon"
]
IZMIR_BRANCH_HINTS = [TextNormalizer.up(b) for b in IZMIR_BRANCHES]
KUŞADASI_HINTS = ["KUSADASI", "KUŞADASI", "AYDIN"]

# CSV'den gelen özel şube isimleri → Excel'deki standart isimler
# Bu mapping sayesinde HARMANDALI→EFESUS, FORUMAVM→FORUM gibi dönüşümler yapılır
BRANCH_NAME_MAPPING = {
    "HARMANDALI": "EFESUS",      # CSV: IZMIR(HARMANDALI) → Excel: EFESUS
    "FORUMAVM": "FORUM",          # CSV: IZMIR(FORUMAVM) → Excel: FORUM
    "FOLKARTVEGA": "FOLKART VEGA", # CSV: IZMIR(FOLKARTVEGA) → Excel: FOLKART VEGA
}

# Birden fazla sevkiyat günü olan şubeler ve hangi Excel sayfalarında bulundukları
# Format: {branch_normalized: [list of possible sheet names]}
MULTI_DAY_BRANCHES = {
    "MAVIBAHCE": ["SALI KARŞIYAKA", "KSK CUMARTESİ"],
    "FORUM": ["SALI KARŞIYAKA", "GÜZELBAHÇE", "KSK CUMARTESİ"],
    "FOLKART": ["SALI KARŞIYAKA", "CUMA İZMİR"],
    "EFESUS": ["SALI KARŞIYAKA", "CUMA İZMİR"],
    "ISTASYON": ["SALI İZMİR", "KSK CUMARTESİ"],
    "GAZIEMIR": ["SALI İZMİR", "CUMA İZMİR"],
    "BALCOVA": ["SALI İZMİR", "CUMA İZMİR"],
    "HATAY": ["SALI İZMİR", "CUMA İZMİR"],
    "FOLKART VEGA": ["SALI İZMİR", "CUMA İZMİR"],
    "KUSADASI": ["KUŞADASI-AYDIN", "KUŞADASI CMERT"],
}

# Sheet name mapping for user-friendly display
SHEET_NAME_MAPPING = {
    "Salı Karşıyaka": "SALI KARŞIYAKA",
    "Salı İzmir": "SALI İZMİR",
    "Cuma İzmir": "CUMA İZMİR",
    "Cumartesi KSK": "KSK CUMARTESİ",
    "Güzelbahçe": "GÜZELBAHÇE",
    "Kuşadası-Aydın": "KUŞADASI-AYDIN",
    "Kuşadası Çmert": "KUŞADASI CMERT",
}


# ------------------ CSV Reader ------------------
@dataclass
class OrderRow:
    stok_kodu: str
    miktar: float
    grup: str


class CsvOrderReader:
    def __init__(self, csv_path: str):
        self.csv_path = csv_path
        self.df = None  # type: Optional[pd.DataFrame]

    def load(self) -> None:
        # Most inputs start with two header lines, so header=2; fall back to 0
        try:
            self.df = pd.read_csv(self.csv_path, encoding="utf-8", delimiter=",", header=2)
        except Exception:
            self.df = pd.read_csv(self.csv_path, encoding="utf-8", delimiter=",", header=0)

    def _find_col(self, poss: Iterable[str]) -> Optional[str]:
        assert self.df is not None
        cols_up = {TextNormalizer.up(c): c for c in self.df.columns}
        for p in poss:
            up = TextNormalizer.up(p)
            if up in cols_up:
                return cols_up[up]
        # try partial
        for p in poss:
            up = TextNormalizer.up(p)
            for cu, orig in cols_up.items():
                if up in cu:
                    return orig
        return None

    def get_branch_name(self) -> Optional[str]:
        assert self.df is not None
        # Possible header names for branch in many CSV exports
        candidates = [
            "SUBE", "SUBE ADI", "SUBEADI", "BAYI", "BAYI ADI", "FIRMA ADI",
        ]
        for c in candidates:
            if c in self.df.columns:
                val = str(self.df[c].iloc[0]) if len(self.df[c]) else ""
                return val
        # Sometimes present in a separate first header row; best-effort parse
        return None

    def iter_rows(self) -> Iterable[OrderRow]:
        assert self.df is not None
        stok = self._find_col(["STOK KODU", "STOKKODU", "STOK KOD", "KOD"])
        miktar = self._find_col(["MIKTAR", "MİKTAR", "ADET"])
        grup = self._find_col(["GRUP", "GRUP ADI", "KATEGORI", "KATEGORI ADI"])
        if not stok or not miktar:
            raise ValueError("CSV'de 'Stok Kodu' veya 'Miktar' sütunu bulunamadı.")
        if not grup:
            # default group if unavailable
            grup = "TATLI"
        for _, r in self.df.iterrows():
            try:
                mikt = float(str(r[miktar]).replace(",", "."))
            except Exception:
                try:
                    mikt = float(r[miktar])
                except Exception:
                    continue
            if pd.isna(mikt) or mikt == 0:
                continue
            yield OrderRow(stok_kodu=str(r[stok]), miktar=mikt, grup=str(r[grup]) if grup in r else "TATLI")


# ------------------ Branch Decision Engine ------------------
class BranchDecisionEngine:
    def __init__(self, branch_name: Optional[str]):
        # Apply branch name mapping for special cases
        self.branch_name = self._apply_branch_mapping(branch_name or "")
        self.branch_up = TextNormalizer.up(self.branch_name)
    
    @staticmethod
    def _apply_branch_mapping(branch_name: str) -> str:
        """Apply special branch name mappings (e.g., HARMANDALI→EFESUS)."""
        if not branch_name:
            return branch_name
        
        branch_up = TextNormalizer.up(branch_name)
        
        # Check exact matches first
        if branch_up in BRANCH_NAME_MAPPING:
            return BRANCH_NAME_MAPPING[branch_up]
        
        # Check if any mapping key is contained in the branch name
        for csv_name, excel_name in BRANCH_NAME_MAPPING.items():
            if csv_name in branch_up:
                return excel_name
        
        return branch_name

    def segment(self) -> str:
        up = self.branch_up
        if any(h in up for h in KUŞADASI_HINTS):
            return "KUSADASI"
        if any(h in up for h in IZMIR_BRANCH_HINTS):
            return "IZMIR"
        return "GENEL"
    
    def requires_day_selection(self) -> bool:
        """Check if this branch appears in multiple sheets (needs day selection)."""
        return self.branch_up in MULTI_DAY_BRANCHES
    
    def get_possible_sheets(self) -> List[str]:
        """Get list of possible sheet names for this branch."""
        return MULTI_DAY_BRANCHES.get(self.branch_up, [])


# ------------------ Writers ------------------
class BaseExcelWriter:
    def __init__(self, output_path: str):
        self.output_path = output_path
        self.wb: Optional[openpyxl.Workbook] = None
        self.ws: Optional[openpyxl.worksheet.worksheet.Worksheet] = None

    def load(self) -> None:
        if os.path.exists(self.output_path):
            self.wb = openpyxl.load_workbook(self.output_path)
            self.ws = self.wb.active
        else:
            self.wb = Workbook()
            self.ws = self.wb.active
            # basic header
            self.ws.cell(row=1, column=1, value="Sevkiyat Tarihi")
            self.ws.cell(row=2, column=1, value=datetime.today().strftime('%d.%m.%Y'))
            self.ws.cell(row=3, column=1, value="Stok Kodu")
            self.ws.cell(row=3, column=2, value="Miktar")
            self.ws.cell(row=3, column=3, value="Grup")
        
    def save(self) -> None:
        assert self.wb is not None
        self.wb.save(self.output_path)

    def clear_values(self) -> int:
        # Generic clear: from DATA_START_ROW onward, all sheets, values only (keep formulas and headers)
        assert self.wb is not None
        cleared = 0
        for ws in self.wb.worksheets:
            for row in ws.iter_rows(min_row=DATA_START_ROW, max_row=ws.max_row):
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith('='):
                        continue
                    if cell.value not in (None, ""):
                        cell.value = None
                        cleared += 1
        return cleared


class SimpleListWriter(BaseExcelWriter):
    """Fallback writer for formats we don't have a template for.
    Appends rows under headers: Stok Kodu | Miktar | Grup
    """

    def append_rows(self, items: Iterable[OrderRow]) -> int:
        assert self.ws is not None
        # Find first empty row after header
        start_row = max(DATA_START_ROW, 4)
        row = start_row
        written = 0
        for it in items:
            # Direct write without safe_write to avoid unmerge-remerge issues
            # Note: We get the cell object first, then set value
            # This prevents "MergedCell" attribute errors
            cell1 = self.ws.cell(row=row, column=1)
            cell2 = self.ws.cell(row=row, column=2)
            cell3 = self.ws.cell(row=row, column=3)
            
            # Check if cells are merged (MergedCell type)
            from openpyxl.cell.cell import MergedCell
            if not isinstance(cell1, MergedCell):
                cell1.value = it.stok_kodu
            if not isinstance(cell2, MergedCell):
                cell2.value = it.miktar
            if not isinstance(cell3, MergedCell):
                cell3.value = it.grup
            
            row += 1
            written += 1
        return written


class LojistikTemplateWriter(BaseExcelWriter):
    """Template-aware writer for the logistics sheet like in lojistik.png.
    - Columns are branch names across the top (row 1 or 2)
    - Rows below are blank lines where free-form item text is appended under the branch column
    """

    def __init__(self, output_path: str, sheet_name: Optional[str] = None):
        super().__init__(output_path)
        self.sheet_name = sheet_name

    def load(self) -> None:
        super().load()
        assert self.wb is not None and self.ws is not None
        # Eğer belirli bir sayfa istenmişse ve varsa, ona geç
        if self.sheet_name and self.sheet_name in [ws.title for ws in self.wb.worksheets]:
            self.ws = self.wb[self.sheet_name]

    def _find_or_add_branch_col(self, branch_name: str) -> int:
        assert self.ws is not None
        up = TextNormalizer.up(branch_name)
        # search first 2 rows, allow fuzzy contains
        best_c = None
        best_score = 0
        for r in (1, 2):
            for c in range(1, self.ws.max_column + 1):
                v = self.ws.cell(row=r, column=c).value
                if not v:
                    continue
                vv = TextNormalizer.up(str(v))
                if vv == up:
                    return c
                # fuzzy score: intersection length
                common = len(set(vv.split()) & set(up.split()))
                if common > best_score:
                    best_score = common
                    best_c = c
        if best_c is not None and best_score > 0:
            return best_c
        # not found: add new column at end of row 1
        col = self.ws.max_column + 1
        self.ws.cell(row=1, column=col, value=branch_name)
        return col

    def append_text_items(self, branch_name: str, items: Iterable[str]) -> int:
        assert self.ws is not None
        col = self._find_or_add_branch_col(self._canonical_branch(branch_name))
        # find first empty row below header, skipping merged regions
        def in_merge(r: int, c: int):
            for mr in self.ws.merged_cells.ranges:
                min_row, min_col, max_row, max_col = mr.bounds
                if min_row <= r <= max_row and min_col <= c <= max_col:
                    return (min_row, min_col, max_row, max_col)
            return None
        row = 2
        while True:
            bounds = in_merge(row, col)
            if bounds is not None:
                _, _, max_row, _ = bounds
                row = max_row + 1
                continue
            v = self.ws.cell(row=row, column=col).value
            if v in (None, ""):
                break
            row += 1
        written = 0
        for t in items:
            self.ws.cell(row=row, column=col, value=str(t).strip())
            row += 1
            written += 1
        return written

    def _canonical_branch(self, branch_name: str) -> str:
        up = TextNormalizer.up(branch_name)
        mapping = {
            "GÜZELBAHÇE": "GÜZELBAHÇE", "GUZELBAHCE": "GÜZELBAHÇE",
            "FORUM": "FORUM",
            "URLA": "URLA",
            "ILICA": "ILICA", "ILIÇA": "ILICA", "ILIÇA": "ILICA",
            "SEFERIHISAR": "SEFERIHISAR", "SEFERİHİSAR": "SEFERIHISAR",
            # İzmir listesiyle uyumlu ek varyantlar
            "MAVIBAHCE": "MAVIBAHCE", "MAVIBAHE": "MAVIBAHCE",
            "POINT": "POINT",
            "FOLKART": "FOLKART",
            "EFESUS": "EFESUS",
            "GAZIEMIR": "GAZIEMIR", "GAZİEMİR": "GAZIEMIR",
            "BALCOVA": "BALCOVA", "BALÇOVA": "BALCOVA",
            "HATAY": "HATAY",
            "FOLKART VEGA": "FOLKART VEGA", "VEGA": "FOLKART VEGA",
            "ISTASYON": "ISTASYON", "İSTASYON": "ISTASYON",
        }
        return mapping.get(up, branch_name)


class ImprovedLojistikWriter(LojistikTemplateWriter):
    """Improved lojistik writer with better branch matching and sheet selection."""
    
    def load(self) -> None:
        super().load()
        assert self.wb is not None and self.ws is not None
        
        # If sheet_hint is provided, try to find matching sheet
        if self.sheet_name:
            # Try exact match first
            if self.sheet_name in [ws.title for ws in self.wb.worksheets]:
                self.ws = self.wb[self.sheet_name]
                return
            
            # Try fuzzy matching for sheet names
            best_sheet = None
            best_score = 0
            hint_up = TextNormalizer.up(self.sheet_name)
            
            for ws in self.wb.worksheets:
                ws_up = TextNormalizer.up(ws.title)
                # Check for keyword matches
                hint_words = set(hint_up.split())
                ws_words = set(ws_up.split())
                common = len(hint_words & ws_words)
                
                if common > best_score:
                    best_score = common
                    best_sheet = ws
            
            if best_sheet and best_score > 0:
                self.ws = best_sheet
                return
        
        # If no specific sheet found, use the first available sheet
        if self.wb.worksheets:
            self.ws = self.wb.worksheets[0]
    
    def _find_sheet_for_branch(self, branch_name: str):
        """Find the Excel sheet that contains a column matching branch_name.
        Respects pre-selected sheet from day selection (self.ws already set in load())."""
        assert self.wb is not None
        
        # CRITICAL: If user selected a specific sheet via day selection,
        # don't override it by searching all sheets. Check if branch exists in current sheet first.
        if self.sheet_name and self.ws:
            branch_up = TextNormalizer.up(branch_name)
            # Check if branch exists in the currently selected sheet
            for r in range(1, 4):
                for c in range(1, min(self.ws.max_column + 1, 30)):
                    val = self.ws.cell(r, c).value
                    if not val:
                        continue
                    val_up = TextNormalizer.up(str(val))
                    
                    # Check if branch matches (exact or partial)
                    if branch_up == val_up or branch_up in val_up or val_up in branch_up:
                        # Branch found in user-selected sheet, use it!
                        return self.ws
            # If branch NOT found in user-selected sheet, log warning but respect user choice
            # This handles cases where user explicitly chose a sheet for multi-day branches
            return self.ws
        
        # No user selection: search all sheets for branch match
        branch_up = TextNormalizer.up(branch_name)
        for ws in self.wb.worksheets:
            # Check first 3 rows for branch headers
            for r in range(1, 4):
                for c in range(1, min(ws.max_column + 1, 30)):
                    val = ws.cell(r, c).value
                    if not val:
                        continue
                    val_up = TextNormalizer.up(str(val))
                    
                    # Check if branch matches (exact or partial)
                    if branch_up == val_up or branch_up in val_up or val_up in branch_up:
                        return ws
        
        # If not found, return current sheet
        return self.ws

    def _find_or_add_branch_col(self, branch_name: str) -> int:
        assert self.ws is not None
        up = TextNormalizer.up(branch_name)
        
        # Enhanced branch matching with better fuzzy logic
        best_c = None
        best_score = 0
        
        # Search in first 3 rows for branch headers
        for r in range(1, min(4, self.ws.max_row + 1)):
            for c in range(1, self.ws.max_column + 1):
                v = self.ws.cell(row=r, column=c).value
                if not v:
                    continue
                    
                vv = TextNormalizer.up(str(v))
                
                # Exact match
                if vv == up:
                    return c
                
                # Check if branch name is contained in cell value or vice versa
                if up in vv or vv in up:
                    return c
                
                # Fuzzy matching with word intersection
                hint_words = set(up.split())
                cell_words = set(vv.split())
                common = len(hint_words & cell_words)
                
                if common > best_score and common > 0:
                    best_score = common
                    best_c = c
        
        # If found a good match, use it
        if best_c is not None and best_score > 0:
            return best_c
        
        # If no match found, add new column
        col = self.ws.max_column + 1
        self.ws.cell(row=1, column=col, value=branch_name)
        return col

    def append_text_items(self, branch_name: str, items: Iterable[str]) -> int:
        assert self.ws is not None and self.wb is not None
        
        # Use canonical branch name for better matching
        canonical_branch = self._canonical_branch(branch_name)
        
        # Find the correct sheet for this branch
        correct_sheet = self._find_sheet_for_branch(canonical_branch)
        if correct_sheet != self.ws:
            self.ws = correct_sheet
        
        col = self._find_or_add_branch_col(canonical_branch)
        
        # Find first empty row below header, skipping merged regions
        def in_merge(r: int, c: int):
            for mr in self.ws.merged_cells.ranges:
                min_row, min_col, max_row, max_col = mr.bounds
                if min_row <= r <= max_row and min_col <= c <= max_col:
                    return (min_row, min_col, max_row, max_col)
            return None
        
        # Start from row 3 to skip potential headers
        row = 3
        while row <= self.ws.max_row:
            bounds = in_merge(row, col)
            if bounds is not None:
                _, _, max_row, _ = bounds
                row = max_row + 1
                continue
            
            v = self.ws.cell(row=row, column=col).value
            if v in (None, "", " "):
                break
            row += 1
        
        # If we've reached the end, we might need to add more rows
        if row > self.ws.max_row:
            # Ensure we have enough rows
            for _ in range(10):  # Add 10 empty rows
                self.ws.cell(row=row, column=1, value=None)
                row += 1
        
        written = 0
        for t in items:
            if t and str(t).strip():
                cell = self.ws.cell(row=row, column=col)
                cell.value = str(t).strip()
                
                # Enable text wrapping for long product names
                cell.alignment = openpyxl.styles.Alignment(wrap_text=True)
                
                row += 1
                written += 1
        
        return written


# ------------------ Coordinator ------------------
class ShipmentCoordinator:
    """Coordinates parsing and writing into three workbooks following the design document.
    - sevkiyat_tatli.xlsx
    - sevkiyat_donuk.xlsx
    - sevkiyat_lojistik.xlsx
    """

    def __init__(self):
        pass

    def process_tatli(self, csv_path: str, output_path: str = "sevkiyat_tatlı.xlsx", sheet_hint: Optional[str] = None) -> Tuple[int, int]:
        """Delegate to existing advanced matcher if available; fall back to simple list."""
        try:
            from parse_gptfix import process_csv as legacy_tatli
            # legacy returns None or counts; normalize to (matched, unmatched)
            # CRITICAL: Pass sheet_hint to legacy_tatli for multi-day branch support
            res = legacy_tatli(csv_path, output_path=output_path, sheet_name=sheet_hint)
            if isinstance(res, tuple) and len(res) == 2:
                return res
            return (0, 0)
        except Exception:
            # Fallback: append all rows where group includes TATLI
            rdr = CsvOrderReader(csv_path)
            rdr.load()
            items = [r for r in rdr.iter_rows() if "TATLI" in TextNormalizer.up(r.grup)]
            wr = SimpleListWriter(output_path)
            wr.load()
            wr.append_rows(items)
            wr.save()
            return (len(items), 0)

    def process_donuk(self, csv_path: str, output_path: str = "sevkiyat_donuk.xlsx", sheet_hint: Optional[str] = None) -> Tuple[int, int]:
        try:
            from parse_gptfix import process_donuk_csv as legacy_donuk
            res = legacy_donuk(csv_path, output_path=output_path, sheet_name=sheet_hint)
            if isinstance(res, tuple) and len(res) == 2:
                return res
            # Fallback: estimate
            rdr = CsvOrderReader(csv_path)
            rdr.load()
            items = [r for r in rdr.iter_rows() if any(k in TextNormalizer.up(r.grup) for k in ["DONDURMA", "PASTA", "BOREK", "TATLI"])]
            return (len(items), 0)
        except Exception:
            rdr = CsvOrderReader(csv_path)
            rdr.load()
            items = [r for r in rdr.iter_rows() if any(k in TextNormalizer.up(r.grup) for k in ["DONDURMA", "PASTA", "BOREK", "TATLI"])]
            wr = SimpleListWriter(output_path)
            wr.load()
            wr.append_rows(items)
            wr.save()
            return (len(items), 0)

    def process_lojistik(self, csv_path: str, output_path: str = "sevkiyat_lojistik.xlsx", sheet_hint: Optional[str] = None) -> Tuple[int, int]:
        rdr = CsvOrderReader(csv_path)
        rdr.load()

        def read_branch_from_file(path: str) -> Optional[str]:
            try:
                with open(path, encoding="utf-8") as f:
                    for line in f:
                        up = TextNormalizer.up(line)
                        # Match "SUBE KODU", "SUBE KODU-ADI", "SUBE ADI", etc.
                        if "SUBE" in up and ("KODU" in up or "ADI" in up):
                            raw = line.split(":", 1)[-1] if ":" in line else line
                            # e.g. "242 - BALÇOVA" or "187 - ANKARA(KIZILAY)"
                            part = raw.split("-", 1)[-1] if "-" in raw else raw
                            part = part.strip()
                            # Remove quotes if present
                            part = part.strip('"').strip("'").strip()
                            # If parens exist, prefer the inner content (ANKARA(KIZILAY) -> KIZILAY)
                            import re
                            m = re.search(r"\(([^)]+)\)", part)
                            if m:
                                return m.group(1).strip()
                            # remove trailing DEPO
                            if part.upper().endswith(" DEPO"):
                                part = part[:-5].strip()
                            return part
            except Exception:
                pass
            return None

        branch_name = read_branch_from_file(csv_path) or rdr.get_branch_name() or "GENEL"
        # Groups mapped as per design doc: SARF MALZEME, KURABIYE, CIKOLATA - HEDIYELIK, ICECEK
        include_keys = ["SARF", "KURABIYE", "CIKOLATA", "HEDIYELIK", "ICECEK", "İCECEK", "İÇECEK"]
        rows = [r for r in rdr.iter_rows() if any(k in TextNormalizer.up(r.grup) for k in include_keys)]

        def clean_display(stok: str) -> str:
            import re
            s = str(stok)
            # remove {...}
            s = re.sub(r"\{[^}]*\}", "", s)
            
            # Clean parentheses content - keep only for specific cases
            # Keep: T-shirt sizes (S, M, L, XL, XXL), colors, and sauce types
            def should_keep_parens(content: str) -> bool:
                content_up = content.upper()
                # Keep if contains t-shirt sizes
                if any(size in content_up for size in ['S)', 'M)', 'L)', 'XL)', 'XXL)', 'BEDEN']):
                    return True
                # Keep if contains color indicators
                if any(color in content_up for color in ['BEYAZ', 'SİYAH', 'MAVİ', 'KIRMIZI', 'YEŞİL', 'SARI', 'GRİ', 'KAHVE', 'MOR', 'TURUNCU', 'PEMBE', 'RENK']):
                    return True
                # Keep if it's sauce type (contains SOS and a product name)
                if 'SOS' in content_up and any(prod in content_up for prod in ['CHOCOLATE', 'CIKOLATA', 'KARAMEL', 'FRAMBUAZ', 'CILEKTE', 'ÇİLEK', 'ORMAN', 'MEYVELI', 'FISTIK', 'ANTEP']):
                    return True
                return False
            
            # Process parentheses: remove unwanted, keep wanted
            def clean_parens(text: str) -> str:
                result = text
                # Find all parentheses content
                pattern = r'\([^)]*\)'
                matches = re.finditer(pattern, text)
                
                for match in reversed(list(matches)):  # Reverse to maintain positions
                    content = match.group(0)
                    inner = content[1:-1]  # Remove outer parens
                    
                    if should_keep_parens(inner):
                        # Keep this parenthesis
                        continue
                    else:
                        # Remove this parenthesis
                        result = result[:match.start()] + result[match.end():]
                
                return result
            
            s = clean_parens(s)
            s = re.sub(r"\s+", " ", s).strip()
            return s

        # Convert to text lines like: "<NAME (…)> - <qty> ADET"
        lines: List[str] = []
        for r in rows:
            name = clean_display(r.stok_kodu)
            qty = int(r.miktar) if float(r.miktar).is_integer() else r.miktar
            lines.append(f"{name} - {qty} ADET")
        
        # Use improved lojistik writer
        wr = ImprovedLojistikWriter(output_path, sheet_name=sheet_hint)
        wr.load()
        count = wr.append_text_items(branch_name, lines)
        wr.save()
        return (count, 0)

    def run_all(self, csv_path: str, izmir_day_sheet: Optional[str] = None) -> Dict[str, Dict[str, int]]:
        out = {}
        matched, unmatched = self.process_tatli(csv_path, output_path="sevkiyat_tatlı.xlsx", sheet_hint=izmir_day_sheet)
        out["tatli"] = {"matched": matched, "unmatched": unmatched, "file": "sevkiyat_tatlı.xlsx"}
        m2, u2 = self.process_donuk(csv_path, output_path="sevkiyat_donuk.xlsx", sheet_hint=izmir_day_sheet)
        out["donuk"] = {"matched": m2, "unmatched": u2, "file": "sevkiyat_donuk.xlsx"}
        m3, u3 = self.process_lojistik(csv_path, output_path="sevkiyat_lojistik.xlsx", sheet_hint=izmir_day_sheet)
        out["lojistik"] = {"matched": m3, "unmatched": u3, "file": "sevkiyat_lojistik.xlsx"}
        return out


# ------------------ Utilities for GUI ------------------

def clear_workbook_values(path: str) -> int:
    """Generic clear - clears all non-formula values. Use specific clear functions instead."""
    if not os.path.exists(path):
        return 0
    wb = openpyxl.load_workbook(path)
    cleared = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=DATA_START_ROW, max_row=ws.max_row):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    continue
                if cell.value not in (None, ""):
                    cell.value = None
                    cleared += 1
    wb.save(path)
    return cleared


def clear_tatli_values(path: str) -> int:
    """Clear tatlı file data cells and sepet (basket) counts in row 1.
    
    Clears:
    1. Row 1 sepet values (basket counts) for each branch
    2. Data cells (TEPSI, ADET columns) starting from row 3
    
    Preserves:
    - Product names (column 1)
    - Headers (row 2)
    - Formulas
    - Merged cell structures
    """
    if not os.path.exists(path):
        return 0
    
    wb = openpyxl.load_workbook(path)
    cleared = 0
    
    for ws in wb.worksheets:
        # Step 1: Read branch columns from row 2 headers
        subeler = {}
        for cell in ws[2][1:]:  # Skip first column (product names)
            if cell.value:
                sube_ad = str(cell.value).strip()
                # Each branch has 4 columns: TEPSI, TEPSI_2, ADET, ADET_2
                subeler[sube_ad] = {
                    "tepsi": cell.column,
                    "tepsi_2": cell.column + 1,
                    "adet": cell.column + 2,
                    "adet_2": cell.column + 3
                }
        
        # Step 2: Clear row 1 sepet values for each branch
        # Sepet is written to the first column of each branch (TEPSI column)
        for sube in subeler.values():
            sepet_col = sube["tepsi"]
            
            # Find the cell to clear (handle merged cells properly)
            # mr.bounds returns (min_col, min_row, max_col, max_row)
            target_cell = ws.cell(row=1, column=sepet_col)
            
            # Check if this cell is part of a merged range
            for mr in ws.merged_cells.ranges:
                # bounds format: (min_col, min_row, max_col, max_row)
                if (mr.min_row <= 1 <= mr.max_row) and (mr.min_col <= sepet_col <= mr.max_col):
                    # This cell is in a merged range, use the master (top-left)
                    target_cell = ws.cell(row=mr.min_row, column=mr.min_col)
                    break
            
            # Clear value if it exists
            if target_cell.value not in (None, ""):
                target_cell.value = None
                cleared += 1
        
        # Step 3: Clear data cells (rows 3+)
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=1):
            ana_cell = row[0]
            if not ana_cell.value:
                continue
            
            # Skip special rows (headers like "SIPARIS TARIHI", etc.)
            ana_ad = str(ana_cell.value).upper()
            skip_keywords = ["SIPARIS TARIHI", "SIPARIS ALAN", "TESLIM TARIHI", "TEYID EDEN"]
            if any(ana_ad.startswith(k) or ana_ad == k for k in skip_keywords):
                continue
            
            # Clear TEPSI, TEPSI_2, ADET, ADET_2 cells for this product
            for sube in subeler.values():
                for col in [sube["tepsi"], sube["tepsi_2"], sube["adet"], sube["adet_2"]]:
                    cell = ws.cell(row=ana_cell.row, column=col)
                    
                    # Skip formulas
                    if isinstance(cell.value, str) and cell.value.startswith('='):
                        continue
                    
                    # Clear value
                    if cell.value not in (None, ""):
                        cell.value = None
                        cleared += 1
    
    wb.save(path)
    return cleared


def clear_lojistik_values(path: str) -> int:
    """Clear only data cells in lojistik file, preserving branch headers and yellow-highlighted cells.
    
    Yellow-highlighted cells contain permanent items (demirbaş) that should never be cleared.
    """
    if not os.path.exists(path):
        return 0
    
    wb = openpyxl.load_workbook(path)
    cleared = 0
    
    for ws in wb.worksheets:
        # Find branch columns in first 2 rows
        branch_cols = set()
        for r in range(1, 3):  # rows 1-2
            for c in range(1, ws.max_column + 1):
                val = ws.cell(row=r, column=c).value
                if val and str(val).strip():
                    # This is a branch header column
                    branch_cols.add(c)
        
        # Clear data rows (from row 3 onwards) in branch columns only
        for c in branch_cols:
            for r in range(3, ws.max_row + 1):
                cell = ws.cell(row=r, column=c)
                
                # Skip formulas
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    continue
                
                # CRITICAL: Skip yellow-highlighted cells (demirbaş items)
                # Check if cell has yellow fill color
                if cell.fill and cell.fill.start_color:
                    # Yellow colors: FFFF00 (pure yellow), FFFFFF00 (yellow), etc.
                    color_rgb = cell.fill.start_color.rgb if hasattr(cell.fill.start_color, 'rgb') else None
                    if color_rgb:
                        # Check if it's a yellow color (starts with FFFF or 00FFFF)
                        color_str = str(color_rgb).upper()
                        # Yellow variations: FFFFFF00, FFFF00, FFFFE0 (light yellow), etc.
                        if 'FFFF' in color_str[:6] or color_str.startswith('00FFFF'):
                            continue  # Skip this cell - it's yellow highlighted
                
                # Clear non-empty cells
                if cell.value not in (None, ""):
                    cell.value = None
                    cleared += 1
    
    wb.save(path)
    return cleared


def clear_donuk_values(path: str) -> int:
    """Clear only quantity cells in donuk file, preserving product names and structure.
    
    Uses a simple heuristic approach:
    1. Skip first 2 rows (headers)
    2. Clear numeric cells (integers, floats)
    3. For text cells containing products with quantities (e.g., "KÜNEFE = 5"), 
       extract and keep only the product name
    """
    if not os.path.exists(path):
        return 0
    
    import re
    wb = openpyxl.load_workbook(path)
    cleared = 0
    
    for ws in wb.worksheets:
        # Process all cells starting from row 3
        for r in range(3, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(row=r, column=c)
                
                # Skip formulas
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    continue
                
                # Case 1: Pure numeric value - clear it
                if isinstance(cell.value, (int, float)):
                    cell.value = None
                    cleared += 1
                    continue
                
                # Case 2: String that can be parsed as number - clear it
                if isinstance(cell.value, str) and cell.value.strip():
                    try:
                        float(cell.value.replace(",", "."))
                        cell.value = None
                        cleared += 1
                        continue
                    except:
                        pass
                    
                    # Case 3: Text containing product name with quantity
                    # Pattern: "KÜNEFE = 5", "KÜNEFE  5", "ŞERBET = 10", "CITIR MANTI  3"
                    # We want to keep only the product name part
                    if "=" in cell.value or re.search(r'\s+\d+', cell.value):
                        # Check if it contains known product keywords
                        val_up = TextNormalizer.up(cell.value)
                        
                        # All DONUK products that need quantity cleanup
                        donuk_products = [
                            "KUNEFE", "SERBET", "TRILECE", "ROKOKO", "EKLER", "MEYVELI",
                            # Çıtır Mantı ve sonrası ürünler
                            "CITIR MANTI", "ÇITIR MANTI", "MANTI",
                            "BOYOZ", "PATATES", 
                            "HAMBURGER", "KOFTE", "KÖFTE",
                            "TAVUK", "BUT",
                            "BAKLAVA", "SOGUK", "CEVIZ", "TAHIN",
                            "DANA", "ASADO",
                            "BONFILE", "MADALYON",
                            "SPAGETTI", "MAKARON",
                            "USTANIN KOFTESI", "KADAYIF", "SNITSEL", "ŞINITSEL",
                            # Makaron varyantları
                            "CIKOLATALI", "ÇIKOLATALI", "HINDCEVIZLI", "HIND.CEVIZLI", "HIND CEV",
                            "KARAMEL", "FRAMBUAZ", "ANTEPL", "ANTEPLI", "MERSIL", "YBNMERSIL"
                        ]
                        
                        if any(prod in val_up for prod in donuk_products):
                            # Extract product name (text before = or trailing numbers)
                            match = re.match(r'^([^=]+?)(?:\s*=\s*\d+|\s+\d+\s*$)', cell.value)
                            if match:
                                product_name = match.group(1).strip()
                                if product_name and product_name != cell.value:
                                    cell.value = product_name
                                    cleared += 1
    
    wb.save(path)
    return cleared


def format_today_in_workbook(path: str) -> None:
    if not os.path.exists(path):
        return
    wb = openpyxl.load_workbook(path)
    for ws in wb.worksheets:
        # Merge-safe write to A2
        target_r, target_c = 2, 1
        cell = ws.cell(row=target_r, column=target_c)
        def master(sheet, r, c):
            for mr in sheet.merged_cells.ranges:
                min_row, min_col, max_row, max_col = mr.bounds
                if min_row <= r <= max_row and min_col <= c <= max_col:
                    return sheet.cell(row=min_row, column=min_col)
            return sheet.cell(row=r, column=c)
        m = master(ws, target_r, target_c)
        m.value = datetime.today().strftime('%d.%m.%Y')
    wb.save(path)
