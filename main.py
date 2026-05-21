import os
import re
import time
from dataclasses import dataclass
from datetime import datetime
from typing import Optional, List, Dict, Tuple

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# ============================================================
# GENERAL CONFIGURATION
# ============================================================

# Supported Excel formats. Old .xls files require xlrd.
EXCEL_EXTS = (".xls", ".xlsx", ".xlsm")

# Archive index cache prevents scanning the whole archive every time.
DEFAULT_INDEX_FILENAME = "calibration_archive_index.csv"
DEFAULT_OUTPUT_DIRNAME = "calibration_automation_output"

# Possible sheet names for calibration result tables.
# If none match, the script will scan all sheets.
RESULT_SHEET_CANDIDATES = [
    "results",
    "result",
    "calibration results",
    "measurement results",
    "test results",
    "data",
    "measurements",
    "calibration data",
]

# Possible sheet names for general certificate information.
INFO_SHEET_CANDIDATES = [
    "cover",
    "summary",
    "certificate",
    "information",
    "info",
    "general",
    "overview",
    "front page",
]

# Date labels that may appear in certificate sheets.
DATE_KEYS = [
    "date",
    "issue date",
    "measurement date",
    "calibration date",
    "date of issue",
    "date of measurement",
    "date of calibration",
]

# Generic metadata fields and possible labels used in certificates.
METADATA_KEYS = {
    "certificate_no": [
        "certificate no",
        "certificate number",
        "certificate id",
        "report no",
        "report number",
    ],
    "job_no": [
        "job no",
        "job number",
        "order no",
        "order number",
        "work order",
        "reference no",
        "reference number",
    ],
    "client": [
        "client",
        "customer",
        "company",
        "organization",
    ],
    "instrument": [
        "instrument",
        "equipment",
        "device",
        "item",
        "unit under test",
        "uut",
    ],
    "model": [
        "model",
        "type",
        "part number",
        "item code",
    ],
    "manufacturer": [
        "manufacturer",
        "maker",
        "brand",
    ],
    "serial": [
        "serial",
        "serial number",
        "s/n",
        "id number",
        "asset id",
        "equipment id",
        "device id",
    ],
}

# Keywords used to detect result-table columns.
# Reference, reading and error are required.
HEADER_KEYWORDS = {
    "point": [
        "point",
        "no",
        "number",
        "step",
        "index",
    ],
    "reference": [
        "reference",
        "reference value",
        "standard value",
        "nominal",
        "set point",
        "ref",
    ],
    "reading": [
        "reading",
        "measured",
        "measured value",
        "indication",
        "displayed value",
        "measurement",
    ],
    "error": [
        "error",
        "deviation",
        "difference",
        "correction",
        "offset",
    ],
    "uncertainty": [
        "uncertainty",
        "expanded uncertainty",
        "u",
        "coverage",
    ],
}

# Files with these words are probably not final certificates.
NON_CERTIFICATE_KEYWORDS = [
    "template",
    "backup",
    "index",
    "database",
    "draft",
    "temporary",
    "~$",
]

# ============================================================
# USER INPUT HELPERS
# ============================================================

# Repeatedly asks the user for a valid file or folder path.
def ask_path(prompt: str, must_exist: bool = True, allow_empty: bool = False) -> str:
    """Ask for a file or folder path."""
    while True:
        p = input(prompt).strip().strip('"').strip("'")

        if not p:
            if allow_empty:
                return ""
            print("Path cannot be empty.")
            continue

        if must_exist and not os.path.exists(p):
            print(f"Path not found: {p}")
            continue

        return p


# Handles yes/no terminal input and applies a default answer when Enter is pressed.
def ask_yes_no(prompt: str, default: bool = True) -> bool:
    """Ask a yes/no question with a default answer."""
    suffix = "Y/n" if default else "y/N"

    while True:
        s = input(f"{prompt} ({suffix}): ").strip().lower()

        if not s:
            return default

        if s in ("y", "yes"):
            return True

        if s in ("n", "no"):
            return False

        print("Please enter y or n.")


# Reads a numeric input from the user and also accepts comma-based decimals.
def ask_float(prompt: str, default: float) -> float:
    """Ask for a numeric value; comma decimals are accepted."""
    while True:
        s = input(f"{prompt} [default={default}]: ").strip()

        if not s:
            return default

        try:
            return float(s.replace(",", "."))
        except ValueError:
            print("Please enter a valid number.")

# ============================================================
# EXCEL READING
# ============================================================

# Selects the correct pandas engine required to read each Excel file type.
def get_excel_engine(path: str) -> Optional[str]:
    """Choose the pandas Excel engine from the file extension."""
    lower = path.lower()

    if lower.endswith(".xlsx") or lower.endswith(".xlsm"):
        return "openpyxl"

    if lower.endswith(".xls"):
        return "xlrd"

    return None

# Reads an Excel sheet safely and gives clear dependency errors when needed.
def read_excel_any(path: str, sheet_name=None, header=None):
    """
    Read an Excel sheet with the correct engine.
    Extra error messages are included because missing xlrd/openpyxl is common.
    """
    engine = get_excel_engine(path)

    try:
        return pd.read_excel(path, sheet_name=sheet_name, header=header, engine=engine)

    except Exception as e:
        msg = str(e).lower()

        if path.lower().endswith(".xls") and ("xlrd" in msg or "missing optional dependency" in msg):
            raise RuntimeError(
                f"{path}\n\n"
                "This file is in .xls format. Install xlrd:\n"
                "    pip install xlrd==2.0.1\n"
                "or:\n"
                "    conda install -c conda-forge xlrd\n"
                f"\nOriginal error: {e}"
            )

        if (path.lower().endswith(".xlsx") or path.lower().endswith(".xlsm")) and "openpyxl" in msg:
            raise RuntimeError(
                f"{path}\n\n"
                "This file requires openpyxl:\n"
                "    pip install openpyxl\n"
                f"\nOriginal error: {e}"
            )

        raise

# Opens the workbook as an ExcelFile object so sheet names can be checked first.
def get_excel_file(path: str) -> pd.ExcelFile:
    """Open workbook so sheet names can be inspected."""
    return pd.ExcelFile(path, engine=get_excel_engine(path))

# Searches workbook sheet names and returns the first candidate match.
def pick_sheet(xls: pd.ExcelFile, candidates: List[str]) -> Optional[str]:
    """Find the first matching sheet name, case-insensitive."""
    sheet_map = {s.lower(): s for s in xls.sheet_names}

    for candidate in candidates:
        if candidate.lower() in sheet_map:
            return sheet_map[candidate.lower()]

    return None

# ============================================================
# GENERAL PARSING HELPERS
# ============================================================

def safe_mkdir(path: str):
    os.makedirs(path, exist_ok=True)


def normalize_text(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()


def tokenize_numeric_filename(name: str) -> List[str]:
    """Extract serial/certificate-like numeric tokens from file names."""
    return re.findall(r"\d{5,}", name)


def parse_date_from_text(s: str) -> Optional[datetime]:
    """Parse common date formats from a text cell."""
    s = str(s).strip()

    # Format: YYYY-MM-DD or YYYY/MM/DD
    m = re.search(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", s)
    if m:
        y, mo, d = map(int, m.groups())
        try:
            return datetime(y, mo, d)
        except ValueError:
            return None

    # Format: DD-MM-YYYY, DD/MM/YYYY, MM-DD-YYYY, etc.
    m = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)
    if m:
        a, b, y = map(int, m.groups())

        # If one number is greater than 12, it must be the day.
        if a > 12 and b <= 12:
            d, mo = a, b
        elif b > 12 and a <= 12:
            d, mo = b, a
        else:
            d, mo = a, b

        try:
            return datetime(y, mo, d)
        except ValueError:
            return None

    return None


def file_mtime_date(path: str) -> datetime:
    """Fallback date when no date is found inside the certificate."""
    return datetime.fromtimestamp(os.path.getmtime(path))


def snap_nominal_value(value: float, tolerance: float) -> float:
    """
    Convert near-nominal reference values into clean nominal values.
    Example: 24.99 -> 25.0 if tolerance is large enough.
    """
    if pd.isna(value):
        return np.nan

    value = float(value)
    rounded = round(value)

    if abs(value - rounded) <= tolerance:
        return float(rounded)

    return float(round(value, 3))


def looks_like_non_certificate_file(path: str) -> bool:
    """Skip templates, backups, drafts, databases and temp files."""
    lower = path.lower()
    name = os.path.basename(path).lower()

    return any(keyword in lower or keyword in name for keyword in NON_CERTIFICATE_KEYWORDS)


# ============================================================
# FAST ARCHIVE INDEXING
# ============================================================

def build_archive_index(root: str, cache_csv_path: str, force_rebuild: bool = False) -> pd.DataFrame:
    """
    Build or load a cached archive index.
    This is the main speed improvement for large folders or network drives.
    """
    if (not force_rebuild) and os.path.exists(cache_csv_path):
        try:
            df = pd.read_csv(cache_csv_path, dtype=str)
            if not df.empty:
                return df
        except Exception:
            # If cache is unreadable, rebuild it.
            pass

    rows = []
    stack = [root]

    # Iterative folder scan is safer for large directory trees.
    while stack:
        current_dir = stack.pop()

        try:
            with os.scandir(current_dir) as iterator:
                for entry in iterator:
                    try:
                        if entry.is_dir(follow_symlinks=False):
                            stack.append(entry.path)

                        elif entry.is_file(follow_symlinks=False):
                            lower_name = entry.name.lower()

                            if lower_name.endswith(EXCEL_EXTS):
                                tokens = tokenize_numeric_filename(entry.name)

                                try:
                                    stat = entry.stat()
                                    mtime = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                                    size = stat.st_size
                                except Exception:
                                    mtime = ""
                                    size = ""

                                rows.append({
                                    "path": entry.path,
                                    "name": entry.name,
                                    "directory": current_dir,
                                    "extension": os.path.splitext(entry.name)[1].lower(),
                                    "mtime": mtime,
                                    "size_bytes": size,
                                    "numeric_tokens": "|".join(tokens),
                                })

                    except Exception:
                        continue

        except Exception:
            continue

    df = pd.DataFrame(rows)

    if df.empty:
        df = pd.DataFrame(
            columns=[
                "path",
                "name",
                "directory",
                "extension",
                "mtime",
                "size_bytes",
                "numeric_tokens",
            ]
        )

    df.to_csv(cache_csv_path, index=False, encoding="utf-8-sig")
    return df

# Searches the cached archive index using extracted identifier tokens.
# This is much faster than rescanning the whole archive every run.
def search_archive_index(
    index_df: pd.DataFrame,
    identifiers: List[str],
    require_all: bool = True
) -> pd.DataFrame:
    """
    Search cached archive index by identifier tokens.
    If require_all=True, all identifiers must appear in the file name.
    """

    # Return immediately if the index is empty.
    if index_df.empty:
        return index_df.copy()

    # Clean identifiers and remove empty values.
    identifiers = [str(x).strip() for x in identifiers if str(x).strip()]

    if not identifiers:
        return index_df.iloc[0:0].copy()

    # Compare identifier list against tokenized file-name numbers.
    def match_tokens(token_str: str) -> bool:
        token_set = set(str(token_str).split("|")) if pd.notna(token_str) else set()

        # Strict mode: every identifier must exist in the filename.
        if require_all:
            return all(identifier in token_set for identifier in identifiers)

        # Relaxed mode: at least one identifier is enough.
        return any(identifier in token_set for identifier in identifiers)

    matched = index_df[index_df["numeric_tokens"].apply(match_tokens)].copy()

    # Sort by modification time to keep historical order cleaner.
    if "mtime" in matched.columns:
        matched = matched.sort_values("mtime")

    return matched.reset_index(drop=True)

# ============================================================
# IDENTIFIER EXTRACTION
# ============================================================

# Extracts possible identifiers directly from the file name.
def extract_identifiers_from_filename(path: str) -> List[str]:
    """Use filename numbers as possible identifiers."""
    return list(dict.fromkeys(tokenize_numeric_filename(os.path.basename(path))))

# Searches workbook contents for serial numbers or equipment identifiers.
def extract_identifiers_from_workbook(path: str) -> List[str]:
    """
    Extract possible serial/equipment identifiers from workbook contents.
    This is heuristic-based because certificate layouts may differ.
    """
    found = []

    try:
        xls = get_excel_file(path)
    except Exception:
        return found

    identifier_keywords = [
        "serial",
        "serial number",
        "s/n",
        "equipment id",
        "device id",
        "asset id",
        "instrument id",
        "unit id",
        "item id",
    ]

    for sheet in xls.sheet_names:
        try:
            df_raw = read_excel_any(path, sheet_name=sheet, header=None)
        except Exception:
            continue

        # Scan a limited top-left region for speed and because IDs are usually there.
        sub = df_raw.iloc[:500, :80].astype(str).fillna("")

        for i in range(sub.shape[0]):
            row_text = " ".join(sub.iloc[i].tolist()).lower()

            if any(keyword in row_text for keyword in identifier_keywords):
                nums = re.findall(r"\d{5,}", row_text)
                found.extend(nums)

        for i in range(sub.shape[0]):
            for j in range(sub.shape[1]):
                cell = sub.iat[i, j].lower()

                if any(keyword in cell for keyword in identifier_keywords):
                    for jj in range(j + 1, min(j + 10, sub.shape[1])):
                        nums = re.findall(r"\d{5,}", str(sub.iat[i, jj]))
                        found.extend(nums)

    found = list(dict.fromkeys(found))
    filtered = [x for x in found if 5 <= len(x) <= 12]

    return list(dict.fromkeys(filtered)) if filtered else found

# Chooses the best available identifier extraction method automatically.
def extract_identifiers_auto(path: str) -> List[str]:
    """Try workbook extraction first; fall back to filename."""
    ids = extract_identifiers_from_workbook(path)

    if ids:
        return ids

    return extract_identifiers_from_filename(path)

# ============================================================
# METADATA EXTRACTION
# ============================================================

def extract_date_from_workbook(path: str) -> Optional[datetime]:
    """
    Extract likely certificate/calibration date.
    If labeled date fields fail, the upper sheet area is scanned as fallback.
    """
    try:
        xls = get_excel_file(path)
    except Exception:
        return None

    # Prefer metadata-like sheets first, otherwise scan all sheets.
    preferred = pick_sheet(xls, INFO_SHEET_CANDIDATES)
    sheets = [preferred] if preferred else xls.sheet_names

    for sheet in sheets:
        try:
            df_raw = read_excel_any(path, sheet_name=sheet, header=None)
        except Exception:
            continue

        # Only scan the likely header/metadata region for speed.
        sub = df_raw.iloc[:350, :80].astype(str).fillna("")

        for i in range(sub.shape[0]):
            for j in range(sub.shape[1]):
                cell = sub.iat[i, j].lower().strip()

                if any(k in cell for k in DATE_KEYS):
                    neighbors = []

                    # Date values are usually close to their labels.
                    for di, dj in [(0, 1), (0, 2), (0, 3), (1, 0), (1, 1), (1, 2)]:
                        ii, jj = i + di, j + dj

                        if 0 <= ii < sub.shape[0] and 0 <= jj < sub.shape[1]:
                            neighbors.append(sub.iat[ii, jj])

                    for neighbor in neighbors:
                        dt = parse_date_from_text(neighbor)
                        if dt:
                            return dt

        # Fallback: search the upper-left area for any date-like value.
        for i in range(min(80, sub.shape[0])):
            for j in range(min(30, sub.shape[1])):
                dt = parse_date_from_text(sub.iat[i, j])
                if dt:
                    return dt

    return None

def extract_metadata_from_workbook(path: str) -> Dict[str, str]:
    """
    Extract certificate metadata using label-neighbor matching.
    Values are usually placed right next to or below labels.
    """
    meta = {
        "source_file": path,
        "certificate_no": "",
        "job_no": "",
        "client": "",
        "instrument": "",
        "model": "",
        "manufacturer": "",
        "serial": "",
    }

    try:
        xls = get_excel_file(path)
    except Exception:
        return meta

    # Metadata is usually stored in the first information/cover sheets.
    preferred = pick_sheet(xls, INFO_SHEET_CANDIDATES)
    sheets = [preferred] if preferred else xls.sheet_names[:3]

    for sheet in sheets:
        try:
            df_raw = read_excel_any(path, sheet_name=sheet, header=None)
        except Exception:
            continue

        # Limit the scan to the upper workbook area where certificate metadata usually exists.
        sub = df_raw.iloc[:350, :80].astype(str).fillna("")

        for key, aliases in METADATA_KEYS.items():
            if key not in meta:
                continue

            # Do not overwrite a metadata field once it has been found.
            if meta[key]:
                continue

            for i in range(sub.shape[0]):
                for j in range(sub.shape[1]):
                    cell = sub.iat[i, j].lower().strip()

                    if any(alias in cell for alias in aliases):
                        values = []

                        # Candidate values are usually to the right or just below the label.
                        for di, dj in [(0, 1), (0, 2), (0, 3), (1, 0), (1, 1), (1, 2)]:
                            ii, jj = i + di, j + dj

                            if 0 <= ii < sub.shape[0] and 0 <= jj < sub.shape[1]:
                                value = normalize_text(sub.iat[ii, jj])

                                if value and value.lower() != cell:
                                    values.append(value)

                        if values:
                            meta[key] = " | ".join(values[:3])
                            break

                if meta[key]:
                    break

    return meta

# ============================================================
# RESULT TABLE EXTRACTION
# ============================================================

@dataclass
class CalibrationResultTable:
    """Extracted result table and where it was found."""
    points: pd.DataFrame
    sheet_name: str
    header_row: int


def row_contains_keywords(row: pd.Series, keyword_groups: List[List[str]]) -> bool:
    """Check whether a row contains all required concept groups."""
    row_text = " ".join(row.astype(str).str.lower().fillna("").tolist())

    for group in keyword_groups:
        if not any(keyword in row_text for keyword in group):
            return False

    return True

def find_result_header_row(df_raw: pd.DataFrame) -> Optional[int]:
    """Find a header row containing reference, reading, and error concepts."""
    groups = [
        HEADER_KEYWORDS["reference"],
        HEADER_KEYWORDS["reading"],
        HEADER_KEYWORDS["error"],
    ]

    for i in range(min(len(df_raw), 350)):
        if row_contains_keywords(df_raw.iloc[i], groups):
            return i

    return None

def infer_result_columns(header_row: pd.Series) -> Dict[str, int]:
    """
    Map logical result fields to Excel column indexes.
    Example: "Deviation" may be mapped to error.
    """
    positions = {}

    for idx, value in enumerate(header_row.astype(str).tolist()):
        text = str(value).lower().strip()

        for logical_name, keywords in HEADER_KEYWORDS.items():
            if logical_name in positions:
                continue

            if any(keyword in text for keyword in keywords):
                positions[logical_name] = idx

    return positions

# Converts Excel cell content into a numeric value before table extraction.
def to_number(x) -> float:
    """Convert Excel value to number; supports comma decimal values."""
    return pd.to_numeric(str(x).replace(",", "."), errors="coerce")

def extract_calibration_result_table(path: str, nominal_tolerance: float) -> CalibrationResultTable:
    """
    Extract calibration result table from workbook.
    Required columns: reference, reading, error.
    """
    xls = get_excel_file(path)

    # Try likely result-sheet names first; otherwise scan every sheet.
    preferred = pick_sheet(xls, RESULT_SHEET_CANDIDATES)
    sheets_to_scan = [preferred] if preferred else xls.sheet_names

    last_error = None

    for sheet in sheets_to_scan:
        try:
            df_raw = read_excel_any(path, sheet_name=sheet, header=None)
        except Exception as e:
            last_error = e
            continue

        # Detect the row containing reference, reading, and error headers.
        header_row = find_result_header_row(df_raw)

        if header_row is None:
            continue

        column_positions = infer_result_columns(df_raw.iloc[header_row])

        required = {"reference", "reading", "error"}

        if not required.issubset(set(column_positions.keys())):
            continue

        numeric_start = header_row + 1

        # Skip unit rows or text rows until numeric result data starts.
        while numeric_start < len(df_raw):
            ref_candidate = to_number(df_raw.iat[numeric_start, column_positions["reference"]])

            if not pd.isna(ref_candidate):
                break

            numeric_start += 1

        rows = []

        # Read each calibration point until the result table ends.
        for i in range(numeric_start, len(df_raw)):
            reference = to_number(df_raw.iat[i, column_positions["reference"]])
            reading = to_number(df_raw.iat[i, column_positions["reading"]])
            error = to_number(df_raw.iat[i, column_positions["error"]])

            if pd.isna(reference):
                if len(rows) >= 2:
                    break
                continue

            point = np.nan
            uncertainty = np.nan

            if "point" in column_positions:
                point = to_number(df_raw.iat[i, column_positions["point"]])

            if "uncertainty" in column_positions:
                uncertainty = to_number(df_raw.iat[i, column_positions["uncertainty"]])

            rows.append({
                "point": point,
                "nominal_value": snap_nominal_value(reference, nominal_tolerance),
                "reference_value": reference,
                "reading_value": reading,
                "error_value": error,
                "uncertainty_value": uncertainty,
            })

        out = pd.DataFrame(rows)

        if out.empty or len(out) < 2:
            continue

        # Return both extracted data and its location for traceability.
        return CalibrationResultTable(
            points=out.reset_index(drop=True),
            sheet_name=sheet,
            header_row=header_row
        )

    if last_error is not None:
        raise RuntimeError(f"Result table could not be read: {path}\nLast error: {last_error}")

    raise RuntimeError(f"No calibration result table found in: {path}")

# ============================================================
# HISTORY BUILDING
# ============================================================

def build_calibration_history(
    candidate_df: pd.DataFrame,
    identifiers: List[str],
    nominal_tolerance: float,
    read_metadata: bool = True
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Build two outputs:
    raw result history and certificate metadata history.
    """
    result_rows = []
    metadata_rows = []

    total = len(candidate_df)

    # Each row in candidate_df is one matched certificate file.
    for idx, row in candidate_df.iterrows():
        path = row["path"]
        print(f"[{idx + 1}/{total}] Reading certificate: {path}")

        # First, try to extract the actual calibration result table.
        # If the file cannot be parsed, skip it and continue with the next one.
        try:
            result_table = extract_calibration_result_table(path, nominal_tolerance)
        except Exception as e:
            print(f"  [WARN] Result table skipped: {e}")
            continue

        date = None

        # Try to read the certificate/calibration date from inside the workbook.
        try:
            date = extract_date_from_workbook(path)
        except Exception:
            date = None

        # If no internal date is found, use file modification time as fallback.
        if date is None:
            date = file_mtime_date(path)

        metadata = {"source_file": path}

        # Metadata extraction is optional because some certificates may have unusual layouts.
        if read_metadata:
            try:
                metadata = extract_metadata_from_workbook(path)
            except Exception:
                metadata = {"source_file": path}

        # Add common tracking fields to the metadata record.
        metadata.update({
            "date": date,
            "date_str": date.strftime("%d-%m-%Y"),
            "result_sheet": result_table.sheet_name,
            "header_row": result_table.header_row,
            "matched_identifiers": "+".join(identifiers),
        })

        metadata_rows.append(metadata)

        # Convert the extracted result table into normalized long format.
        # One calibration point becomes one row, which makes pivoting/plotting easier later.
        for _, r in result_table.points.iterrows():
            result_rows.append({
                "file": path,
                "file_name": os.path.basename(path),
                "date": date,
                "date_str": date.strftime("%d-%m-%Y"),
                "matched_identifiers": "+".join(identifiers),
                "result_sheet": result_table.sheet_name,
                "point": r.get("point", np.nan),
                "nominal_value": r["nominal_value"],
                "reference_value": r["reference_value"],
                "reading_value": r["reading_value"],
                "error_value": r["error_value"],
                "uncertainty_value": r.get("uncertainty_value", np.nan),
                "certificate_no": metadata.get("certificate_no", ""),
                "job_no": metadata.get("job_no", ""),
                "instrument": metadata.get("instrument", ""),
                "model": metadata.get("model", ""),
                "manufacturer": metadata.get("manufacturer", ""),
                "serial": metadata.get("serial", ""),
                "client": metadata.get("client", ""),
            })

    # Stop clearly if none of the matched files produced usable calibration data.
    if not result_rows:
        raise RuntimeError("No calibration result data could be extracted from candidate files.")

    # Main output table: all calibration points from all certificates.
    raw_df = pd.DataFrame(result_rows)
    raw_df = raw_df.sort_values(["nominal_value", "date", "file_name"]).reset_index(drop=True)

    # Secondary output table: one metadata record per certificate.
    metadata_df = pd.DataFrame(metadata_rows)
    metadata_df = metadata_df.drop_duplicates(subset=["source_file"]).reset_index(drop=True)

    return raw_df, metadata_df

# ============================================================
# ANALYSIS TABLES
# ============================================================

# Builds higher-level analysis and statistical summary tables from raw calibration history.
def build_analysis_tables(raw_df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """Create pivot tables and summary tables."""
    raw_df = raw_df.copy()

    # Pivot table: error values by date and nominal calibration point.
    pivot_error = raw_df.pivot_table(
        index="date_str",
        columns="nominal_value",
        values="error_value",
        aggfunc="first"
    ).sort_index()

    # Pivot table: uncertainty values by date and nominal point.
    pivot_uncertainty = raw_df.pivot_table(
        index="date_str",
        columns="nominal_value",
        values="uncertainty_value",
        aggfunc="first"
    ).sort_index()

    # Pivot table: original reference values used during calibration.
    pivot_reference = raw_df.pivot_table(
        index="date_str",
        columns="nominal_value",
        values="reference_value",
        aggfunc="first"
    ).sort_index()

    # Statistical summary grouped by nominal calibration value.
    summary_by_nominal = raw_df.groupby("nominal_value").agg(
        number_of_records=("error_value", "count"),
        mean_error=("error_value", "mean"),
        std_error=("error_value", "std"),
        min_error=("error_value", "min"),
        max_error=("error_value", "max"),
        mean_uncertainty=("uncertainty_value", "mean"),
        min_uncertainty=("uncertainty_value", "min"),
        max_uncertainty=("uncertainty_value", "max"),
    ).reset_index().sort_values("nominal_value")

    # Certificate-level performance summary for quick comparison between calibrations.
    summary_by_certificate = raw_df.groupby(["date_str", "file_name"]).agg(
        number_of_points=("error_value", "count"),
        max_abs_error=("error_value", lambda x: np.nanmax(np.abs(x))),
        mean_abs_error=("error_value", lambda x: np.nanmean(np.abs(x))),
        mean_uncertainty=("uncertainty_value", "mean"),
    ).reset_index()

    return {
        "pivot_error": pivot_error,
        "pivot_uncertainty": pivot_uncertainty,
        "pivot_reference": pivot_reference,
        "summary_by_nominal": summary_by_nominal,
        "summary_by_certificate": summary_by_certificate,
    }

# ============================================================
# PLOTTING
# ============================================================

def plot_error_by_nominal(raw_df: pd.DataFrame, out_dir: str):
    """Create one error trend plot for each nominal value."""
    safe_mkdir(out_dir)

    nominal_values = sorted(raw_df["nominal_value"].dropna().unique())

    for nominal in nominal_values:
        df = raw_df[raw_df["nominal_value"] == nominal].sort_values("date")

        if df.empty:
            continue

        plt.figure(figsize=(10, 5))
        plt.plot(df["date"], df["error_value"], marker="o")

        for _, r in df.iterrows():
            plt.annotate(
                r["date_str"],
                (r["date"], r["error_value"]),
                textcoords="offset points",
                xytext=(0, 6),
                ha="center",
                fontsize=8
            )

        plt.title(f"Error trend | Nominal value = {nominal}")
        plt.xlabel("Calibration date")
        plt.ylabel("Error")
        plt.grid(True, alpha=0.3)
        plt.xticks(rotation=45)
        plt.tight_layout()

        safe_nominal = str(nominal).replace(".", "p").replace("-", "minus")
        plt.savefig(os.path.join(out_dir, f"error_trend_nominal_{safe_nominal}.png"), dpi=220)
        plt.close()


def plot_all_nominal_errors(raw_df: pd.DataFrame, out_dir: str):
    """Create a combined trend plot for all nominal values."""
    safe_mkdir(out_dir)

    plt.figure(figsize=(12, 6))

    nominal_values = sorted(raw_df["nominal_value"].dropna().unique())

    for nominal in nominal_values:
        df = raw_df[raw_df["nominal_value"] == nominal].sort_values("date")

        if df.empty:
            continue

        plt.plot(df["date"], df["error_value"], marker="o", label=str(nominal))

    plt.title("All nominal points | Error vs calibration date")
    plt.xlabel("Calibration date")
    plt.ylabel("Error")
    plt.grid(True, alpha=0.3)
    plt.legend(ncol=2, fontsize=8)
    plt.xticks(rotation=45)
    plt.tight_layout()

    plt.savefig(os.path.join(out_dir, "all_nominal_error_trends.png"), dpi=220)
    plt.close()

def plot_error_heatmap(pivot_error: pd.DataFrame, out_dir: str):
    """Create heatmap of error values by date and nominal value."""
    safe_mkdir(out_dir)

    if pivot_error.empty:
        return

    data = pivot_error.values.astype(float)

    plt.figure(figsize=(max(8, 0.8 * pivot_error.shape[1]), max(4, 0.5 * pivot_error.shape[0])))

    im = plt.imshow(data, aspect="auto")
    plt.colorbar(im, label="Error")

    plt.title("Error heatmap by calibration date and nominal value")
    plt.xlabel("Nominal value")
    plt.ylabel("Calibration date")

    plt.xticks(range(len(pivot_error.columns)), [str(c) for c in pivot_error.columns], rotation=45)
    plt.yticks(range(len(pivot_error.index)), list(pivot_error.index))

    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "error_heatmap.png"), dpi=220)
    plt.close()

def plot_max_abs_error_by_certificate(summary_by_certificate: pd.DataFrame, out_dir: str):
    """Create certificate-level maximum absolute error plot."""
    safe_mkdir(out_dir)

    if summary_by_certificate.empty:
        return

    df = summary_by_certificate.copy()
    labels = df["date_str"] + "\n" + df["file_name"].astype(str).str[:25]

    plt.figure(figsize=(max(10, 0.8 * len(df)), 5))
    plt.bar(range(len(df)), df["max_abs_error"])
    plt.xticks(range(len(df)), labels, rotation=45, ha="right")
    plt.ylabel("Maximum absolute error")
    plt.title("Maximum absolute error by certificate")
    plt.tight_layout()

    plt.savefig(os.path.join(out_dir, "max_abs_error_by_certificate.png"), dpi=220)
    plt.close()

# ============================================================
# EXPORT
# ============================================================

# Exports all generated outputs into CSV files and one combined Excel workbook.
def export_outputs(
    output_dir: str,
    raw_df: pd.DataFrame,
    metadata_df: pd.DataFrame,
    analysis_tables: Dict[str, pd.DataFrame]
):
    """Export CSV files and a combined Excel report."""
    safe_mkdir(output_dir)

    # Main detailed dataset: one row per calibration point.
    raw_df.to_csv(
        os.path.join(output_dir, "raw_calibration_history.csv"),
        index=False,
        encoding="utf-8-sig"
    )

    # Certificate-level metadata extracted from each file.
    metadata_df.to_csv(
        os.path.join(output_dir, "certificate_metadata.csv"),
        index=False,
        encoding="utf-8-sig"
    )

    # Export each pivot/summary table as an individual CSV file.
    for name, table in analysis_tables.items():
        table.to_csv(
            os.path.join(output_dir, f"{name}.csv"),
            encoding="utf-8-sig",
            index=not name.startswith("summary")
        )

    xlsx_path = os.path.join(output_dir, "calibration_automation_report.xlsx")

    try:
        # Single Excel workbook is useful for quick manual review.
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            raw_df.to_excel(writer, sheet_name="raw_history", index=False)
            metadata_df.to_excel(writer, sheet_name="metadata", index=False)

            for name, table in analysis_tables.items():
                # Excel sheet names cannot exceed 31 characters.
                sheet_name = name[:31]
                table.to_excel(writer, sheet_name=sheet_name, index=not name.startswith("summary"))

    except Exception as e:
        print(f"[WARN] Excel report could not be written: {e}")

# ============================================================
# MAIN PROGRAM
# ============================================================

def main():
    """Run the full calibration automation pipeline."""
    print("\n=== Generic Calibration & Certification Process Automation Tool ===\n")

    current_file = ask_path(
        "1) Enter current workflow file, certificate file, or related Excel file: ",
        must_exist=True
    )

    archive_root = ask_path(
        "2) Enter calibration certificate archive root folder: ",
        must_exist=True
    )

    output_dir = ask_path(
        f"3) Enter output folder. Leave empty for ./{DEFAULT_OUTPUT_DIRNAME}: ",
        must_exist=False,
        allow_empty=True
    )

    if not output_dir:
        output_dir = os.path.join(os.getcwd(), DEFAULT_OUTPUT_DIRNAME)

    nominal_tolerance = ask_float(
        "4) Nominal value snapping tolerance",
        default=0.05
    )

    require_all = ask_yes_no(
        "5) Should all identifiers be required in candidate file names?",
        default=True
    )

    auto_ids = ask_yes_no(
        "6) Try to extract identifiers automatically from the current file?",
        default=True
    )

    identifiers = []

    if auto_ids:
        identifiers = extract_identifiers_auto(current_file)
        print(f"   Automatically detected identifiers: {identifiers}")

    # Manual override is useful when automatic extraction captures extra numbers.
    if (not identifiers) or ask_yes_no("7) Manually override identifiers?", default=False):
        s = input("   Enter identifiers. Example: 23014357+23014081 or 21040250,21033263: ").strip()
        identifiers = list(dict.fromkeys(re.findall(r"\d{5,}", s)))

    if not identifiers:
        raise RuntimeError("No identifiers were found or entered.")

    safe_mkdir(output_dir)
    cache_csv_path = os.path.join(output_dir, DEFAULT_INDEX_FILENAME)

    force_rebuild = ask_yes_no(
        "8) Rebuild archive index from scratch?",
        default=False
    )

    print("\n--- Preparing archive index ---")
    t0 = time.time()

    archive_index = build_archive_index(
        archive_root,
        cache_csv_path,
        force_rebuild=force_rebuild
    )

    print(f"[OK] Archive index ready. Total Excel files: {len(archive_index)}")
    print(f"Elapsed time: {time.time() - t0:.1f} seconds")

    print("\n--- Searching matching certificate files ---")

    candidates = search_archive_index(
        archive_index,
        identifiers,
        require_all=require_all
    )

    candidates = candidates[
        ~candidates["path"].apply(looks_like_non_certificate_file)
    ].reset_index(drop=True)

    if candidates.empty:
        raise RuntimeError("No matching certificate files found for the given identifiers.")

    candidates_path = os.path.join(output_dir, "matched_certificate_candidates.csv")
    candidates.to_csv(candidates_path, index=False, encoding="utf-8-sig")

    print(f"[OK] Candidate certificate count: {len(candidates)}")
    print(f"Candidate list saved to: {candidates_path}")

    print("\n--- Extracting calibration result data ---")

    raw_df, metadata_df = build_calibration_history(
        candidate_df=candidates,
        identifiers=identifiers,
        nominal_tolerance=nominal_tolerance,
        read_metadata=True
    )

    print("\n--- Building analysis tables ---")

    analysis_tables = build_analysis_tables(raw_df)

    print("\n--- Creating plots ---")

    plots_dir = os.path.join(output_dir, "plots")

    plot_error_by_nominal(
        raw_df,
        os.path.join(plots_dir, "per_nominal_value")
    )

    plot_all_nominal_errors(
        raw_df,
        plots_dir
    )

    plot_error_heatmap(
        analysis_tables["pivot_error"],
        plots_dir
    )

    plot_max_abs_error_by_certificate(
        analysis_tables["summary_by_certificate"],
        plots_dir
    )

    print("\n--- Exporting reports ---")

    export_outputs(
        output_dir=output_dir,
        raw_df=raw_df,
        metadata_df=metadata_df,
        analysis_tables=analysis_tables
    )

    print("\n=== PROCESS COMPLETED ===")
    print("Matched identifiers:", "+".join(identifiers))
    print("Candidate certificate count:", len(candidates))
    print("Successfully parsed certificate count:", raw_df["file"].nunique())
    print("Nominal point count:", raw_df["nominal_value"].nunique())
    print("Total extracted rows:", len(raw_df))
    print("Output folder:", output_dir)

if __name__ == "__main__":
    main() 
