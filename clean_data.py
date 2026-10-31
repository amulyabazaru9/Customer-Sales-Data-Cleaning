"""
clean_data.py
-------------
End-to-end data cleaning and analysis pipeline for messy customer sales data.

Pipeline:
    1. Load the raw CSV (every column read as text so nothing is silently lost)
    2. Inspect it (shape, missing values) and log the findings
    3. Standardise text: whitespace, capitalisation, IDs, locations, products
    4. Clean emails and phone numbers
    5. Normalise mixed date formats into real dates
    6. Validate quantities and prices
    7. Remove exact and near-duplicate records
    8. Separate invalid records (kept in a 'Rejected Records' sheet, never
       silently deleted) and calculate total sales
    9. Export a formatted Excel workbook, a clean CSV and three charts

Usage (from the project root):
    python clean_data.py
    python clean_data.py --input data/raw_sales_data.csv --output-dir output
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import matplotlib
import numpy as np

matplotlib.use("Agg")  # render charts to files without needing a display window
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402
from matplotlib.ticker import FuncFormatter  # noqa: E402
from openpyxl import Workbook  # noqa: E402
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

# ---------------------------------------------------------------------------
# Configuration - relative paths only, so the project runs on any computer
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = BASE_DIR / "data" / "raw_sales_data.csv"
DEFAULT_OUTPUT_DIR = BASE_DIR / "output"

EXCEL_NAME = "sales_report.xlsx"
CLEAN_CSV_NAME = "cleaned_sales_data.csv"
LOG_NAME = "cleaning_log.txt"
CHART_DIR_NAME = "charts"

REQUIRED_COLUMNS = [
    "order_id", "order_date", "customer_name", "customer_email", "phone",
    "product", "category", "quantity", "unit_price", "city", "state",
]

# Business rules for validation (easy to adjust per client)
MAX_QUANTITY = 1_000        # more than this per line is treated as a typo
MAX_UNIT_PRICE = 10_000.00  # the most expensive catalogue item is far below this

# Every date format we expect to see. Each one is structurally unambiguous
# (different separators / field order), so trying them in turn is safe.
DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%m-%d-%Y", "%B %d, %Y", "%Y/%m/%d", "%d-%b-%Y"]

EMAIL_PATTERN = re.compile(r"^[a-z0-9._%+-]+@[a-z0-9-]+(\.[a-z0-9-]+)*\.[a-z]{2,}$")

UNKNOWN_CUSTOMER = "Unknown Customer"
UNKNOWN_LOCATION = "Unknown"

# Chart styling (single-series charts -> one accent colour, quiet axes)
CHART_COLOR = "#2a78d6"
CHART_SURFACE = "#fcfcfb"
INK_PRIMARY = "#0b0b0b"
INK_SECONDARY = "#52514e"
GRID_COLOR = "#e1e0d9"

# Excel styling
HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14, color="1F3A5F")
SECTION_FONT = Font(bold=True, size=12, color="1F3A5F")
NOTE_FILL = PatternFill("solid", fgColor="FFF4CE")
THIN_BORDER = Border(bottom=Side(style="thin", color="D9D9D9"))
CURRENCY_FMT = '"$"#,##0.00'
INTEGER_FMT = "#,##0"
PERCENT_FMT = "0.0%"
DATE_FMT = "yyyy-mm-dd"

logger = logging.getLogger("sales_cleaner")


class PipelineError(Exception):
    """Raised for problems the user can fix (missing file, wrong columns...)."""


def changed(new: pd.Series, old: pd.Series) -> pd.Series:
    """True where a value changed. Missing -> missing does not count as a change.
    Works the same on pandas 2.x and 3.x (which handle missing text differently)."""
    return new.fillna("\x00").astype(str) != old.fillna("\x00").astype(str)


def text_mask(result: pd.Series) -> pd.Series:
    """Turn a string-method result containing missing values into a clean True/False mask."""
    return result.eq(True).astype(bool)


@dataclass
class CleaningContext:
    """
    Collects everything that happens during cleaning so it can be reported:
      - log_entries: one row per cleaning action (goes to the Excel report)
      - notes:   non-fatal issues per row (row kept, issue documented)
      - rejects: fatal issues per row (row moved to 'Rejected Records')
    Rows are identified by their original CSV line number ('source_row').
    """
    log_entries: list[dict] = field(default_factory=list)
    notes: dict[int, list[str]] = field(default_factory=dict)
    rejects: dict[int, list[str]] = field(default_factory=dict)

    def log(self, step: str, issue: str, rows: int, action: str) -> None:
        self.log_entries.append(
            {"Step": step, "Issue": issue, "Rows Affected": int(rows), "Action Taken": action}
        )
        logger.info("[%s] %s: %d row(s) -> %s", step, issue, rows, action)

    def add_note(self, rows, text: str) -> None:
        for r in rows:
            self.notes.setdefault(int(r), []).append(text)

    def add_reject(self, rows, text: str) -> None:
        for r in rows:
            self.rejects.setdefault(int(r), []).append(text)


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
def setup_logging(log_path: Path) -> logging.FileHandler:
    """Log to the console and to a text file (overwritten on each run)."""
    logger.setLevel(logging.INFO)
    logger.propagate = False
    # Remove handlers from a previous run so messages are not duplicated
    for handler in list(logger.handlers):
        handler.close()
        logger.removeHandler(handler)

    fmt = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", "%Y-%m-%d %H:%M:%S")
    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(fmt)
    file_handler = logging.FileHandler(log_path, mode="w", encoding="utf-8")
    file_handler.setFormatter(fmt)
    logger.addHandler(console)
    logger.addHandler(file_handler)
    return file_handler


# ---------------------------------------------------------------------------
# 1-2. Load and inspect
# ---------------------------------------------------------------------------
def load_data(path: Path) -> pd.DataFrame:
    """Load the CSV as text. Reading as text keeps values like '$1,299.00' intact."""
    if not path.exists():
        raise PipelineError(
            f"Input file not found: {path}\nRun 'python generate_dataset.py' first."
        )
    try:
        df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8")
    except pd.errors.EmptyDataError as exc:
        raise PipelineError(f"Input file is empty: {path}") from exc
    except UnicodeDecodeError:
        # Files saved from Excel on Windows are often cp1252, not UTF-8
        df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="cp1252")

    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise PipelineError(f"Input file is missing required columns: {missing}")
    if df.empty:
        raise PipelineError("Input file has a header but no data rows.")

    # Line number in the CSV (header = line 1) so every row can be traced back
    df.insert(0, "source_row", range(2, len(df) + 2))
    df = df.set_index("source_row", drop=False)
    logger.info("Loaded %d rows x %d columns from %s", len(df), len(REQUIRED_COLUMNS), path.name)
    return df


def inspect_data(df: pd.DataFrame) -> pd.DataFrame:
    """Return missing-value counts per column (blank or whitespace = missing)."""
    data = df[REQUIRED_COLUMNS]
    blank = data.apply(lambda col: col.str.strip().eq(""))
    summary = pd.DataFrame({
        "Column": REQUIRED_COLUMNS,
        "Missing (Raw)": blank.sum().values,
        "Missing % (Raw)": (blank.sum() / len(data)).values,
    })
    logger.info("Missing values in raw data:\n%s",
                summary[summary["Missing (Raw)"] > 0].to_string(index=False))
    return summary


# ---------------------------------------------------------------------------
# 3. Text standardisation
# ---------------------------------------------------------------------------
def strip_whitespace(df: pd.DataFrame, ctx: CleaningContext) -> pd.DataFrame:
    """Trim leading/trailing spaces, collapse repeated spaces, blank -> missing."""
    df = df.copy()
    any_change = pd.Series(False, index=df.index)
    for col in REQUIRED_COLUMNS:
        cleaned = df[col].str.replace(r"\s+", " ", regex=True).str.strip()
        any_change |= changed(cleaned, df[col])
        # Empty text becomes a real missing value (NaN) so pandas can detect it
        df[col] = cleaned.replace("", np.nan).astype(object)
    ctx.log("Whitespace", "Leading/trailing or repeated spaces", any_change.sum(),
            "Trimmed and collapsed to single spaces")
    return df


def canonical_spelling(series: pd.Series) -> pd.Series:
    """
    Map every spelling variant to the most common one, ignoring case.
    e.g. 'usb-c hub', 'USB-C HUB', 'USB-C Hub' -> 'USB-C Hub' (most frequent).
    This keeps brand-style capitalisation (USB-C, 4K) that .title() would break.
    """
    key = series.str.lower()
    most_common = (
        pd.DataFrame({"key": key, "value": series})
        .dropna()
        .groupby("key")["value"]
        .agg(lambda v: v.value_counts().index[0])
    )
    return key.map(most_common)


def standardize_text(df: pd.DataFrame, ctx: CleaningContext) -> pd.DataFrame:
    """Fix capitalisation of IDs, names, products, categories and locations."""
    df = df.copy()

    new_ids = df["order_id"].str.upper()
    ctx.log("Standardise", "Order IDs in lowercase", changed(new_ids, df["order_id"]).sum(),
            "Converted to uppercase (e.g. ORD-10001)")
    df["order_id"] = new_ids

    new_names = df["customer_name"].str.title()
    ctx.log("Standardise", "Customer names with incorrect capitalisation",
            changed(new_names, df["customer_name"]).sum(), "Converted to Title Case")
    df["customer_name"] = new_names

    for col in ["product", "category", "city"]:
        new_values = canonical_spelling(df[col])
        ctx.log("Standardise", f"Inconsistent capitalisation in '{col}'",
                changed(new_values, df[col]).sum(),
                "Mapped to the most common spelling")
        df[col] = new_values

    new_state = df["state"].str.upper()
    ctx.log("Standardise", "State codes in lowercase", changed(new_state, df["state"]).sum(),
            "Converted to uppercase")
    df["state"] = new_state
    return df


# ---------------------------------------------------------------------------
# 4. Emails and phone numbers
# ---------------------------------------------------------------------------
def clean_emails(df: pd.DataFrame, ctx: CleaningContext) -> pd.DataFrame:
    """Lowercase emails, remove inner spaces and blank out invalid addresses."""
    df = df.copy()
    original = df["customer_email"]
    cleaned = original.str.lower().str.replace(" ", "", regex=False)
    ctx.log("Email", "Uppercase letters or spaces in email",
            changed(cleaned, original).sum(), "Lowercased and removed spaces")

    missing = cleaned.isna()
    invalid = ~missing & ~text_mask(cleaned.fillna("").str.match(EMAIL_PATTERN))
    ctx.log("Email", "Missing email", missing.sum(), "Kept row, flagged in Data Notes")
    ctx.log("Email", "Invalid email format (e.g. missing @ or domain)", invalid.sum(),
            "Removed invalid value, kept row, flagged in Data Notes")
    ctx.add_note(df.index[missing], "Missing email")
    for row, bad in cleaned[invalid].items():
        ctx.add_note([row], f"Invalid email removed ('{bad}')")

    df["customer_email"] = cleaned.mask(invalid)
    return df


def clean_phones(df: pd.DataFrame, ctx: CleaningContext) -> pd.DataFrame:
    """Normalise US phone numbers to (555) 123-4567; blank out invalid ones."""
    df = df.copy()
    digits = df["phone"].str.replace(r"\D", "", regex=True)
    # Drop the +1 country code if present
    digits = digits.where(~(text_mask(digits.str.len().eq(11)) & text_mask(digits.str.startswith("1"))),
                           digits.str[1:])
    valid = text_mask(digits.str.len().eq(10))
    formatted = "(" + digits.str[:3] + ") " + digits.str[3:6] + "-" + digits.str[6:]

    missing = df["phone"].isna()
    invalid = ~missing & ~valid
    ctx.log("Phone", "Phone numbers in mixed formats", (valid & changed(formatted, df["phone"])).sum(),
            "Normalised to (XXX) XXX-XXXX")
    ctx.log("Phone", "Missing phone", missing.sum(), "Kept row, flagged in Data Notes")
    ctx.log("Phone", "Invalid phone (wrong number of digits)", invalid.sum(),
            "Removed invalid value, kept row, flagged in Data Notes")
    ctx.add_note(df.index[missing], "Missing phone")
    ctx.add_note(df.index[invalid], "Invalid phone removed")

    df["phone"] = formatted.where(valid)
    return df


# ---------------------------------------------------------------------------
# 5. Dates
# ---------------------------------------------------------------------------
def parse_dates(df: pd.DataFrame, ctx: CleaningContext) -> pd.DataFrame:
    """Try each known format and keep the first that parses. Unparseable -> rejected."""
    df = df.copy()
    raw = df["order_date"]
    parsed = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
    for fmt in DATE_FORMATS:
        attempt = pd.to_datetime(raw, format=fmt, errors="coerce")
        parsed = parsed.fillna(attempt)

    non_iso = parsed.notna() & changed(parsed.dt.strftime("%Y-%m-%d"), raw)
    missing = raw.isna()
    invalid = ~missing & parsed.isna()
    ctx.log("Dates", "Dates not in YYYY-MM-DD format", non_iso.sum(),
            f"Parsed {len(DATE_FORMATS)} known formats into one date format")
    ctx.log("Dates", "Missing order date", missing.sum(), "Rejected (date is required)")
    ctx.log("Dates", "Unparseable/impossible date (e.g. 2025-13-45)", invalid.sum(),
            "Rejected (date is required)")
    ctx.add_reject(df.index[missing], "Missing order date")
    for row, bad in raw[invalid].items():
        ctx.add_reject([row], f"Invalid order date ('{bad}')")

    df["order_date"] = parsed
    return df


# ---------------------------------------------------------------------------
# 6. Quantities and prices
# ---------------------------------------------------------------------------
def validate_quantities(df: pd.DataFrame, ctx: CleaningContext) -> pd.DataFrame:
    """Quantities must be whole numbers between 1 and MAX_QUANTITY."""
    df = df.copy()
    raw = df["quantity"]
    qty = pd.to_numeric(raw, errors="coerce")

    missing = raw.isna()
    non_numeric = ~missing & qty.isna()
    out_of_range = qty.notna() & ((qty < 1) | (qty > MAX_QUANTITY))
    fractional = qty.notna() & ~out_of_range & (qty % 1 != 0)
    decimal_text = qty.notna() & (qty % 1 == 0) & text_mask(raw.str.contains(".", regex=False))

    ctx.log("Quantity", "Whole-number quantity written as decimal (e.g. '2.0')",
            decimal_text.sum(), "Converted to integer")
    for mask, label in [
        (missing, "Missing quantity"),
        (non_numeric, "Non-numeric quantity"),
        (out_of_range, f"Quantity zero, negative or above {MAX_QUANTITY}"),
        (fractional, "Fractional quantity"),
    ]:
        ctx.log("Quantity", label, mask.sum(), "Rejected")
        for row in df.index[mask]:
            ctx.add_reject([row], f"{label} ('{raw[row]}')" if pd.notna(raw[row]) else label)

    valid = qty.notna() & ~out_of_range & ~fractional
    df["quantity"] = qty.where(valid).astype("Int64")  # nullable integer
    return df


def validate_prices(df: pd.DataFrame, ctx: CleaningContext) -> pd.DataFrame:
    """Prices must be numeric, above zero and below MAX_UNIT_PRICE."""
    df = df.copy()
    raw = df["unit_price"]
    stripped = raw.str.replace(r"[$,]", "", regex=True)
    price = pd.to_numeric(stripped, errors="coerce")

    had_symbols = price.notna() & changed(stripped, raw)
    missing = raw.isna()
    non_numeric = ~missing & price.isna()
    non_positive = price.notna() & (price <= 0)
    outlier = price.notna() & (price > MAX_UNIT_PRICE)

    ctx.log("Price", "Price stored as text with '$' or ','", had_symbols.sum(),
            "Removed symbols and converted to number")
    for mask, label in [
        (missing, "Missing unit price"),
        (non_numeric, "Non-numeric unit price"),
        (non_positive, "Zero or negative unit price"),
        (outlier, f"Unit price above ${MAX_UNIT_PRICE:,.0f} (likely typo)"),
    ]:
        ctx.log("Price", label, mask.sum(), "Rejected")
        for row in df.index[mask]:
            ctx.add_reject([row], f"{label} ('{raw[row]}')" if pd.notna(raw[row]) else label)

    df["unit_price"] = price.where(~(missing | non_numeric | non_positive | outlier)).round(2)
    return df


# ---------------------------------------------------------------------------
# 7. Missing required / optional fields
# ---------------------------------------------------------------------------
def handle_missing_fields(df: pd.DataFrame, ctx: CleaningContext) -> pd.DataFrame:
    """Reject rows without an order ID or product; fill optional fields with labels."""
    df = df.copy()
    for col, label in [("order_id", "Missing order ID"), ("product", "Missing product")]:
        mask = df[col].isna()
        ctx.log("Missing data", label, mask.sum(), "Rejected (field is required)")
        ctx.add_reject(df.index[mask], label)

    no_name = df["customer_name"].isna()
    ctx.log("Missing data", "Missing customer name", no_name.sum(),
            f"Filled with '{UNKNOWN_CUSTOMER}', flagged in Data Notes")
    ctx.add_note(df.index[no_name], "Missing customer name")
    df["customer_name"] = df["customer_name"].fillna(UNKNOWN_CUSTOMER)

    no_city = df["city"].isna()
    ctx.log("Missing data", "Missing city", no_city.sum(),
            f"Filled with '{UNKNOWN_LOCATION}', flagged in Data Notes")
    ctx.add_note(df.index[no_city], "Missing city")
    df["city"] = df["city"].fillna(UNKNOWN_LOCATION)
    df["state"] = df["state"].fillna(UNKNOWN_LOCATION)
    return df


# ---------------------------------------------------------------------------
# 8. Duplicates
# ---------------------------------------------------------------------------
def remove_duplicates(df: pd.DataFrame, ctx: CleaningContext) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Remove duplicates AFTER standardisation, so near-duplicates such as
    'JOHN SMITH ' vs 'John Smith' are caught. The first occurrence is kept.
    Then flag any order ID that still appears more than once with different data.
    """
    is_dup = df.duplicated(subset=REQUIRED_COLUMNS, keep="first")
    duplicates = df[is_dup].copy()
    df = df[~is_dup].copy()
    ctx.log("Duplicates", "Exact or near-duplicate records", is_dup.sum(),
            "Removed (first occurrence kept)")

    conflict = df["order_id"].notna() & df.duplicated(subset=["order_id"], keep="first")
    ctx.log("Duplicates", "Same order ID with different details", conflict.sum(),
            "Rejected (needs manual review)")
    for row in df.index[conflict]:
        ctx.add_reject([row], f"Duplicate order ID with conflicting data ({df.at[row, 'order_id']})")

    # Rejection reasons from duplicate rows are no longer relevant
    for row in duplicates.index:
        ctx.rejects.pop(int(row), None)
        ctx.notes.pop(int(row), None)
    return df, duplicates


# ---------------------------------------------------------------------------
# 9. Split, totals, analysis tables
# ---------------------------------------------------------------------------
def split_valid_rejected(df: pd.DataFrame, raw: pd.DataFrame, ctx: CleaningContext):
    """Return (clean_df, rejected_df). Rejected rows show their ORIGINAL values."""
    rejected_mask = df.index.isin(list(ctx.rejects))
    clean = df[~rejected_mask].copy()

    rejected_rows = df.index[rejected_mask]
    rejected = raw.loc[rejected_rows, ["source_row"] + REQUIRED_COLUMNS].copy()
    rejected.insert(1, "rejection_reason", ["; ".join(ctx.rejects[int(r)]) for r in rejected_rows])

    clean["total_sales"] = (clean["quantity"].astype(float) * clean["unit_price"]).round(2)
    clean["data_notes"] = ["; ".join(ctx.notes.get(int(r), [])) for r in clean.index]
    clean = clean.sort_values(["order_date", "order_id"]).reset_index(drop=True)
    clean["order_date"] = clean["order_date"].dt.date
    clean["quantity"] = clean["quantity"].astype(int)

    columns = ["source_row", "order_id", "order_date", "customer_name", "customer_email",
               "phone", "product", "category", "quantity", "unit_price", "total_sales",
               "city", "state", "data_notes"]
    logger.info("Valid rows: %d | Rejected rows: %d", len(clean), len(rejected))
    return clean[columns], rejected.reset_index(drop=True)


def build_product_analysis(clean: pd.DataFrame) -> pd.DataFrame:
    prod = (
        clean.groupby(["product", "category"], as_index=False)
        .agg(orders=("order_id", "count"), units_sold=("quantity", "sum"),
             revenue=("total_sales", "sum"), avg_unit_price=("unit_price", "mean"))
        .sort_values("revenue", ascending=False)
        .reset_index(drop=True)
    )
    prod["revenue_share"] = prod["revenue"] / prod["revenue"].sum()
    prod.insert(0, "rank", range(1, len(prod) + 1))
    return prod


def build_category_analysis(clean: pd.DataFrame) -> pd.DataFrame:
    cat = (
        clean.groupby("category", as_index=False)
        .agg(orders=("order_id", "count"), units_sold=("quantity", "sum"),
             revenue=("total_sales", "sum"))
        .sort_values("revenue", ascending=False)
        .reset_index(drop=True)
    )
    cat["revenue_share"] = cat["revenue"] / cat["revenue"].sum()
    return cat


def build_monthly_sales(clean: pd.DataFrame) -> pd.DataFrame:
    month = pd.to_datetime(clean["order_date"]).dt.to_period("M")
    monthly = (
        clean.assign(month=month)
        .groupby("month", as_index=False)
        .agg(orders=("order_id", "count"), units_sold=("quantity", "sum"),
             revenue=("total_sales", "sum"))
        .sort_values("month")
        .reset_index(drop=True)
    )
    monthly["mom_change"] = monthly["revenue"].pct_change()
    monthly["month"] = monthly["month"].dt.strftime("%Y-%m")
    return monthly


def build_location_sales(clean: pd.DataFrame) -> pd.DataFrame:
    # Rows with an unknown city are grouped together, whatever their state
    state = clean["state"].where(clean["city"] != UNKNOWN_LOCATION, "-")
    loc = (
        clean.assign(state=state)
        .groupby(["city", "state"], as_index=False)
        .agg(orders=("order_id", "count"), revenue=("total_sales", "sum"))
        .sort_values("revenue", ascending=False)
        .reset_index(drop=True)
    )
    loc["revenue_share"] = loc["revenue"] / loc["revenue"].sum()
    return loc


def build_kpis(clean: pd.DataFrame, raw_rows: int) -> list[tuple[str, object, str]]:
    known = clean.loc[clean["customer_name"] != UNKNOWN_CUSTOMER, "customer_name"]
    return [
        ("Total Revenue", float(clean["total_sales"].sum()), CURRENCY_FMT),
        ("Valid Orders", int(len(clean)), INTEGER_FMT),
        ("Units Sold", int(clean["quantity"].sum()), INTEGER_FMT),
        ("Average Order Value", float(clean["total_sales"].mean()), CURRENCY_FMT),
        ("Unique Customer Names", int(known.nunique()), INTEGER_FMT),
        ("Date Range", f"{clean['order_date'].min()} to {clean['order_date'].max()}", "@"),
        ("Raw Rows Received", int(raw_rows), INTEGER_FMT),
    ]


# ---------------------------------------------------------------------------
# 10. Excel export (OpenPyXL)
# ---------------------------------------------------------------------------
def _pretty(name: str) -> str:
    """'unit_price' -> 'Unit Price'."""
    pretty = name.replace("_", " ").title().replace("Mom ", "MoM ")
    return re.sub(r"\bId\b", "ID", pretty)


def write_table(ws, df: pd.DataFrame, start_row: int, start_col: int = 1,
                title: str | None = None, formats: dict[str, str] | None = None) -> tuple[int, int]:
    """
    Write a DataFrame as a formatted table. Returns (first_data_row, last_data_row).
    `formats` maps column name -> Excel number format.
    """
    formats = formats or {}
    row = start_row
    if title:
        ws.cell(row=row, column=start_col, value=title).font = SECTION_FONT
        row += 1

    for j, col in enumerate(df.columns):
        cell = ws.cell(row=row, column=start_col + j, value=_pretty(col))
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    header_row = row

    for record in df.itertuples(index=False):
        row += 1
        for j, value in enumerate(record):
            if pd.isna(value):
                value = None
            elif hasattr(value, "item"):
                value = value.item()  # numpy scalar -> plain Python value
            cell = ws.cell(row=row, column=start_col + j, value=value)
            cell.border = THIN_BORDER
            fmt = formats.get(df.columns[j])
            if fmt:
                cell.number_format = fmt
    return header_row + 1, row


def autofit_columns(ws, min_width: int = 10, max_width: int = 60) -> None:
    """Approximate Excel's 'AutoFit' by measuring the longest value per column."""
    widths: dict[int, int] = {}
    for row in ws.iter_rows():
        for cell in row:
            # Skip empty cells and section titles (titles may overflow into empty cells)
            if cell.value is None or (cell.font and cell.font.sz and cell.font.sz >= 12):
                continue
            text = cell.value.strftime("%Y-%m-%d") if hasattr(cell.value, "strftime") else str(cell.value)
            if isinstance(cell.value, float):
                text = f"{cell.value:,.2f}$"
            widths[cell.column] = max(widths.get(cell.column, 0), len(text))
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = max(min_width, min(max_width, width + 3))


def add_data_bar(ws, col: int, first: int, last: int, color: str = "5B9BD5") -> None:
    if last >= first:
        letter = get_column_letter(col)
        ws.conditional_formatting.add(
            f"{letter}{first}:{letter}{last}",
            DataBarRule(start_type="min", end_type="max", color=color),
        )


def sheet_cleaned_data(wb: Workbook, clean: pd.DataFrame) -> None:
    ws = wb.active
    ws.title = "Cleaned Data"
    formats = {"order_date": DATE_FMT, "unit_price": CURRENCY_FMT,
               "total_sales": CURRENCY_FMT, "quantity": INTEGER_FMT}
    first, last = write_table(ws, clean, start_row=1, formats=formats)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    cols = list(clean.columns)
    add_data_bar(ws, cols.index("total_sales") + 1, first, last)
    # Highlight rows that were kept but have a documented issue
    notes_letter = get_column_letter(cols.index("data_notes") + 1)
    last_letter = get_column_letter(len(cols))
    ws.conditional_formatting.add(
        f"A{first}:{last_letter}{last}",
        FormulaRule(formula=[f'LEN(${notes_letter}{first})>0'], fill=NOTE_FILL),
    )
    autofit_columns(ws)
    ws.column_dimensions[notes_letter].width = 55


def sheet_quality_report(wb: Workbook, overview: list[tuple[str, object, str]],
                         missing: pd.DataFrame, reject_reasons: pd.DataFrame,
                         log_df: pd.DataFrame) -> None:
    ws = wb.create_sheet("Data Quality Report")
    ws["A1"] = "Data Quality Report"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Generated {datetime.now():%Y-%m-%d %H:%M}"
    ws["A2"].font = Font(italic=True, color="595959")

    ws["A4"] = "Dataset Overview"
    ws["A4"].font = SECTION_FONT
    for i, (label, value, fmt) in enumerate(overview, start=5):
        ws.cell(row=i, column=1, value=label).font = Font(bold=True)
        cell = ws.cell(row=i, column=2, value=value)
        cell.number_format = fmt
        cell.border = THIN_BORDER

    row = 5 + len(overview) + 1
    _, end = write_table(ws, missing, row, title="Missing Values by Column",
                         formats={"Missing % (Raw)": PERCENT_FMT})
    first = row + 2
    ws.conditional_formatting.add(
        f"B{first}:B{end}", CellIsRule(operator="greaterThan", formula=["0"], fill=NOTE_FILL))

    row = end + 2
    _, end = write_table(ws, reject_reasons, row, title="Rejection Reasons")

    row = end + 2
    ws.cell(row=row, column=1, value="Cleaning Log (every action taken)").font = SECTION_FONT
    ws.cell(row=row + 1, column=1, value=(
        "Note: steps before 'Duplicates' are counted on all raw rows, so a problem that also "
        "appears in a duplicate row is counted twice there. The overview above shows final counts."
    )).font = Font(italic=True, color="595959")
    first, end = write_table(ws, log_df, row + 2)
    ws.conditional_formatting.add(
        f"C{first}:C{end}", CellIsRule(operator="greaterThan", formula=["0"], fill=NOTE_FILL))
    autofit_columns(ws)
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 48


def sheet_sales_summary(wb: Workbook, kpis, monthly: pd.DataFrame, locations: pd.DataFrame) -> None:
    ws = wb.create_sheet("Sales Summary")
    ws["A1"] = "Sales Summary"
    ws["A1"].font = TITLE_FONT

    ws["A3"] = "Key Metrics"
    ws["A3"].font = SECTION_FONT
    for i, (label, value, fmt) in enumerate(kpis, start=4):
        ws.cell(row=i, column=1, value=label).font = Font(bold=True)
        cell = ws.cell(row=i, column=2, value=value)
        cell.number_format, cell.font = fmt, Font(size=12)
        cell.alignment = Alignment(horizontal="right")

    row = 4 + len(kpis) + 1
    first, last = write_table(ws, monthly, row, title="Monthly Sales", formats={
        "revenue": CURRENCY_FMT, "mom_change": PERCENT_FMT,
        "orders": INTEGER_FMT, "units_sold": INTEGER_FMT})
    add_data_bar(ws, 4, first, last)
    # Month-over-month change: green when up, red when down
    ws.conditional_formatting.add(f"E{first}:E{last}", CellIsRule(
        operator="greaterThan", formula=["0"], font=Font(color="1E7B34")))
    ws.conditional_formatting.add(f"E{first}:E{last}", CellIsRule(
        operator="lessThan", formula=["0"], font=Font(color="C00000")))

    row = last + 2
    first, last = write_table(ws, locations, row, title="Sales by Location", formats={
        "revenue": CURRENCY_FMT, "revenue_share": PERCENT_FMT, "orders": INTEGER_FMT})
    add_data_bar(ws, 4, first, last)
    autofit_columns(ws)
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 26


def sheet_product_analysis(wb: Workbook, products: pd.DataFrame, categories: pd.DataFrame) -> None:
    ws = wb.create_sheet("Product Analysis")
    ws["A1"] = "Product Analysis"
    ws["A1"].font = TITLE_FONT
    fmts = {"revenue": CURRENCY_FMT, "avg_unit_price": CURRENCY_FMT,
            "revenue_share": PERCENT_FMT, "orders": INTEGER_FMT, "units_sold": INTEGER_FMT}

    first, last = write_table(ws, products, 3, title="Revenue by Product", formats=fmts)
    cols = list(products.columns)
    rev = get_column_letter(cols.index("revenue") + 1)
    ws.conditional_formatting.add(f"{rev}{first}:{rev}{last}", ColorScaleRule(
        start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50,
        mid_color="FFEB84", end_type="max", end_color="63BE7B"))
    add_data_bar(ws, cols.index("revenue_share") + 1, first, last)
    ws.freeze_panes = "A5"

    first, last = write_table(ws, categories, last + 3, title="Revenue by Category", formats=fmts)
    add_data_bar(ws, list(categories.columns).index("revenue") + 1, first, last)
    autofit_columns(ws)


def sheet_rejected(wb: Workbook, rejected: pd.DataFrame) -> None:
    ws = wb.create_sheet("Rejected Records")
    write_table(ws, rejected, 1)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    autofit_columns(ws)


def save_workbook(wb: Workbook, path: Path) -> None:
    try:
        wb.save(path)
    except PermissionError as exc:
        raise PipelineError(
            f"Cannot save {path.name} - it is probably open in Excel. Close it and run again."
        ) from exc


# ---------------------------------------------------------------------------
# 11. Charts (Matplotlib)
# ---------------------------------------------------------------------------
def _style_axes(ax, title: str, subtitle: str) -> None:
    ax.set_facecolor(CHART_SURFACE)
    for side in ["top", "right", "left"]:
        ax.spines[side].set_visible(False)
    ax.spines["bottom"].set_color(GRID_COLOR)
    ax.tick_params(colors=INK_SECONDARY, length=0, labelsize=10)
    ax.set_title(title, loc="left", fontsize=14, fontweight="bold", color=INK_PRIMARY, pad=24)
    ax.text(0, 1.02, subtitle, transform=ax.transAxes, fontsize=10, color=INK_SECONDARY)


def _money(x, _pos=None) -> str:
    return f"${x / 1000:,.0f}K" if abs(x) >= 1000 else f"${x:,.0f}"


def horizontal_bar_chart(labels, values, title: str, subtitle: str, path: Path) -> None:
    """Sorted horizontal bars (largest at the top) with value labels at the bar ends."""
    fig, ax = plt.subplots(figsize=(10, 0.45 * len(labels) + 1.8), facecolor=CHART_SURFACE)
    y = range(len(labels))[::-1]  # largest value at the top
    ax.barh(list(y), values, color=CHART_COLOR, height=0.68)
    ax.set_yticks(list(y), labels)
    ax.xaxis.set_major_formatter(FuncFormatter(_money))
    ax.grid(axis="x", color=GRID_COLOR, linewidth=0.8)
    ax.set_axisbelow(True)
    for yi, v in zip(y, values):
        ax.text(v, yi, f"  ${v:,.0f}", va="center", fontsize=9, color=INK_SECONDARY)
    ax.set_xlim(0, max(values) * 1.15)
    _style_axes(ax, title, subtitle)
    fig.tight_layout()
    fig.savefig(path, dpi=150, facecolor=CHART_SURFACE)
    plt.close(fig)


def monthly_trend_chart(monthly: pd.DataFrame, path: Path) -> None:
    fig, ax = plt.subplots(figsize=(10, 5), facecolor=CHART_SURFACE)
    months = pd.to_datetime(monthly["month"])
    ax.plot(months, monthly["revenue"], color=CHART_COLOR, linewidth=2,
            marker="o", markersize=7, markeredgecolor=CHART_SURFACE, markeredgewidth=2)
    ax.yaxis.set_major_formatter(FuncFormatter(_money))
    ax.set_ylim(0, monthly["revenue"].max() * 1.2)
    ax.grid(axis="y", color=GRID_COLOR, linewidth=0.8)
    ax.set_axisbelow(True)
    ax.set_xticks(months, [m.strftime("%b\n%Y") if i == 0 else m.strftime("%b")
                           for i, m in enumerate(months)])
    # Label only the highest and lowest months to keep the chart readable
    for idx, offset in [(monthly["revenue"].idxmax(), 12), (monthly["revenue"].idxmin(), -20)]:
        ax.annotate(f"${monthly.at[idx, 'revenue']:,.0f}", (months[idx], monthly.at[idx, "revenue"]),
                    textcoords="offset points", xytext=(0, offset), ha="center",
                    fontsize=9, color=INK_PRIMARY)
    _style_axes(ax, "Monthly Sales Trend", "Revenue per month, valid orders only")
    fig.tight_layout()
    fig.savefig(path, dpi=150, facecolor=CHART_SURFACE)
    plt.close(fig)


def create_charts(products: pd.DataFrame, locations: pd.DataFrame,
                  monthly: pd.DataFrame, chart_dir: Path) -> list[Path]:
    chart_dir.mkdir(parents=True, exist_ok=True)
    paths = [chart_dir / "sales_by_product.png", chart_dir / "sales_by_location.png",
             chart_dir / "monthly_sales_trend.png"]
    horizontal_bar_chart(products["product"].tolist(), products["revenue"].tolist(),
                         "Sales by Product", "Total revenue per product, valid orders only", paths[0])
    # Combine every row with a missing city into one "Unknown city" bar
    known = locations["city"] != UNKNOWN_LOCATION
    labels = [f"{c}, {s}" for c, s in zip(locations.loc[known, "city"], locations.loc[known, "state"])]
    values = locations.loc[known, "revenue"].tolist()
    if (~known).any():
        labels.append("Unknown city")
        values.append(float(locations.loc[~known, "revenue"].sum()))
    horizontal_bar_chart(labels, values,
                         "Sales by Location", "Total revenue per customer city, valid orders only",
                         paths[1])
    monthly_trend_chart(monthly, paths[2])
    for p in paths:
        logger.info("Saved chart: %s", p.name)
    return paths


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------
def run_pipeline(input_path: Path = DEFAULT_INPUT, output_dir: Path = DEFAULT_OUTPUT_DIR) -> dict:
    """Run every step and return a summary dict (used by main() and the tests)."""
    output_dir.mkdir(parents=True, exist_ok=True)
    file_handler = setup_logging(output_dir / LOG_NAME)
    try:
        ctx = CleaningContext()
        logger.info("=== Customer sales data cleaning started ===")

        raw = load_data(input_path)
        missing_raw = inspect_data(raw)

        df = strip_whitespace(raw, ctx)
        df = standardize_text(df, ctx)
        df = clean_emails(df, ctx)
        df = clean_phones(df, ctx)
        df = parse_dates(df, ctx)
        df = validate_quantities(df, ctx)
        df = validate_prices(df, ctx)
        df = handle_missing_fields(df, ctx)
        df, duplicates = remove_duplicates(df, ctx)
        clean, rejected = split_valid_rejected(df, raw, ctx)

        if clean.empty:
            raise PipelineError("No valid records remain after cleaning - check the input file.")

        products = build_product_analysis(clean)
        categories = build_category_analysis(clean)
        monthly = build_monthly_sales(clean)
        locations = build_location_sales(clean)
        kpis = build_kpis(clean, len(raw))

        # Missing values before vs after (after = blanks left in the clean data)
        clean_missing = [int(clean[c].isna().sum()) if c in clean else 0 for c in REQUIRED_COLUMNS]
        missing_raw["Blank After Cleaning"] = clean_missing

        reasons = (pd.Series([r.split(" (")[0] for lst in ctx.rejects.values() for r in lst],
                             dtype=object)
                   .value_counts().rename_axis("Reason").reset_index(name="Rows"))

        overview = [
            ("Raw rows", len(raw), INTEGER_FMT),
            ("Duplicate rows removed", len(duplicates), INTEGER_FMT),
            ("Invalid rows rejected", len(rejected), INTEGER_FMT),
            ("Clean rows exported", len(clean), INTEGER_FMT),
            ("Rows kept with a data note", int((clean["data_notes"] != "").sum()), INTEGER_FMT),
            ("Data retention (clean / raw)", len(clean) / len(raw), PERCENT_FMT),
        ]
        log_df = pd.DataFrame(ctx.log_entries)

        wb = Workbook()
        sheet_cleaned_data(wb, clean)
        sheet_quality_report(wb, overview, missing_raw, reasons, log_df)
        sheet_sales_summary(wb, kpis, monthly, locations)
        sheet_product_analysis(wb, products, categories)
        sheet_rejected(wb, rejected)
        excel_path = output_dir / EXCEL_NAME
        save_workbook(wb, excel_path)
        logger.info("Saved Excel report: %s", excel_path.name)

        csv_path = output_dir / CLEAN_CSV_NAME
        try:
            clean.to_csv(csv_path, index=False, encoding="utf-8")
        except PermissionError as exc:
            raise PipelineError(f"Cannot save {csv_path.name} - close it in Excel and retry.") from exc

        charts = create_charts(products, locations, monthly, output_dir / CHART_DIR_NAME)

        summary = {
            "raw_rows": len(raw), "duplicates_removed": len(duplicates),
            "rejected_rows": len(rejected), "clean_rows": len(clean),
            "total_revenue": round(float(clean["total_sales"].sum()), 2),
            "excel_path": excel_path, "csv_path": csv_path, "charts": charts,
        }
        logger.info("=== Done: %d raw -> %d duplicates removed, %d rejected, %d clean rows ===",
                    summary["raw_rows"], summary["duplicates_removed"],
                    summary["rejected_rows"], summary["clean_rows"])
        return summary
    finally:
        # Release the log file (important on Windows, where open files are locked)
        file_handler.close()
        logger.removeHandler(file_handler)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Clean and analyse messy sales data.")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT,
                        help="Raw CSV file (default data/raw_sales_data.csv)")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR,
                        help="Folder for the report, charts and log (default output/)")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        summary = run_pipeline(args.input, args.output_dir)
    except PipelineError as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        return 1
    except Exception:  # unexpected bug: show the full traceback for debugging
        logger.exception("Unexpected error - the pipeline stopped.")
        return 2

    print("\nCleaning complete.")
    print(f"  Raw rows:            {summary['raw_rows']}")
    print(f"  Duplicates removed:  {summary['duplicates_removed']}")
    print(f"  Invalid rows:        {summary['rejected_rows']} (see 'Rejected Records' sheet)")
    print(f"  Clean rows:          {summary['clean_rows']}")
    print(f"  Total revenue:       ${summary['total_revenue']:,.2f}")
    print(f"  Excel report:        {summary['excel_path']}")
    print(f"  Charts folder:       {summary['charts'][0].parent}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
