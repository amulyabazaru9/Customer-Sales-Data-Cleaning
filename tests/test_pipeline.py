"""
Automated checks for the data cleaning pipeline.

Run from the project root:
    python -m pytest -v
"""

import re

import pandas as pd
import pytest
from openpyxl import load_workbook

import clean_data
import generate_dataset

EXPECTED_SHEETS = ["Cleaned Data", "Data Quality Report", "Sales Summary",
                   "Product Analysis", "Rejected Records"]


# ---------------------------------------------------------------------------
# Fixtures: generate a fresh dataset and run the pipeline in a temp folder,
# so the tests never touch the real data/ or output/ folders.
# ---------------------------------------------------------------------------
@pytest.fixture(scope="module")
def workspace(tmp_path_factory):
    tmp = tmp_path_factory.mktemp("run")
    raw_csv = tmp / "raw_sales_data.csv"
    generate_dataset.save_csv(generate_dataset.generate_rows(500, seed=42), raw_csv)
    summary = clean_data.run_pipeline(raw_csv, tmp / "output")
    clean = pd.read_excel(summary["excel_path"], sheet_name="Cleaned Data")
    rejected = pd.read_excel(summary["excel_path"], sheet_name="Rejected Records")
    return {"tmp": tmp, "raw_csv": raw_csv, "summary": summary,
            "clean": clean, "rejected": rejected}


# ---------------------------------------------------------------------------
# Dataset generation
# ---------------------------------------------------------------------------
def test_generator_creates_500_rows_with_required_columns(workspace):
    raw = pd.read_csv(workspace["raw_csv"], dtype=str, keep_default_na=False)
    assert len(raw) == 500
    assert list(raw.columns) == generate_dataset.COLUMNS


def test_generator_is_reproducible():
    assert generate_dataset.generate_rows(100, seed=7) == generate_dataset.generate_rows(100, seed=7)


def test_generated_data_actually_contains_quality_problems(workspace):
    raw = pd.read_csv(workspace["raw_csv"], dtype=str, keep_default_na=False)
    assert (raw["customer_name"] == "").any(), "expected missing names"
    assert raw["unit_price"].str.startswith("$").any(), "expected '$' prices"
    assert raw["unit_price"].str.startswith("-").any(), "expected negative prices"
    assert raw["customer_name"].str.contains(r"^\s|\s$|\s{2,}").any(), "expected extra spaces"
    assert not raw["order_date"].str.match(r"^\d{4}-\d{2}-\d{2}$").all(), "expected mixed dates"


# ---------------------------------------------------------------------------
# Cleaning results
# ---------------------------------------------------------------------------
def test_every_raw_row_is_accounted_for(workspace):
    s = workspace["summary"]
    assert s["raw_rows"] == s["duplicates_removed"] + s["rejected_rows"] + s["clean_rows"]
    assert s["duplicates_removed"] > 0
    assert s["rejected_rows"] > 0


def test_duplicates_are_removed(workspace):
    clean = workspace["clean"]
    assert not clean["Order ID"].duplicated().any()
    assert not clean.drop(columns=["Source Row"]).duplicated().any()


def test_prices_and_quantities_are_valid(workspace):
    clean = workspace["clean"]
    assert (clean["Unit Price"] > 0).all()
    assert (clean["Unit Price"] <= clean_data.MAX_UNIT_PRICE).all()
    assert (clean["Quantity"] >= 1).all()
    assert (clean["Quantity"] % 1 == 0).all()
    expected_total = (clean["Quantity"] * clean["Unit Price"]).round(2)
    assert (clean["Total Sales"] - expected_total).abs().max() < 0.01


def test_text_fields_are_standardised(workspace):
    clean = workspace["clean"]
    for col in ["Customer Name", "Product", "City", "Order ID"]:
        values = clean[col].astype(str)
        assert not values.str.contains(r"^\s|\s$|\s{2,}").any(), f"extra spaces in {col}"
    assert clean["Order ID"].str.match(r"^ORD-\d+$").all()
    assert clean["State"].str.isupper().all()
    # Each product should now have exactly one spelling
    assert clean["Product"].str.lower().nunique() == clean["Product"].nunique()


def test_emails_and_phones_are_valid_or_blank(workspace):
    clean = workspace["clean"]
    emails = clean["Customer Email"].dropna()
    assert emails.str.match(clean_data.EMAIL_PATTERN).all()
    phones = clean["Phone"].dropna()
    assert phones.str.match(r"^\(\d{3}\) \d{3}-\d{4}$").all()


def test_dates_are_real_dates(workspace):
    dates = pd.to_datetime(workspace["clean"]["Order Date"], errors="coerce")
    assert dates.notna().all()


def test_invalid_records_are_kept_with_a_reason(workspace):
    rejected = workspace["rejected"]
    assert len(rejected) == workspace["summary"]["rejected_rows"]
    assert rejected["Rejection Reason"].notna().all()
    # A rejected row must never also appear in the clean data
    assert not set(rejected["Source Row"]) & set(workspace["clean"]["Source Row"])


# ---------------------------------------------------------------------------
# Outputs
# ---------------------------------------------------------------------------
def test_excel_has_all_sheets_and_formatting(workspace):
    wb = load_workbook(workspace["summary"]["excel_path"])
    assert wb.sheetnames == EXPECTED_SHEETS
    ws = wb["Cleaned Data"]
    assert ws.auto_filter.ref is not None
    assert ws.freeze_panes == "A2"
    assert len(ws.conditional_formatting) > 0


def test_charts_are_saved(workspace):
    charts = workspace["summary"]["charts"]
    assert len(charts) == 3
    for chart in charts:
        assert chart.exists() and chart.stat().st_size > 10_000
        assert chart.read_bytes()[:8] == b"\x89PNG\r\n\x1a\n"


def test_repeated_runs_give_identical_results(workspace):
    again = clean_data.run_pipeline(workspace["raw_csv"], workspace["tmp"] / "output")
    for key in ["raw_rows", "duplicates_removed", "rejected_rows", "clean_rows", "total_revenue"]:
        assert again[key] == workspace["summary"][key]


# ---------------------------------------------------------------------------
# Error handling and individual functions
# ---------------------------------------------------------------------------
def test_missing_input_file_gives_clear_error(tmp_path):
    with pytest.raises(clean_data.PipelineError, match="not found"):
        clean_data.run_pipeline(tmp_path / "nope.csv", tmp_path / "out")


def test_missing_columns_gives_clear_error(tmp_path):
    bad = tmp_path / "bad.csv"
    bad.write_text("order_id,product\nORD-1,Mouse\n", encoding="utf-8")
    with pytest.raises(clean_data.PipelineError, match="missing required columns"):
        clean_data.run_pipeline(bad, tmp_path / "out")


def test_parse_dates_handles_every_format():
    raw = ["2025-03-14", "14/03/2025", "03-14-2025", "March 14, 2025", "2025/03/14",
           "14-Mar-2025", "2025-13-45"]
    df = pd.DataFrame({"order_date": raw}, index=range(2, 2 + len(raw)))
    ctx = clean_data.CleaningContext()
    out = clean_data.parse_dates(df, ctx)
    assert (out["order_date"].iloc[:6] == pd.Timestamp("2025-03-14")).all()
    assert pd.isna(out["order_date"].iloc[6])
    assert 8 in ctx.rejects  # the impossible date is rejected


def test_clean_phones_normalises_formats():
    raw = ["(555) 123-4567", "555.123.4567", "+1-555-123-4567", "5551234567", "12345"]
    df = pd.DataFrame({"phone": raw}, index=range(2, 2 + len(raw)))
    out = clean_data.clean_phones(df, clean_data.CleaningContext())
    assert out["phone"].iloc[:4].eq("(555) 123-4567").all()
    assert pd.isna(out["phone"].iloc[4])
    assert re.match(r"^\(\d{3}\)", out["phone"].iloc[0])
