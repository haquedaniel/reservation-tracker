"""Airbnb reservation tracker.

This script reads the reservation workbook, validates the input sheets,
builds clean analytical outputs for Excel/Power BI, and optionally sends
a weekly HTML email.

Run modes:
    python reservation-tracker.py <workbook_path> update
    python reservation-tracker.py <workbook_path> weekly-email

Design notes:
- Excel remains the input layer.
- Python owns validation and transformation.
- Power BI / email consume clean output tables.
"""

from pathlib import Path
from datetime import datetime, timedelta
from email.message import EmailMessage

import calendar
import os
import smtplib
import sys
import argparse
import json

import pandas as pd
from dotenv import load_dotenv

import matplotlib.pyplot as plt
import base64

import gspread
from google.oauth2.service_account import Credentials

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None


BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")


def load_table(source: str, input_path: Path | None, sheet_id: str | None, tab_name: str) -> pd.DataFrame:
    """
    Load a named table from either Excel or Google Sheets.
    """

    if source == "excel":
        if input_path is None:
            raise ValueError("input_path is required for Excel source")
        return pd.read_excel(input_path, sheet_name=tab_name)

    if source == "google-sheets":
        if sheet_id is None:
            raise ValueError("sheet_id is required for Google Sheets source")
        return load_google_sheet_tab(sheet_id, tab_name)

    raise ValueError(f"Unknown source: {source}")



# =============================================================================
# General utilities
# =============================================================================

def parse_args():
    parser = argparse.ArgumentParser(description="Reservation tracker")

    parser.add_argument(
        "--input",
        type=str,
        default="Reservations.xlsm",
        help="Path to input workbook"
    )

    parser.add_argument(
        "--output",
        type=str,
        default=None,
        help="Output directory (default = iCloud folder)"
    )

    parser.add_argument(
        "--mode",
        type=str,
        default="weekly-email",
        choices=["update", "weekly-email", "finance-email"],
        help="Run mode"
    )

    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Generate email but do not send"
    )

    parser.add_argument(
    "--source",
    choices=["excel", "google-sheets"],
    default="excel",
    )

    parser.add_argument(
        "--sheet-id",
        default=os.environ.get("GOOGLE_SHEET_ID"),
    )

    parser.add_argument(
        "--groups",
        type=str,
        default=None,
        help="Comma-separated listing groups to include in finance reports, e.g. Voilerie,Peskerezh",
    )

    parser.add_argument(
        "--ai-insights",
        action="store_true",
        help="Add optional GenAI insights to the finance email. Requires OPENAI_API_KEY.",
    )

    return parser.parse_args()






GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly"
]


def get_google_sheets_client():
    credentials_env = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]

    try:
        credentials_dict = json.loads(credentials_env)

        credentials = Credentials.from_service_account_info(
            credentials_dict,
            scopes=GOOGLE_SCOPES,
        )

        print("Using service account JSON from environment")

    except json.JSONDecodeError:
        credentials_path = credentials_env.strip()

        credentials = Credentials.from_service_account_file(
            credentials_path,
            scopes=GOOGLE_SCOPES,
        )

        print("Using service account JSON file")

    return gspread.authorize(credentials)


def load_google_sheet_tab(sheet_id: str, tab_name: str) -> pd.DataFrame:
    """
    Load one Google Sheet tab into a pandas DataFrame.

    The first row must contain headers, just like your Excel sheets.
    """

    client = get_google_sheets_client()
    spreadsheet = client.open_by_key(sheet_id)
    worksheet = spreadsheet.worksheet(tab_name)

    rows = worksheet.get_all_records()

    return pd.DataFrame(rows)



def normalise_listing_key(value) -> str | None:
    """
    Convert Excel listing values to a stable string key.
    Examples:
    - 2 -> "2"
    - 2.0 -> "2"
    - blank -> None
    """
    if pd.isna(value):
        return None

    try:
        number = float(value)
        if number.is_integer():
            return str(int(number))
    except (ValueError, TypeError):
        pass

    return str(value).strip()


def parse_group_filter(groups_arg: str | None) -> list[str] | None:
    """Parse comma-separated group names supplied on the command line."""
    if not groups_arg:
        return None

    groups = [group.strip() for group in groups_arg.split(",") if group.strip()]
    return groups or None


def filter_listings_by_group(listings: pd.DataFrame, selected_groups: list[str] | None) -> pd.DataFrame:
    """Return listings belonging to the selected groups. If no groups are supplied, return all listings."""
    listings = listings.copy()

    if "group" not in listings.columns:
        listings["group"] = "All"

    listings["group"] = listings["group"].fillna("Ungrouped").astype(str).str.strip()

    if not selected_groups:
        return listings

    selected_lower = {group.lower() for group in selected_groups}
    filtered = listings[listings["group"].str.lower().isin(selected_lower)].copy()

    if filtered.empty:
        available = sorted(listings["group"].dropna().unique())
        raise ValueError(
            f"No listings matched --groups={selected_groups}. Available groups: {available}"
        )

    return filtered


def group_label(selected_groups: list[str] | None) -> str:
    """Human-readable label for report titles and filenames."""
    if not selected_groups:
        return "All groups"

    return ", ".join(selected_groups)


def group_slug(selected_groups: list[str] | None) -> str:
    """Filesystem-safe label for grouped output files."""
    if not selected_groups:
        return "all-groups"

    slug = "-".join(selected_groups).lower()
    for char in [" ", "/", "\\", ":", ";", ","]:
        slug = slug.replace(char, "-")
    while "--" in slug:
        slug = slug.replace("--", "-")
    return slug.strip("-") or "selected-groups"


def get_output_dir() -> Path:
    """
    Return the iCloud Drive output folder.

    Adjust the last two folder names if you want a different location.
    """
    output_dir = (
        Path.home()
        / "Library"
        / "Mobile Documents"
        / "com~apple~CloudDocs"
        / "airbnb-tracker"
        / "output"
    )
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def validate_required_columns(df: pd.DataFrame, required_columns: list[str]) -> None:
    """
    Check that the expected Excel columns exist before we do any transformations.

    This is a schema check:
    - are the columns we depend on actually present?
    """
    missing = [col for col in required_columns if col not in df.columns]

    if missing:
        raise ValueError(f"Missing required columns: {missing}")


def drop_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remove rows that are completely empty.

    Why:
    - Excel sheets often contain blank lines
    - they add noise and can trigger false validation errors
    """
    df = df.copy()
    df = df.dropna(how="all")
    return df


def parse_types(df: pd.DataFrame, date_columns, numeric_columns) -> pd.DataFrame:
    """
    Convert dates and numeric columns to the right types.

    Important:
    - Excel data may come in as text, mixed types, or blanks
    - errors='coerce' turns invalid values into NaN / NaT instead of crashing
    """
    df = df.copy()

    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)


    for col in numeric_columns:
        if col in df.columns:
            # Replace comma decimals if Excel/text import gives strings like "218,85"
            if df[col].dtype == "object":
                df[col] = df[col].astype(str).str.replace(",", ".", regex=False)

            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


# =============================================================================
# Workbook loaders
# =============================================================================

def load_reservations(source: str, input_path: Path | None, sheet_id: str | None, sheet_name: str = "Reservations") -> pd.DataFrame:
    return load_table(source, input_path, sheet_id, sheet_name)


def load_targets(source: str, input_path: Path | None, sheet_id: str | None, sheet_name: str = "Monthly Targets") -> pd.DataFrame:
    return load_table(source, input_path, sheet_id, sheet_name)


def load_fixed_expenses(source: str, input_path: Path | None, sheet_id: str | None, sheet_name: str = "Fixed Expenses") -> pd.DataFrame:
    return load_table(source, input_path, sheet_id, sheet_name)


def load_variable_expenses(source: str, input_path: Path | None, sheet_id: str | None, sheet_name: str = "Variable Expenses") -> pd.DataFrame:
    return load_table(source, input_path, sheet_id, sheet_name)


def load_listings_lookup(source: str, input_path: Path | None, sheet_id: str | None) -> pd.DataFrame:
    df = load_table(source, input_path, sheet_id, "Listings")

    df.columns = df.columns.str.strip()

    df = df.rename(columns={
        "Listing": "listing",
        "Name": "listing_name",
        "Description AirBnB": "listing_description",
        "Group": "group",
        "Groupe": "group",
        "group": "group",
    })

    validate_required_columns(df, ["listing", "listing_name"])

    if "listing_description" not in df.columns:
        df["listing_description"] = pd.NA

    if "group" not in df.columns:
        df["group"] = "All"

    df["listing"] = df["listing"].apply(normalise_listing_key)
    df["group"] = df["group"].fillna("Ungrouped").astype(str).str.strip()
    df.loc[df["group"].eq(""), "group"] = "Ungrouped"

    return df



# =============================================================================
# Listing lookup / enrichment
# =============================================================================


def enrich_with_listing_info(df: pd.DataFrame, listings: pd.DataFrame) -> pd.DataFrame:
    """
    Add listing_name and listing_description to a dataframe that already has
    a standardised 'listing' column.
    """
    df = df.copy()
    listings = listings.copy()

    df["listing"] = df["listing"].apply(normalise_listing_key)
    listings["listing"] = listings["listing"].apply(normalise_listing_key)

    listing_cols = ["listing", "listing_name", "listing_description"]
    if "group" in listings.columns:
        listing_cols.append("group")

    enriched = df.merge(
        listings[listing_cols],
        on="listing",
        how="left"
    )

    enriched["listing_name"] = enriched["listing_name"].fillna(enriched["listing"])
    if "group" in enriched.columns:
        enriched["group"] = enriched["group"].fillna("Ungrouped")

    return enriched


# =============================================================================
# Reservations cleaning, validation, and expansion
# =============================================================================


def rename_reservation_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Rename Excel-style column names to Python-friendly names.

    Why do this?
    - avoids spaces in column names
    - makes code easier to read and write
    """
    df = df.copy()

    column_map = {
        "Guest Name": "guest_name",
        "Listing": "listing",
        "Booking Source": "booking_source",
        "Confirmation Code": "confirmation_code",
        "Check in Date": "checkin_date",
        "Number of Nights": "nights",
        "Total Revenue": "total_revenue",
        "Cleaning Fees": "cleaning_fees",
        "Concierge Commission": "concierge_commission",
        "Revenue Net": "revenue_net",
        "Booking Date": "booking_date",
        "Reservation Date": "booking_date",
    }

    df = df.rename(columns=column_map)
    return df


def add_optional_reservation_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ensure optional reservation columns exist so downstream code does not fail
    when the Excel workbook does not contain them.
    """
    df = df.copy()

    optional_columns = [
        "guest_name",
        "booking_source",
        "confirmation_code",
        "total_revenue",
        "cleaning_fees",
        "concierge_commission",
        "booking_date",
    ]

    for col in optional_columns:
        if col not in df.columns:
            df[col] = pd.NA

    return df


def add_checkout_date(df: pd.DataFrame) -> pd.DataFrame:
    """
    Derive checkout_date using hotel logic:

    checkout_date = checkin_date + nights

    Example:
    - checkin: 2026-01-30
    - nights: 3
    - checkout: 2026-02-02
    """
    df = df.copy()

    df["checkout_date"] = df["checkin_date"] + pd.to_timedelta(df["nights"], unit="D")
    return df


def validate_reservations(df: pd.DataFrame) -> None:
    """
    Validate business rules on the reservations table.

    Hard errors:
    - listing must be present
    - checkin_date must be valid
    - nights must be > 0
    - revenue_net must be numeric
    - booking_date, if present, must be <= checkin_date
    """
    errors = []

    for idx, row in df.iterrows():
        row_num = idx + 2  # +2 because Excel row numbers usually start after header

        # Listing must exist
        if pd.isna(row["listing"]):
            errors.append(f"Row {row_num}: missing listing")

        # Check-in date must exist and be valid
        if pd.isna(row["checkin_date"]):
            errors.append(f"Row {row_num}: invalid or missing checkin_date")

        # Nights must exist and be > 0
        if pd.isna(row["nights"]):
            errors.append(f"Row {row_num}: invalid or missing nights")
        elif row["nights"] <= 0:
            errors.append(f"Row {row_num}: nights must be > 0")

        # Revenue net must exist and be numeric
        if pd.isna(row["revenue_net"]):
            errors.append(f"Row {row_num}: invalid or missing revenue_net")

        # Booking date can be empty, but if present it must be <= check-in date
        if not pd.isna(row["booking_date"]) and not pd.isna(row["checkin_date"]):
            if row["booking_date"] > row["checkin_date"]:
                errors.append(
                    f"Row {row_num}: booking_date is after checkin_date"
                )

    if errors:
        error_text = "\n".join(errors)
        raise ValueError(f"Reservation validation failed:\n{error_text}")


def find_overlaps(df: pd.DataFrame) -> pd.DataFrame:
    """
    Detect overlapping bookings for the same listing.

    Logic:
    - group by listing
    - sort by checkin_date
    - compare each booking to the previous one
    - overlap exists if current checkin < previous checkout
    """
    overlap_rows = []

    for listing, group in df.groupby("listing"):
        group = group.sort_values("checkin_date")

        previous_row = None

        for _, row in group.iterrows():
            if previous_row is None:
                previous_row = row
                continue

            if row["checkin_date"] < previous_row["checkout_date"]:
                overlap_rows.append(
                    {
                        "listing": listing,
                        "previous_guest": previous_row.get("guest_name"),
                        "previous_checkin": previous_row["checkin_date"],
                        "previous_checkout": previous_row["checkout_date"],
                        "current_guest": row.get("guest_name"),
                        "current_checkin": row["checkin_date"],
                        "current_checkout": row["checkout_date"],
                    }
                )

            previous_row = row

    return pd.DataFrame(overlap_rows)


def expand_booking_nights(df: pd.DataFrame) -> pd.DataFrame:
    """
    Expand each booking into one row per occupied night.

    Output:
    - one row per listing per night
    """

    rows = []  # we collect rows here, then build a DataFrame at the end

    for _, row in df.iterrows():

        start_date = row["checkin_date"]
        end_date = row["checkout_date"]

        # Compute daily values once per booking
        daily_revenue_net = row["revenue_net"] / row["nights"]
 
        if not pd.isna(row.get("concierge_commission")):
            daily_concierge_commission = row["concierge_commission"] / row["nights"]
        else:
            daily_concierge_commission = 0

        if not pd.isna(row.get("cleaning_fees")):
            daily_cleaning_fees = row["cleaning_fees"] / row["nights"]
        else:
            daily_cleaning_fees = 0 
        
        if not pd.isna(row.get("total_revenue")):
            daily_revenue_total = row["total_revenue"] / row["nights"]
        else:
            daily_revenue_total = pd.NA

        # Create date range 
        dates = pd.date_range(start=start_date, end=end_date, inclusive="left")

        for date in dates:
            rows.append({
                "date": date,
                "year": date.year,
                "month": date.month,
                "listing": normalise_listing_key(row["listing"]),
                "guest_name": row.get("guest_name"),
                "booking_source": row.get("booking_source"),
                "confirmation_code": row.get("confirmation_code"),
                "booking_date": row.get("booking_date"),
                "checkin_date": start_date,
                "checkout_date": end_date,
                "nights": row["nights"],
                "occupied": 1,
                "daily_concierge_commission": daily_concierge_commission,
                "daily_cleaning_fees": daily_cleaning_fees,
                "daily_revenue_net": daily_revenue_net,
                "daily_revenue_total": daily_revenue_total,
            })

    # Convert list of dicts into DataFrame
    expanded_df = pd.DataFrame(rows)
    if not expanded_df.empty:
        expanded_df["listing"] = expanded_df["listing"].apply(normalise_listing_key)
    return expanded_df


# =============================================================================
# Calendar, occupancy, and targets
# =============================================================================


def build_calendar(year, df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    df = df.copy()
    start_date = datetime(int(year), 1, 1)
    end_date = datetime(int(year), 12, 31)
    dates = pd.date_range(start=start_date, end=end_date)

    for listing in df["listing"].dropna().apply(normalise_listing_key).unique():

        for date in dates:
            rows.append({
                "date": date,
                "year": date.year,
                "month": date.month,
                "listing": listing,
            })

    calendar = pd.DataFrame(rows)
    calendar = calendar.sort_values("listing")
    calendar["listing"] = calendar["listing"].apply(normalise_listing_key)
    return calendar


def build_daily_occupancy(calendar: pd.DataFrame, booking_nights: pd.DataFrame) -> pd.DataFrame:
    """
    Merge full calendar with booking nights to create a complete daily occupancy table.

    Output:
    - one row per date + listing
    - includes both occupied and unoccupied nights
    """
    calendar = calendar.copy()
    booking_nights = booking_nights.copy()
    calendar["listing"] = calendar["listing"].apply(normalise_listing_key)
    booking_nights["listing"] = booking_nights["listing"].apply(normalise_listing_key)
    daily = calendar.merge(
        booking_nights,
        on=["date", "month", "year", "listing"],
        how="left"
    )

    # Fill missing values (i.e. unoccupied nights)
    daily["occupied"] = daily["occupied"].fillna(0).astype(int)
    daily["daily_revenue_net"] = daily["daily_revenue_net"].fillna(0)
    daily["daily_revenue_total"] = daily["daily_revenue_total"].fillna(0)
    daily["daily_cleaning_fees"] = daily["daily_cleaning_fees"].fillna(0)
    daily["daily_concierge_commission"] = daily["daily_concierge_commission"].fillna(0)

    return daily


def rename_target_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Rename Excel-style column names to Python-friendly names.

    Why do this?
    - avoids spaces in column names
    - makes code easier to read and write
    """
    df = df.copy()

    column_map = {
        "Month": "month",
        "Year": "year",
    }

    df = df.rename(columns=column_map)
    return df


def expand_targets(df: pd.DataFrame) -> pd.DataFrame:
    """
    targets=[]
    for _,row in df.iterrows():

        for col in df.columns[2:]:
            listing = str(col)
            targets.append([int(row["year"]), int(row["month"]), listing, row[col]])

    df = pd.DataFrame(targets)
    return df
    """
    df = df.dropna(subset=["year", "month"])
    melted = df.melt(id_vars=["year", "month"], var_name="listing", value_name="target")
    melted["listing"] = melted["listing"].apply(normalise_listing_key)
    return melted


def build_daily_targets(df_calendar: pd.DataFrame, monthly_targets: pd.DataFrame) -> pd.DataFrame:
    """
    Merge full calendar with targets to create a complete daily targets table.

    Output:
    - one row per date + listing
    """
    df_calendar = df_calendar.copy()
    monthly_targets = monthly_targets.copy()
    df_calendar["listing"] = df_calendar["listing"].apply(normalise_listing_key)
    monthly_targets["listing"] = monthly_targets["listing"].apply(normalise_listing_key)

    monthly_targets["days_in_month"] = monthly_targets.apply(
        lambda row: calendar.monthrange(int(row["year"]), int(row["month"]))[1],
        axis=1
    )

    monthly_targets["daily_target"]=monthly_targets["target"] / monthly_targets["days_in_month"]

    daily_targets = df_calendar.merge(
        monthly_targets,
        on=["year", "month", "listing"],
        how="left"
    )

    return daily_targets


# =============================================================================
# Expense processing
# =============================================================================


def _month_columns(df: pd.DataFrame) -> list:
    """Return columns that look like YYYY-MM month columns."""
    return [col for col in df.columns if str(col).startswith("2026-")]


def expand_fixed_expenses(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convert the Fixed Expenses sheet from wide monthly format to long monthly format.

    Fixed expenses are real monthly costs. They are summed directly for monthly/YTD
    cost reporting.
    """
    df = df.copy()
    df.columns = df.columns.str.strip()

    id_vars = ["Description", "Listing"]
    month_cols = _month_columns(df)

    df_long = df.melt(
        id_vars=id_vars,
        value_vars=month_cols,
        var_name="year_month",
        value_name="monthly_expense"
    )

    df_long["monthly_expense"] = pd.to_numeric(df_long["monthly_expense"], errors="coerce").fillna(0)
    df_long["year"] = df_long["year_month"].astype(str).str[:4].astype(int)
    df_long["month"] = df_long["year_month"].astype(str).str[5:7].astype(int)
    df_long["listing"] = df_long["Listing"].apply(normalise_listing_key)
    df_long["description"] = df_long["Description"]
    df_long["expense_type"] = "fixed"
    df_long["version"] = "actual"
    df_long["selected_version"] = "actual"
    df_long["allocated_monthly_expense"] = df_long["monthly_expense"]

    df_long["days_in_month"] = df_long.apply(
        lambda row: calendar.monthrange(int(row["year"]), int(row["month"]))[1],
        axis=1
    )
    df_long["daily_expense"] = df_long["allocated_monthly_expense"] / df_long["days_in_month"]

    return df_long


def expand_variable_expenses(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convert the Variable Expenses sheet from wide monthly format to long monthly format.

    The sheet contains estimate and actual rows. For each listing + description + month:
    - use actual when present
    - otherwise use estimate

    The resulting monthly_expense is the selected monthly variable cost used for
    reporting and allocation to bookings.
    """
    df = df.copy()
    df.columns = df.columns.str.strip()

    id_vars = ["Description", "Listing", "Type"]
    month_cols = _month_columns(df)

    df_long = df.melt(
        id_vars=id_vars,
        value_vars=month_cols,
        var_name="year_month",
        value_name="monthly_expense"
    )

    df_long["monthly_expense"] = pd.to_numeric(df_long["monthly_expense"], errors="coerce")
    df_long["year"] = df_long["year_month"].astype(str).str[:4].astype(int)
    df_long["month"] = df_long["year_month"].astype(str).str[5:7].astype(int)
    df_long["listing"] = df_long["Listing"].apply(normalise_listing_key)
    df_long["description"] = df_long["Description"]
    df_long["version"] = (
        df_long["Type"]
        .astype(str)
        .str.strip()
        .str.lower()
    )
    df_long["expense_type"] = "variable"

    # Future actual cells should be blank, not zero. Blank means "no actual yet".
    estimates = df_long[df_long["version"] == "estimate"].copy()
    actuals = df_long[(df_long["version"] == "actual") & (df_long["monthly_expense"].notna())].copy()

    estimates["version_priority"] = 0
    actuals["version_priority"] = 1

    selected = pd.concat([estimates, actuals], ignore_index=True)
    selected = selected.sort_values("version_priority")
    selected = selected.drop_duplicates(
        subset=["listing", "description", "year", "month"],
        keep="last"
    )

    selected["selected_version"] = selected["version"]
    selected["monthly_expense"] = selected["monthly_expense"].fillna(0)
    selected["allocated_monthly_expense"] = selected["monthly_expense"]

    selected["days_in_month"] = selected.apply(
        lambda row: calendar.monthrange(int(row["year"]), int(row["month"]))[1],
        axis=1
    )

    # This daily value is used when allocating variable costs to occupied reservation nights.
    selected["daily_expense"] = selected["allocated_monthly_expense"] / selected["days_in_month"]

    return selected.drop(columns=["version_priority"], errors="ignore")


def combine_monthly_expenses(
    fixed_expenses: pd.DataFrame,
    variable_expenses: pd.DataFrame
) -> pd.DataFrame:
    """Combine selected fixed and variable monthly expenses for reporting."""
    return pd.concat([fixed_expenses, variable_expenses], ignore_index=True)


def expand_daily_expenses(df_long: pd.DataFrame) -> pd.DataFrame:
    rows = []

    for _, row in df_long.iterrows():
        start_date = datetime(int(row["year"]), int(row["month"]), 1)
        end_date = datetime(
            int(row["year"]),
            int(row["month"]),
            calendar.monthrange(int(row["year"]), int(row["month"]))[1]
        )

        dates = pd.date_range(start=start_date, end=end_date)

        for date in dates:
            rows.append({
                "date": date,
                "year": date.year,
                "month": date.month,
                "listing": normalise_listing_key(row["listing"]),
                "description": row.get("description"),
                "expense_type": row.get("expense_type"),
                "daily_expense": row["daily_expense"],
            })

    return pd.DataFrame(rows)


# =============================================================================
# Weekly reporting calculations
# =============================================================================


def get_week_bounds(today: datetime | None = None) -> tuple[pd.Timestamp, pd.Timestamp]:
    """
    Return Monday-Sunday bounds for the current week.

    Later we can change this to 'next week' if useful for operations.
    """
    if today is None:
        today = datetime.today()

    start = today - timedelta(days=today.weekday())  # Monday
    end = start + timedelta(days=6)                  # Sunday

    return pd.Timestamp(start.date()), pd.Timestamp(end.date())


def build_weekly_arrivals_departures_summary(
    reservations: pd.DataFrame,
    today: datetime | None = None
) -> dict:
    """
    Build arrivals/departures summary from the reservations table.

    Expected columns:
    - guest_name
    - listing
    - checkin_date
    - checkout_date
    - nights
    """

    df = reservations.copy()

    start_of_week, end_of_week = get_week_bounds(today)

    arrivals = df[
        (df["checkin_date"] >= start_of_week) &
        (df["checkin_date"] <= end_of_week)
    ].sort_values(["checkin_date", "listing"])

    departures = df[
        (df["checkout_date"] >= start_of_week) &
        (df["checkout_date"] <= end_of_week)
    ].sort_values(["checkout_date", "listing"])

    return {
        "start_of_week": start_of_week,
        "end_of_week": end_of_week,
        "arrivals": arrivals,
        "departures": departures,
        "num_arrivals": len(arrivals),
        "num_departures": len(departures),
    }


def build_weekly_occupancy(
    daily: pd.DataFrame,
    today: datetime | None = None,
    weeks_ahead: int = 8
) -> pd.DataFrame:

    if today is None:
        today = datetime.today()

    today_ts = pd.Timestamp(today.date())

    # Start from Monday of the current week, not today,
    # so the first displayed week has a full 7-day denominator.
    start_date = today_ts - pd.to_timedelta(today_ts.weekday(), unit="D")

    # Include full weeks
    end_date = start_date + pd.Timedelta(weeks=weeks_ahead)

    df = daily.copy()
    df["date"] = pd.to_datetime(df["date"])

    df = df[
        (df["date"] >= start_date) &
        (df["date"] < end_date)
    ]

    df["week_start"] = df["date"] - pd.to_timedelta(df["date"].dt.weekday, unit="D")
    df["week_end"] = df["week_start"] + pd.Timedelta(days=6)

    weekly = (
        df.groupby(
            ["week_start", "week_end", "listing"],
            as_index=False
        )
        .agg(
            occupied_nights=("occupied", "sum"),
            available_nights=("date", "count"),
        )
    )

    weekly["occupancy_rate"] = (
        weekly["occupied_nights"] / weekly["available_nights"]
    )

    return weekly


def pivot_weekly_occupancy(weekly: pd.DataFrame) -> pd.DataFrame:

    df = weekly.copy()

    # Create sortable key
    df["week_start"] = pd.to_datetime(df["week_start"])

    # Create display label
    df["week_label"] = df["week_start"].dt.strftime("%d %b")

    # Sort BEFORE pivot
    df = df.sort_values("week_start")

    pivot = df.pivot_table(
        index=["listing_name"],
        columns="week_label",
        values="occupancy_rate",
        aggfunc="first"
    )

    # Reorder columns properly
    ordered_cols = (
        df[["week_start", "week_label"]]
        .drop_duplicates()
        .sort_values("week_start")["week_label"]
        .tolist()
    )

    pivot = pivot[ordered_cols]

    pivot = pivot.reset_index()

    # Format as %
    for col in ordered_cols:
        pivot[col] = (pivot[col] * 100).round(0).astype("Int64").astype(str) + "%"

    return pivot


def build_revenue_this_week(
    reservations: pd.DataFrame,
    daily_expenses: pd.DataFrame,
    today: datetime | None = None
) -> pd.DataFrame:

    df = reservations.copy()
    expenses = daily_expenses.copy()

    start_of_week, end_of_week = get_week_bounds(today)

    week = df[
        (df["checkin_date"] >= start_of_week) &
        (df["checkin_date"] <= end_of_week)
    ].copy()

    rows = []

    for _, row in week.iterrows():
        # Allocate only variable expenses to an individual reservation.
        # Fixed monthly costs such as mortgage or insurance are kept for
        # monthly/YTD profitability, not booking contribution.
        stay_expenses = expenses[
            (expenses["listing"].astype(str) == str(row["listing"])) &
            (expenses["date"] >= row["checkin_date"]) &
            (expenses["date"] < row["checkout_date"]) &
            (expenses["expense_type"] == "variable")
        ]

        estimated_variable_expenses = stay_expenses["daily_expense"].sum()

        gross_revenue = 0 if pd.isna(row.get("total_revenue")) else row.get("total_revenue")
        cleaning_fees = 0 if pd.isna(row.get("cleaning_fees")) else row.get("cleaning_fees")
        concierge_fees = 0 if pd.isna(row.get("concierge_commission")) else row.get("concierge_commission")

        revenue_net_before_tax = (
            gross_revenue
            - cleaning_fees
            - concierge_fees
            - estimated_variable_expenses
        )

         
        rows.append({
           # "checkin_date": row["checkin_date"],
            "Guest_Name": row.get("guest_name"),
           # "listing": row["listing"],
            "Listing_Name": row.get("listing_name"),
            "Nights": row["nights"],
            "Gross_Revenue": gross_revenue,
            "Cleaning_Fees": cleaning_fees,
            "Concierge_Fees": concierge_fees,
            "Variable_Expenses": estimated_variable_expenses,
            "Net_Before_Fixed_Costs": revenue_net_before_tax,
        })

    cols_to_round = [
        "Gross_Revenue",
        "Cleaning_Fees",
        "Concierge_Fees",
        "Variable_Expenses",
        "Net_Before_Fixed_Costs",
    ]

    result = pd.DataFrame(rows)
    result[cols_to_round] = result[cols_to_round].round(0).astype("Int64")

    total_row = {
    "Guest_Name": "TOTAL",
    "Listing_Name": "",
    "Nights": result["Nights"].sum(),
}

    for col in cols_to_round:
        total_row[col] = result[col].sum()

    result = pd.concat(
        [result, pd.DataFrame([total_row])],
        ignore_index=True
    )

    result["_sort"] = result["Listing_Name"].replace("", "ZZZ_TOTAL")
    result = result.sort_values("_sort").drop(columns="_sort")
    return result


# =============================================================================
# Email rendering and sending
# =============================================================================


def dataframe_to_html_table(
    df: pd.DataFrame,
    columns: list[str] | None = None,
    currency_columns: list[str] | None = None,
    color_percentages: bool = True,
) -> str:

    if df.empty:
        return "<p><em>No data available</em></p>"

    if columns is not None:
        df = df[columns]

    currency_columns = currency_columns or []

    table_style = "border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 13px;"
    th_style = "border: 1px solid #ddd; padding: 8px; background-color: #f3f4f6; text-align: left;"
    td_style = "border: 1px solid #ddd; padding: 8px;"

    html = f'<table style="{table_style}"><thead><tr>'

    for col in df.columns:
        label = col.replace("_", " ")
        html += f'<th style="{th_style}">{label}</th>'

    html += "</tr></thead><tbody>"

    for _, row in df.iterrows():
        is_total = str(row.iloc[0]).upper() == "TOTAL"
        row_style = "font-weight:bold; background-color:#f9fafb;" if is_total else ""

        html += f'<tr style="{row_style}">'

        for col in df.columns:
            value = row[col]
            cell_style = td_style

            if isinstance(value, pd.Timestamp):
                value = value.strftime("%a %d %b")

            if color_percentages and isinstance(value, str) and value.strip().endswith("%"):
                pct = int(value.strip().replace("%", ""))

                if pct == 0:
                    cell_style += " color:#dc2626; font-weight:bold;"
                elif pct < 30:
                    cell_style += " color:#f59e0b;"
                elif pct > 70:
                    cell_style += " color:#16a34a;"

            if col in currency_columns and pd.notna(value):
                try:
                    value = f"{float(value):,.0f} €"
                except (ValueError, TypeError):
                    pass

            current_month = datetime.today().strftime("%Y-%m")

            if (
                isinstance(col, str)
                and len(col) == 7
                and col[4] == "-"
                and col[:4].isdigit()
                and col[5:7].isdigit()
            ):
                if col > current_month:
                    cell_style += " opacity:0.5; font-style:italic;"
                    
            align = "right" if col in currency_columns or col == "Nights" else "left"

            # IMPORTANT: this line must be inside the for col loop
            html += f'<td style="{cell_style} text-align:{align};">{value}</td>'

        html += "</tr>"

    html += "</tbody></table>"

    return html


def build_weekly_email_html(summary: dict, weekly_occupancy_pivot: pd.DataFrame, revenue_this_week: pd.DataFrame, report_group_label: str = "All groups") -> str:
    start = summary["start_of_week"].strftime("%d %b")
    end = summary["end_of_week"].strftime("%d %b %Y")


    total_gross = revenue_this_week.loc[
        revenue_this_week["Guest_Name"] == "TOTAL",
        "Gross_Revenue"
    ].sum()

    total_net = revenue_this_week.loc[
        revenue_this_week["Guest_Name"] == "TOTAL",
        "Net_Before_Fixed_Costs"
    ].sum()

    arrivals_table = dataframe_to_html_table(
        summary["arrivals"],
        ["checkin_date", "guest_name", "listing_name", "nights", "booking_source"]
    )

    departures_table = dataframe_to_html_table(
        summary["departures"],
        ["checkout_date", "guest_name", "listing_name", "nights", "booking_source"]
    )

    occupancy_table = dataframe_to_html_table(weekly_occupancy_pivot)

    revenue_table = dataframe_to_html_table(
        revenue_this_week,
        currency_columns=[
            "Gross_Revenue",
            "Cleaning_Fees",
            "Concierge_Fees",
            "Other_Expenses",
            "Variable_Expenses",
            "Net_Before_Fixed_Costs",
        ],
    )
    table_style = f"""(
        "border-collapse: collapse; "
        "width: 100%; "
        "min-width: 760px; "
        "margin-top: 16px; "
        "font-family: Arial, sans-serif; "
        "font-size: 13px;"
        "border-spacing:12px;"
    )
    """
    #<table style="width:100%; margin-top:16px; border-spacing:12px;">

    kpi_cards = f"""
        <table role="presentation" style="width:100%; margin-top:16px; border-collapse:separate; border-spacing:8px;">
        <tr>
            <td style="width:25%; background:white; padding:14px; border-radius:8px;">
            <div style="font-size:12px; color:#666;">Arrivals</div>
            <div style="font-size:24px; font-weight:bold;">{summary["num_arrivals"]}</div>
            </td>

            <td style="width:25%; background:white; padding:14px; border-radius:8px;">
            <div style="font-size:12px; color:#666;">Departures</div>
            <div style="font-size:24px; font-weight:bold;">{summary["num_departures"]}</div>
            </td>

            <td style="width:25%; background:white; padding:14px; border-radius:8px;">
            <div style="font-size:12px; color:#666;">Gross revenue</div>
            <div style="font-size:24px; font-weight:bold;">{total_gross:,.0f} €</div>
            </td>

            <td style="width:25%; background:white; padding:14px; border-radius:8px;">
            <div style="font-size:12px; color:#666;">Net before tax</div>
            <div style="font-size:24px; font-weight:bold;">{total_net:,.0f} €</div>
            </td>
        </tr>
        </table>
        """

    html = f"""
    <html>
    <body style="margin:0; padding:0; background:#f6f7f9; font-family:Arial, sans-serif; color:#222;">

    <div style="max-width:900px; margin:0 auto; padding:24px;">

        <div style="background:#1f2937; color:white; padding:20px; border-radius:8px;">
        <h2 style="margin:0;">Weekly Reservations Summary</h2>
        <p style="margin:6px 0 0 0;">{start} – {end} — {report_group_label}</p>
        </div>

        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px;">
        <h3 style="margin-top:0;">Overview</h3>
        <p>
            <strong>Arrivals:</strong> {summary["num_arrivals"]}<br>
            <strong>Departures:</strong> {summary["num_departures"]}
        </p>
        </div>
        {kpi_cards}
        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px;">
        <h3 style="margin-top:0;">Arrivals</h3>
        {arrivals_table}
        </div>

        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px;">
        <h3 style="margin-top:0;">Departures</h3>
        {departures_table}
        </div>

        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px;">
        <h3 style="margin-top:0;">Revenue this week</h3>
        <div style="overflow-x:auto; width:100%;">
            {revenue_table}
        </div>
        </div>

        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px;">
        <h3 style="margin-top:0;">Occupancy outlook — next 8 weeks</h3>
        {occupancy_table}
        </div>

        <p style="font-size:12px; color:#777; margin-top:20px;">
        Generated automatically from the reservations workbook.
        </p>

    </div>
    </body>
    </html>
    """

    return html


def send_html_email(subject: str, html_body: str) -> None:
    """
    Send an HTML email using SMTP credentials from environment variables.

    This is portable:
    - local Mac: environment variables / .env
    - GitHub Actions: repository secrets
    """
    print("Looking for .env at:", BASE_DIR / ".env")

    print(BASE_DIR / ".env", (BASE_DIR / ".env").exists())

    print("Exists:", (BASE_DIR / ".env").exists())
    print("SMTP_HOST:", os.environ.get("SMTP_HOST"))
    smtp_host = os.environ["SMTP_HOST"]
    smtp_port = int(os.environ.get("SMTP_PORT", "465"))
    smtp_user = os.environ["SMTP_USER"]
    smtp_password = os.environ["SMTP_PASSWORD"]

    email_from = os.environ.get("EMAIL_FROM", smtp_user)
    email_to = os.environ["EMAIL_TO"]

    recipients = [email.strip() for email in email_to.split(",")]

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = email_from
    msg["To"] = ", ".join(recipients)

    msg.set_content("This email requires an HTML-compatible email client.")
    msg.add_alternative(html_body, subtype="html")

    with smtplib.SMTP_SSL(smtp_host, smtp_port) as server:
        server.login(smtp_user, smtp_password)
        server.send_message(msg)




def build_monthly_listing_pivot(
    df: pd.DataFrame,
    value_col: str,
    listing_col: str = "listing_name",
    month_col: str = "year_month",
) -> pd.DataFrame:
    pivot = df.pivot_table(
        index=listing_col,
        columns=month_col,
        values=value_col,
        aggfunc="sum",
        fill_value=0,
    )

    pivot["Total"] = pivot.sum(axis=1)

    total_row = pivot.sum(axis=0)
    total_row.name = "TOTAL"

    pivot = pd.concat([pivot, total_row.to_frame().T])
    pivot.index.name = "listing_name"
    pivot = pivot.reset_index()

    return pivot



def build_monthly_finance_base(
    booking_nights: pd.DataFrame,
    monthly_expenses: pd.DataFrame,
    listings: pd.DataFrame,
) -> pd.DataFrame:
    """
    Build monthly finance metrics by listing.

    Finance view logic:
    - revenue comes from bookings
    - cleaning/concierge come from bookings
    - variable expenses use the selected full monthly amount
    - fixed expenses use the selected full monthly amount
    - every listing/month exists even if there are no bookings
    """

    bookings = booking_nights.copy()
    expenses = monthly_expenses.copy()
    listings = listings.copy()

    bookings["year_month"] = bookings["date"].dt.strftime("%Y-%m")
    expenses["year_month"] = (
        expenses["year"].astype(int).astype(str)
        + "-"
        + expenses["month"].astype(int).astype(str).str.zfill(2)
    )

    # Revenue / fees from booking nights
    revenue_monthly = (
        bookings.groupby(["listing", "year_month"], as_index=False)
        .agg(
            Gross_Revenue=("daily_revenue_total", "sum"),
            Cleaning_Fees=("daily_cleaning_fees", "sum"),
            Concierge_Fees=("daily_concierge_commission", "sum"),
        )
    )

    # Average Daily Rate

    ADR_monthly = (
        bookings.groupby(["listing", "year_month"], as_index=False)
        .agg(
            adr_revenue=("daily_revenue_total", "sum"),
            Occupied_Nights=("occupied", "sum"),
        )
    )

    ADR_monthly["ADR"] = (
        ADR_monthly["adr_revenue"] / ADR_monthly["Occupied_Nights"].replace(0, pd.NA)
    )

    # Full monthly variable expenses for finance view
    variable_monthly = (
        expenses[expenses["expense_type"] == "variable"]
        .groupby(["listing", "year_month"], as_index=False)
        .agg(Variable_Expenses=("allocated_monthly_expense", "sum"))
    )

    # Full monthly fixed expenses
    fixed_monthly = (
        expenses[expenses["expense_type"] == "fixed"]
        .groupby(["listing", "year_month"], as_index=False)
        .agg(Fixed_Expenses=("allocated_monthly_expense", "sum"))
    )

    # Build complete listing x month grid from listings and expense/revenue months
    all_listings = listings["listing"].dropna().apply(normalise_listing_key).unique()

    all_months = sorted(
        set(revenue_monthly["year_month"].dropna())
        | set(expenses["year_month"].dropna())
    )

    full_grid = pd.MultiIndex.from_product(
        [all_listings, all_months],
        names=["listing", "year_month"]
    ).to_frame(index=False)

    finance = full_grid.merge(
        revenue_monthly,
        on=["listing", "year_month"],
        how="left"
    )

    finance = finance.merge(
        variable_monthly,
        on=["listing", "year_month"],
        how="left"
    )

    finance = finance.merge(
        fixed_monthly,
        on=["listing", "year_month"],
        how="left"
    )

    finance = finance.merge(
        ADR_monthly,
        on=["listing", "year_month"],
        how="left"
    )

    money_cols = [
        "Gross_Revenue",
        "Cleaning_Fees",
        "Concierge_Fees",
        "Variable_Expenses",
        "Fixed_Expenses",
        "ADR",
        "Occupied_Nights",
    ]

    finance[money_cols] = finance[money_cols].fillna(0)

    finance["Contribution_Before_Fixed_Costs"] = (
        finance["Gross_Revenue"]
        - finance["Cleaning_Fees"]
        - finance["Concierge_Fees"]
        - finance["Variable_Expenses"]
    )

    finance["Net_After_Fixed_Costs"] = (
        finance["Contribution_Before_Fixed_Costs"]
        - finance["Fixed_Expenses"]
    )

    finance = finance.merge(
        listings[["listing", "listing_name"]],
        on="listing",
        how="left"
    )

    finance["listing_name"] = finance["listing_name"].fillna(finance["listing"])

    return finance


def build_monthly_occupancy_base(
    daily: pd.DataFrame,
    listings: pd.DataFrame,
) -> pd.DataFrame:
    df = daily.copy()

    df["year_month"] = df["date"].dt.strftime("%Y-%m")

    monthly = (
        df.groupby(["listing", "year_month"], as_index=False)
        .agg(
            occupied_nights=("occupied", "sum"),
            available_nights=("date", "count"),
        )
    )

    monthly["Occupancy_Rate"] = (
        monthly["occupied_nights"] / monthly["available_nights"]
    )

    monthly = monthly.merge(
        listings[["listing", "listing_name"]],
        on="listing",
        how="left",
    )

    monthly["listing_name"] = monthly["listing_name"].fillna(monthly["listing"])

    return monthly






def build_cleaning_pct_net_pivot(monthly_finance: pd.DataFrame) -> pd.DataFrame:
    """
    Build Cleaning Fees as a % of net revenue by listing and month.

    Net revenue here means gross revenue after concierge, before cleaning.
    This KPI highlights listings/months where cleaning is consuming too much
    of the revenue, usually because stays are too short or prices are too low.
    """
    df = monthly_finance.copy()

    df["Net_Revenue_Before_Cleaning"] = (
        df["Gross_Revenue"] - df["Concierge_Fees"]
    )

    df["Cleaning_Pct_Net"] = (
        df["Cleaning_Fees"]
        / df["Net_Revenue_Before_Cleaning"].replace(0, pd.NA)
    )

    pivot = df.pivot_table(
        index="listing_name",
        columns="year_month",
        values="Cleaning_Pct_Net",
        aggfunc="first",
        fill_value=0,
    )

    pivot.index.name = "listing_name"
    pivot = pivot.reset_index()

    month_cols = [col for col in pivot.columns if col != "listing_name"]
    for col in month_cols:
        pivot[col] = (
            pivot[col] * 100
        ).round(0).astype("Int64").astype(str) + "%"

    return pivot


def build_listing_ytd_net_chart(
    daily_for_report: pd.DataFrame,
    daily_expenses_for_report: pd.DataFrame,
    listings: pd.DataFrame,
    output_dir: Path,
    filename_suffix: str = "all-groups",
    today: datetime | None = None,
) -> Path:
    """
    Build a line chart with one line per listing showing cumulative net revenue.

    This is not pure YTD:
    - solid line = actuals up to yesterday
    - dashed line = future on-the-books values
    - vertical line = today
    """
    if today is None:
        today = datetime.today()

    today_ts = pd.Timestamp(today).normalize()

    daily = daily_for_report.copy()
    expenses = daily_expenses_for_report.copy()
    listings = listings.copy()

    daily["date"] = pd.to_datetime(daily["date"])
    expenses["date"] = pd.to_datetime(expenses["date"])
    daily["listing"] = daily["listing"].apply(normalise_listing_key)
    expenses["listing"] = expenses["listing"].apply(normalise_listing_key)
    listings["listing"] = listings["listing"].apply(normalise_listing_key)

    revenue_daily = (
        daily.groupby(["listing", "date"], as_index=False)
        .agg(
            gross_revenue=("daily_revenue_total", "sum"),
            cleaning_fees=("daily_cleaning_fees", "sum"),
            concierge_fees=("daily_concierge_commission", "sum"),
        )
    )

    if expenses.empty:
        expenses_daily = pd.DataFrame(columns=["listing", "date", "variable", "fixed"])
    else:
        expenses_daily = (
            expenses.groupby(["listing", "date", "expense_type"], as_index=False)
            .agg(expense=("daily_expense", "sum"))
            .pivot_table(
                index=["listing", "date"],
                columns="expense_type",
                values="expense",
                aggfunc="sum",
                fill_value=0,
            )
            .reset_index()
        )

    if "variable" not in expenses_daily.columns:
        expenses_daily["variable"] = 0
    if "fixed" not in expenses_daily.columns:
        expenses_daily["fixed"] = 0

    chart_df = revenue_daily.merge(
        expenses_daily[["listing", "date", "variable", "fixed"]],
        on=["listing", "date"],
        how="left",
    )

    chart_df[["variable", "fixed"]] = chart_df[["variable", "fixed"]].fillna(0)

    chart_df["net_after_fixed"] = (
        chart_df["gross_revenue"]
        - chart_df["cleaning_fees"]
        - chart_df["concierge_fees"]
        - chart_df["variable"]
        - chart_df["fixed"]
    )

    chart_df = chart_df.merge(
        listings[["listing", "listing_name"]],
        on="listing",
        how="left",
    )

    chart_df["listing_name"] = chart_df["listing_name"].fillna(chart_df["listing"])
    chart_df = chart_df.sort_values(["listing_name", "date"])
    chart_df["cumulative_net_after_fixed"] = (
        chart_df.groupby("listing_name")["net_after_fixed"].cumsum()
    )

    chart_path = output_dir / f"listing_net_revenue_actual_otb_{filename_suffix}.png"

    plt.figure(figsize=(10, 5))

    for listing_name, group in chart_df.groupby("listing_name"):
        actual = group[group["date"] < today_ts]
        future = group[group["date"] >= today_ts]

        if not actual.empty:
            line = plt.plot(
                actual["date"],
                actual["cumulative_net_after_fixed"],
                label=str(listing_name),
            )
            line_color = line[0].get_color()
        else:
            line = plt.plot([], [], label=str(listing_name))
            line_color = line[0].get_color()

        if not future.empty:
            if not actual.empty:
                bridge = pd.concat([actual.tail(1), future], ignore_index=True)
            else:
                bridge = future

            plt.plot(
                bridge["date"],
                bridge["cumulative_net_after_fixed"],
                linestyle="--",
                color=line_color,
            )

    plt.axvline(today_ts, linestyle="--", linewidth=1)
    plt.title("Net Revenue by Listing — Actual + On the Books")
    plt.xlabel("")
    plt.ylabel("€")
    plt.legend(fontsize=8)
    plt.grid(True, alpha=0.3)
    plt.tight_layout()

    plt.savefig(chart_path, dpi=150)
    plt.close()

    return chart_path


def build_monthly_occupancy_chart(
    monthly_occupancy: pd.DataFrame,
    output_dir: Path,
    filename_suffix: str = "all-groups",
    today: datetime | None = None,
) -> Path:
    """
    Build a monthly occupancy chart with one line per listing.

    Solid line = completed months before the current month.
    Dashed line = current/future on-the-books months.
    Vertical line = today.
    """
    if today is None:
        today = datetime.today()

    today_ts = pd.Timestamp(today).normalize()
    current_month_start = today_ts.replace(day=1)

    df = monthly_occupancy.copy()
    df["month_start"] = pd.to_datetime(df["year_month"] + "-01")
    df = df.sort_values(["listing_name", "month_start"])

    chart_path = output_dir / f"occupancy_by_listing_actual_otb_{filename_suffix}.png"

    plt.figure(figsize=(10, 5))

    for listing_name, group in df.groupby("listing_name"):
        actual = group[group["month_start"] < current_month_start]
        future = group[group["month_start"] >= current_month_start]

        if not actual.empty:
            line = plt.plot(
                actual["month_start"],
                actual["Occupancy_Rate"] * 100,
                label=str(listing_name),
            )
            line_color = line[0].get_color()
        else:
            line = plt.plot([], [], label=str(listing_name))
            line_color = line[0].get_color()

        if not future.empty:
            if not actual.empty:
                bridge = pd.concat([actual.tail(1), future], ignore_index=True)
            else:
                bridge = future

            plt.plot(
                bridge["month_start"],
                bridge["Occupancy_Rate"] * 100,
                linestyle="--",
                color=line_color,
            )

    plt.axvline(today_ts, linestyle="--", linewidth=1)
    plt.title("Occupancy by Listing — Actual + On the Books")
    plt.xlabel("")
    plt.ylabel("Occupancy %")
    plt.ylim(0, 105)
    plt.legend(fontsize=8)
    plt.grid(True, alpha=0.3)
    plt.tight_layout()

    plt.savefig(chart_path, dpi=150)
    plt.close()

    return chart_path


def _format_kpi_value(value, value_type: str) -> str:
    if pd.isna(value):
        return "—"
    if value_type == "currency":
        return f"{value:,.0f} €"
    if value_type == "currency_per_night":
        return f"{value:,.0f} €/night"
    if value_type == "percent":
        return f"{value * 100:,.0f}%"
    return str(value)


def _kpi_color(value, good_when_positive: bool = True) -> str:
    if pd.isna(value):
        return "#111827"
    if good_when_positive:
        return "#16a34a" if value >= 0 else "#dc2626"
    return "#dc2626" if value >= 0 else "#16a34a"


def _calculate_period_metrics(
    daily: pd.DataFrame,
    expenses: pd.DataFrame,
) -> dict:
    """
    Calculate profitability metrics for a given daily slice.

    The caller decides the period:
    - YTD actuals: dates before today
    - On the books: dates from today onward
    """
    daily = daily.copy()
    expenses = expenses.copy()

    gross = daily["daily_revenue_total"].sum()
    cleaning = daily["daily_cleaning_fees"].sum()
    concierge = daily["daily_concierge_commission"].sum()

    variable_expenses = expenses.loc[
        expenses["expense_type"] == "variable",
        "daily_expense"
    ].sum()

    fixed_expenses = expenses.loc[
        expenses["expense_type"] == "fixed",
        "daily_expense"
    ].sum()

    occupied_nights = daily["occupied"].sum()
    available_nights = len(daily)

    contribution = gross - cleaning - concierge - variable_expenses
    net_after_fixed = contribution - fixed_expenses

    net_revenue_before_cleaning = gross - concierge
    operating_costs = cleaning + concierge + variable_expenses

    return {
        "Gross_Revenue": gross,
        "Cleaning_Fees": cleaning,
        "Concierge_Fees": concierge,
        "Variable_Expenses": variable_expenses,
        "Fixed_Expenses": fixed_expenses,
        "Operating_Costs": operating_costs,
        "Contribution_Before_Fixed_Costs": contribution,
        "Net_After_Fixed_Costs": net_after_fixed,
        "Net_Revenue_Before_Cleaning": net_revenue_before_cleaning,
        "Occupied_Nights": occupied_nights,
        "Available_Nights": available_nights,
        "ADR": gross / occupied_nights if occupied_nights else pd.NA,
        "Occupancy_Rate": occupied_nights / available_nights if available_nights else pd.NA,
        "Variable_Cost_Per_Night": operating_costs / occupied_nights if occupied_nights else pd.NA,
        "Fixed_Cost_Per_Night": fixed_expenses / occupied_nights if occupied_nights else pd.NA,
        "Net_Per_Night": net_after_fixed / occupied_nights if occupied_nights else pd.NA,
        "Cleaning_Pct_Net": cleaning / net_revenue_before_cleaning if net_revenue_before_cleaning else pd.NA,
    }


def _split_ytd_and_otb(
    daily_for_report: pd.DataFrame,
    daily_expenses_for_report: pd.DataFrame,
    today: datetime | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Split daily data into:
    - YTD actuals: dates strictly before today
    - On the books: dates from today onward
    """
    if today is None:
        today = datetime.today()

    today_ts = pd.Timestamp(today).normalize()

    daily = daily_for_report.copy()
    expenses = daily_expenses_for_report.copy()

    daily["date"] = pd.to_datetime(daily["date"])
    expenses["date"] = pd.to_datetime(expenses["date"])

    daily_ytd = daily[daily["date"] < today_ts].copy()
    daily_otb = daily[daily["date"] >= today_ts].copy()
    expenses_ytd = expenses[expenses["date"] < today_ts].copy()
    expenses_otb = expenses[expenses["date"] >= today_ts].copy()

    return daily_ytd, daily_otb, expenses_ytd, expenses_otb


def build_finance_kpis(
    daily_for_report: pd.DataFrame,
    daily_expenses_for_report: pd.DataFrame,
    today: datetime | None = None,
) -> list[list[dict]]:
    """
    Headline finance KPI cards for the selected report group.

    Layout:
    1. headline outcome: YTD, OTB, full-year landing, occupancy
    2. revenue and cleaning pressure
    3. profitability funnel per occupied night

    YTD = dates before today.
    OTB = dates from today onward, confirmed bookings only.
    Full Year = YTD + OTB.
    """
    daily_ytd, daily_otb, expenses_ytd, expenses_otb = _split_ytd_and_otb(
        daily_for_report,
        daily_expenses_for_report,
        today=today,
    )

    ytd = _calculate_period_metrics(daily_ytd, expenses_ytd)
    otb = _calculate_period_metrics(daily_otb, expenses_otb)

    full_year = {
        "Gross_Revenue": ytd["Gross_Revenue"] + otb["Gross_Revenue"],
        "Cleaning_Fees": ytd["Cleaning_Fees"] + otb["Cleaning_Fees"],
        "Net_After_Fixed_Costs": ytd["Net_After_Fixed_Costs"] + otb["Net_After_Fixed_Costs"],
    }

    variable_cost_warning = (
        pd.notna(ytd["Variable_Cost_Per_Night"])
        and pd.notna(ytd["ADR"])
        and ytd["Variable_Cost_Per_Night"] >= ytd["ADR"] * 0.35
    )

    fixed_cost_warning = (
        pd.notna(ytd["Fixed_Cost_Per_Night"])
        and pd.notna(ytd["ADR"])
        and ytd["Fixed_Cost_Per_Night"] >= ytd["ADR"] * 0.35
    )

    return [
        [
            {
                "label": "Net Profit Full Year",
                "display": _format_kpi_value(full_year["Net_After_Fixed_Costs"], "currency"),
                "color": _kpi_color(full_year["Net_After_Fixed_Costs"]),
                "note": "YTD + on the books",
            },
            {
                "label": "Net Profit YTD",
                "display": _format_kpi_value(ytd["Net_After_Fixed_Costs"], "currency"),
                "color": _kpi_color(ytd["Net_After_Fixed_Costs"]),
                "note": "actuals to yesterday",
            },
            {
                "label": "Net Profit rest of year",
                "display": _format_kpi_value(otb["Net_After_Fixed_Costs"], "currency"),
                "color": _kpi_color(otb["Net_After_Fixed_Costs"]),
                "note": "future confirmed bookings",
            },
            {
                "label": "Occupancy YTD / OTB",
                "display": (
                    f'{_format_kpi_value(ytd["Occupancy_Rate"], "percent")} / '
                    f'{_format_kpi_value(otb["Occupancy_Rate"], "percent")}'
                ),
                "color": "#111827",
                "note": "actual / on the books",
            },
        ],
        [
            {
                "label": "Gross Revenue Full Year",
                "display": _format_kpi_value(full_year["Gross_Revenue"], "currency"),
                "color": "#111827",
                "note": "YTD + on the books",
            },
            {
                "label": "Gross Revenue to date",
                "display": _format_kpi_value(ytd["Gross_Revenue"], "currency"),
                "color": "#111827",
                "note": "actual gross revenue",
            },
            {
                "label": "Gross Revenue rest of year",
                "display": _format_kpi_value(otb["Gross_Revenue"], "currency"),
                "color": "#111827",
                "note": "future booked gross revenue",
            },


            {
                "label": "Cleaning Fees Full Year",
                "display": _format_kpi_value(full_year["Cleaning_Fees"], "currency"),
                "color": "#111827",
                "note": "YTD + on the books",
            },
        ],
        [
            {
                "label": "ADR YTD",
                "display": _format_kpi_value(ytd["ADR"], "currency"),
                "color": "#111827",
                "note": "gross / occupied night",
            },
            {
                "label": "Variable Costs €/night",
                "display": _format_kpi_value(ytd["Variable_Cost_Per_Night"], "currency_per_night"),
                "color": "#dc2626" if variable_cost_warning else "#111827",
                "note": "cleaning + concierge + energy*",
            },
            {
                "label": "Fixed Costs €/night",
                "display": _format_kpi_value(ytd["Fixed_Cost_Per_Night"], "currency_per_night"),
                "color": "#dc2626" if fixed_cost_warning else "#111827",
                "note": "fixed costs / occupied night",
            },
            {
                "label": "Net €/night",
                "display": _format_kpi_value(ytd["Net_Per_Night"], "currency_per_night"),
                "color": _kpi_color(ytd["Net_Per_Night"]),
                "note": "after variable + fixed costs",
            },
        ],
    ]


def build_apartment_kpi_table(
    daily_for_report: pd.DataFrame,
    daily_expenses_for_report: pd.DataFrame,
    listings: pd.DataFrame,
    today: datetime | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Build two diagnostic tables by apartment/listing.

    Table 1: performance by listing.
    Table 2: funnel by listing.
    """
    daily_ytd, daily_otb, expenses_ytd, expenses_otb = _split_ytd_and_otb(
        daily_for_report,
        daily_expenses_for_report,
        today=today,
    )

    listings = listings.copy()
    listings["listing"] = listings["listing"].apply(normalise_listing_key)

    performance_rows = []
    funnel_rows = []

    for _, listing_row in listings.sort_values("listing_name").iterrows():
        listing = normalise_listing_key(listing_row["listing"])
        listing_name = listing_row["listing_name"]

        ytd = _calculate_period_metrics(
            daily_ytd[daily_ytd["listing"].apply(normalise_listing_key) == listing],
            expenses_ytd[expenses_ytd["listing"].apply(normalise_listing_key) == listing],
        )

        otb = _calculate_period_metrics(
            daily_otb[daily_otb["listing"].apply(normalise_listing_key) == listing],
            expenses_otb[expenses_otb["listing"].apply(normalise_listing_key) == listing],
        )

        performance_rows.append({
            "Listing": listing_name,
            "Net Profit YTD": ytd["Net_After_Fixed_Costs"],
            "Net Profit on the Books": otb["Net_After_Fixed_Costs"],
            "Gross YTD": ytd["Gross_Revenue"],
            "Gross on the Books": otb["Gross_Revenue"],
            "Occupancy YTD": ytd["Occupancy_Rate"],
            "Occupancy on the Books": otb["Occupancy_Rate"],
        })

        funnel_rows.append({
            "Listing": listing_name,
            "ADR YTD": ytd["ADR"],
            "Variable Costs €/night": ytd["Variable_Cost_Per_Night"],
            "Cleaning % net": ytd["Cleaning_Pct_Net"],
            "Fixed Costs €/night": ytd["Fixed_Cost_Per_Night"],
            "Net €/night": ytd["Net_Per_Night"],
        })

    performance_table = pd.DataFrame(performance_rows)
    funnel_table = pd.DataFrame(funnel_rows)

    performance_currency_cols = [
        "Net Profit YTD",
        "Net Profit on the Books",
        "Gross YTD",
        "Gross on the Books",
    ]

    performance_percent_cols = [
        "Occupancy YTD",
        "Occupancy on the Books",
    ]

    funnel_currency_cols = [
        "ADR YTD",
        "Variable Costs €/night",
        "Fixed Costs €/night",
        "Net €/night",
    ]

    funnel_percent_cols = ["Cleaning % net"]

    for col in performance_currency_cols:
        performance_table[col] = performance_table[col].apply(
            lambda value: "—" if pd.isna(value) else f"{value:,.0f} €"
        )

    for col in performance_percent_cols:
        performance_table[col] = performance_table[col].apply(
            lambda value: "—" if pd.isna(value) else f"{value * 100:,.0f}%"
        )

    for col in funnel_currency_cols:
        funnel_table[col] = funnel_table[col].apply(
            lambda value: "—" if pd.isna(value) else f"{value:,.0f} €"
        )

    for col in funnel_percent_cols:
        funnel_table[col] = funnel_table[col].apply(
            lambda value: "—" if pd.isna(value) else f"{value * 100:,.0f}%"
        )

    return performance_table, funnel_table


def render_finance_kpi_cards(kpi_rows: list[list[dict]]) -> str:
    """Render KPI rows. Each inner list becomes one row of equal-width cards."""
    rows_html = []

    for row in kpi_rows:
        width = 100 / max(len(row), 1)
        cells = []

        for item in row:
            cells.append(
                f'<td style="width:{width:.2f}%; background:white; padding:14px; border-radius:8px; border:1px solid #e5e7eb; vertical-align:top;">'
                f'<div style="font-size:12px; color:#666;">{item["label"]}</div>'
                f'<div style="font-size:24px; font-weight:bold; color:{item["color"]}; margin-top:4px;">{item["display"]}</div>'
                f'<div style="font-size:11px; color:#777; margin-top:4px;">{item["note"]}</div>'
                f'</td>'
            )

        rows_html.append("<tr>" + "".join(cells) + "</tr>")

    return (
        '<table role="presentation" style="width:100%; margin-top:16px; border-collapse:separate; border-spacing:8px;">'
        + "".join(rows_html)
        + "</table>"
    )




def _html_escape(value) -> str:
    """Minimal HTML escaping for model output and error messages."""
    text = "" if value is None else str(value)
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _safe_number(value):
    """Return JSON-safe rounded numbers; convert NaN/NA to None."""
    if pd.isna(value):
        return None
    try:
        return round(float(value), 4)
    except (TypeError, ValueError):
        return value


def classify_rental_model(avg_stay_length: float | None, max_stay_length: float | None) -> str:
    """
    Simple heuristic to help the AI distinguish short-stay vs medium/long-stay activity.
    """
    if avg_stay_length is None or pd.isna(avg_stay_length):
        return "unknown"

    if max_stay_length is not None and max_stay_length >= 90:
        return "medium_or_long_term"

    if avg_stay_length >= 30:
        return "medium_or_long_term"

    if avg_stay_length >= 14:
        return "extended_short_stay"

    return "short_stay"


def build_listing_booking_profile(
    reservations: pd.DataFrame,
    listings: pd.DataFrame,
) -> list[dict]:

    df = reservations.copy()
    listings = listings.copy()

    df["listing"] = df["listing"].apply(normalise_listing_key)
    listings["listing"] = listings["listing"].apply(normalise_listing_key)

    # Avoid duplicate listing_name columns if reservations is already enriched
    if "listing_name" not in df.columns:
        df = df.merge(
            listings[["listing", "listing_name"]],
            on="listing",
            how="left",
        )

    df["listing_name"] = df["listing_name"].fillna(df["listing"])

    rows = []

    for listing_name, group in df.groupby("listing_name"):
        avg_stay = group["nights"].mean()
        max_stay = group["nights"].max()

        rows.append({
            "listing": listing_name,
            "bookings_count": int(len(group)),
            "average_stay_length": _safe_number(avg_stay),
            "max_stay_length": _safe_number(max_stay),
            "rental_model": classify_rental_model(avg_stay, max_stay),
        })

    return rows

def build_monthly_listing_performance(
    daily_for_report: pd.DataFrame,
    daily_expenses_for_report: pd.DataFrame,
    listings: pd.DataFrame,
    today: datetime | None = None,
) -> list[dict]:
    """
    Build compact monthly performance by listing for AI analysis.

    This gives the model the shape of performance over time without sending
    raw daily data.
    """
    if today is None:
        today = datetime.today()

    today_ts = pd.Timestamp(today).normalize()

    daily = daily_for_report.copy()
    expenses = daily_expenses_for_report.copy()
    listings = listings.copy()

    daily["date"] = pd.to_datetime(daily["date"])
    expenses["date"] = pd.to_datetime(expenses["date"])

    daily["listing"] = daily["listing"].apply(normalise_listing_key)
    expenses["listing"] = expenses["listing"].apply(normalise_listing_key)
    listings["listing"] = listings["listing"].apply(normalise_listing_key)

    daily["year_month"] = daily["date"].dt.strftime("%Y-%m")
    expenses["year_month"] = expenses["date"].dt.strftime("%Y-%m")

    revenue_monthly = (
        daily.groupby(["listing", "year_month"], as_index=False)
        .agg(
            available_nights=("date", "count"),
            booked_nights=("occupied", "sum"),
            gross_revenue=("daily_revenue_total", "sum"),
            cleaning_fees=("daily_cleaning_fees", "sum"),
            concierge_fees=("daily_concierge_commission", "sum"),
        )
    )

    expenses_monthly = (
        expenses.groupby(["listing", "year_month", "expense_type"], as_index=False)
        .agg(expense=("daily_expense", "sum"))
        .pivot_table(
            index=["listing", "year_month"],
            columns="expense_type",
            values="expense",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )

    if "variable" not in expenses_monthly.columns:
        expenses_monthly["variable"] = 0

    if "fixed" not in expenses_monthly.columns:
        expenses_monthly["fixed"] = 0

    monthly = revenue_monthly.merge(
        expenses_monthly[["listing", "year_month", "variable", "fixed"]],
        on=["listing", "year_month"],
        how="left",
    )

    monthly[["variable", "fixed"]] = monthly[["variable", "fixed"]].fillna(0)

    monthly["occupancy"] = (
        monthly["booked_nights"] / monthly["available_nights"].replace(0, pd.NA)
    )

    monthly["adr"] = (
        monthly["gross_revenue"] / monthly["booked_nights"].replace(0, pd.NA)
    )

    monthly["contribution_before_fixed"] = (
        monthly["gross_revenue"]
        - monthly["cleaning_fees"]
        - monthly["concierge_fees"]
        - monthly["variable"]
    )

    monthly["net_after_fixed"] = (
        monthly["contribution_before_fixed"]
        - monthly["fixed"]
    )

    monthly["period_type"] = monthly["year_month"].apply(
        lambda ym: (
            "past_or_current"
            if pd.Timestamp(f"{ym}-01") <= today_ts.replace(day=1)
            else "future_on_the_books"
        )
    )

    monthly = monthly.merge(
        listings[["listing", "listing_name"]],
        on="listing",
        how="left",
    )

    monthly["listing_name"] = monthly["listing_name"].fillna(monthly["listing"])

    keep_cols = [
        "listing_name",
        "year_month",
        "period_type",
        "available_nights",
        "booked_nights",
        "occupancy",
        "gross_revenue",
        "adr",
        "cleaning_fees",
        "concierge_fees",
        "variable",
        "contribution_before_fixed",
        "net_after_fixed",
    ]

    monthly = monthly[keep_cols].sort_values(["listing_name", "year_month"])

    records = []
    for _, row in monthly.iterrows():
        records.append({
            "listing": row["listing_name"],
            "month": row["year_month"],
            "period_type": row["period_type"],
            "available_nights": int(row["available_nights"]),
            "booked_nights": int(row["booked_nights"]),
            "occupancy": _safe_number(row["occupancy"]),
            "gross_revenue": _safe_number(row["gross_revenue"]),
            "average_nightly_rate": _safe_number(row["adr"]),
            "cleaning_fees": _safe_number(row["cleaning_fees"]),
            "concierge_fees": _safe_number(row["concierge_fees"]),
            "variable_costs": _safe_number(row["variable"]),
            "contribution_before_fixed": _safe_number(row["contribution_before_fixed"]),
            "net_after_fixed": _safe_number(row["net_after_fixed"]),
        })

    return records


def build_listing_signals(
    monthly_listing_performance: list[dict],
    booking_profiles: list[dict],
) -> list[dict]:
    """
    Pre-compute simple diagnostic signals to guide the AI.

    This helps avoid the AI over-focusing on fixed costs.
    """
    performance = pd.DataFrame(monthly_listing_performance)
    profiles = pd.DataFrame(booking_profiles)

    if performance.empty:
        return []

    signals = []

    for listing, group in performance.groupby("listing"):
        future = group[group["period_type"] == "future_on_the_books"]
        past = group[group["period_type"] == "past_or_current"]

        avg_future_occupancy = future["occupancy"].dropna().mean() if not future.empty else pd.NA
        avg_past_occupancy = past["occupancy"].dropna().mean() if not past.empty else pd.NA

        summer = group[group["month"].astype(str).str[5:7].isin(["06", "07", "08"])]
        summer_occupancy = summer["occupancy"].dropna().mean() if not summer.empty else pd.NA

        profile_row = profiles[profiles["listing"] == listing]
        rental_model = (
            profile_row["rental_model"].iloc[0]
            if not profile_row.empty
            else "unknown"
        )

        listing_signals = []

        if pd.notna(avg_future_occupancy) and avg_future_occupancy < 0.25:
            listing_signals.append("low_future_bookings")

        if pd.notna(summer_occupancy) and summer_occupancy < 0.40:
            listing_signals.append("weak_peak_season_bookings")

        if rental_model == "medium_or_long_term":
            listing_signals.append("medium_or_long_term_pattern")

        if pd.notna(avg_past_occupancy) and pd.notna(avg_future_occupancy):
            if avg_future_occupancy < avg_past_occupancy * 0.6:
                listing_signals.append("future_bookings_weaker_than_past")

        signals.append({
            "listing": listing,
            "rental_model": rental_model,
            "average_past_occupancy": _safe_number(avg_past_occupancy),
            "average_future_occupancy": _safe_number(avg_future_occupancy),
            "summer_occupancy_on_books": _safe_number(summer_occupancy),
            "signals": listing_signals,
        })

    return signals





def build_ai_analysis_payload(
    daily_for_report: pd.DataFrame,
    daily_expenses_for_report: pd.DataFrame,
    listings: pd.DataFrame,
    report_group_label: str,
    reservations_for_report: pd.DataFrame | None = None,
    monthly_targets: pd.DataFrame | None = None,
    today: datetime | None = None,
) -> dict:
    """
    Build a compact, deterministic payload for GenAI analysis.

    The model receives Python-computed metrics only. It should explain and
    prioritise, not calculate.
    """
    daily_ytd, daily_otb, expenses_ytd, expenses_otb = _split_ytd_and_otb(
        daily_for_report,
        daily_expenses_for_report,
        today=today,
    )

    def clean_metrics(metrics: dict) -> dict:
        return {key: _safe_number(value) for key, value in metrics.items()}

    ytd = _calculate_period_metrics(daily_ytd, expenses_ytd)
    otb = _calculate_period_metrics(daily_otb, expenses_otb)

    listings = listings.copy()
    listings["listing"] = listings["listing"].apply(normalise_listing_key)

    listing_rows = []

    for _, listing_row in listings.sort_values("listing_name").iterrows():
        listing = normalise_listing_key(listing_row["listing"])
        listing_name = listing_row["listing_name"]

        listing_ytd = _calculate_period_metrics(
            daily_ytd[daily_ytd["listing"].apply(normalise_listing_key) == listing],
            expenses_ytd[expenses_ytd["listing"].apply(normalise_listing_key) == listing],
        )

        listing_otb = _calculate_period_metrics(
            daily_otb[daily_otb["listing"].apply(normalise_listing_key) == listing],
            expenses_otb[expenses_otb["listing"].apply(normalise_listing_key) == listing],
        )

        listing_rows.append({
            "listing": listing_name,
            "YTD": clean_metrics(listing_ytd),
            "OTB": clean_metrics(listing_otb),
        })

    monthly_listing_performance = build_monthly_listing_performance(
        daily_for_report,
        daily_expenses_for_report,
        listings,
        today=today,
    )

    monthly_vs_target = (
        build_monthly_vs_target_payload(
            monthly_listing_performance,
            monthly_targets,
            listings,
        )
        if monthly_targets is not None
        else []
    )

    booking_profiles = (
        build_listing_booking_profile(reservations_for_report, listings)
        if reservations_for_report is not None
        else []
    )

    listing_signals = build_listing_signals(
        monthly_listing_performance,
        booking_profiles,
    )

    target_payload = (
    build_monthly_target_payload(monthly_targets, listings)
    if monthly_targets is not None
    else []
)

    return {
        "report_group": report_group_label,
        "definitions": {
            "YTD": "actual dates before today",
            "OTB": "future confirmed bookings from today onward",
            "Full Year": "YTD actuals plus OTB confirmed future bookings",
            "ADR": "gross revenue divided by occupied nights",
            "Variable_Cost_Per_Night": "cleaning + concierge + energy/usage-based variable expenses divided by occupied nights",
            "Fixed_Cost_Per_Night": "fixed expenses divided by occupied nights",
            "Cleaning_Pct_Net": "cleaning fees divided by gross revenue minus concierge fees",
            "Net_Per_Night": "net after variable and fixed costs divided by occupied nights",
            "monthly_targets": target_payload,
            "monthly_vs_target": monthly_vs_target,
        },
        "headline_metrics": {
            "YTD": clean_metrics(ytd),
            "OTB": clean_metrics(otb),
            "Full_Year": {
                "Gross_Revenue": _safe_number(
                    ytd["Gross_Revenue"] + otb["Gross_Revenue"]
                ),
                "Net_After_Fixed_Costs": _safe_number(
                    ytd["Net_After_Fixed_Costs"] + otb["Net_After_Fixed_Costs"]
                ),
                "Cleaning_Fees": _safe_number(
                    ytd["Cleaning_Fees"] + otb["Cleaning_Fees"]
                ),
            },
        },
        "listing_metrics": listing_rows,
        "monthly_listing_performance": monthly_listing_performance,
        "booking_profiles": booking_profiles,
        "listing_signals": listing_signals,
        "analysis_instruction": (
            "Prioritise actionable insights. Use monthly performance and listing signals "
            "to distinguish seasonality, future booking weakness, medium/long-stay patterns, "
            "and genuine listing underperformance."
        ),
    }


def _render_ai_list(items) -> str:
    if not items:
        return "<p style='font-size:13px; color:#777; margin:0;'>No specific points returned.</p>"
    lis = "".join(f"<li>{_html_escape(item)}</li>" for item in items)
    return f"<ul style='margin:6px 0 0 18px; padding:0; font-size:13px; line-height:1.45;'>{lis}</ul>"


def render_ai_insights_json(ai_text: str) -> str:
    import json
    import re
    import html

    cleaned = ai_text.strip()

    # Remove ```json ... ``` or ``` ... ```
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned.strip(), flags=re.IGNORECASE)
    cleaned = re.sub(r"\s*```$", "", cleaned.strip())

    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        escaped = html.escape(ai_text)
        return f"<pre style='white-space:pre-wrap; font-family:Arial, sans-serif;'>{escaped}</pre>"

    sections = [
        ("🔴 Critical Issues", "critical_issues", "#dc2626"),
        ("🟡 Key Drivers", "key_drivers", "#f59e0b"),
        ("🟢 Recommended Actions", "recommended_actions", "#16a34a"),
        ("⚪ Data Gaps", "data_gaps", "#6b7280"),
    ]

    output = ""

    for title, key, color in sections:
        items = data.get(key, [])
        if not items:
            continue

        output += f"<h4 style='margin:12px 0 6px 0; color:{color};'>{title}</h4>"
        output += "<ul style='margin-top:0; padding-left:20px;'>"

        for item in items:
            output += f"<li>{html.escape(str(item))}</li>"

        output += "</ul>"

    return output

def _plain_ai_fallback_to_html(text: str) -> str:
    lines = [_html_escape(line) for line in text.splitlines() if line.strip()]
    body = "".join(f"<p style='font-size:13px; line-height:1.45; margin:0 0 8px 0;'>{line}</p>" for line in lines)
    return f"""
    <div style="background:white; padding:16px; margin-top:16px; border-radius:8px; border:1px solid #e5e7eb;">
      <h3 style="margin-top:0;">AI Insights & Actions</h3>
      <p style="font-size:12px; color:#666; margin-top:0;">
      Generated from KPI tables and listing-level metrics. Calculations are still performed by Python.
      </p>
      {body}
    </div>
    """

def build_monthly_target_payload(
    monthly_targets: pd.DataFrame,
    listings: pd.DataFrame,
) -> list[dict]:
    targets = monthly_targets.copy()
    listings = listings.copy()

    targets["listing"] = targets["listing"].apply(normalise_listing_key)
    listings["listing"] = listings["listing"].apply(normalise_listing_key)

    targets = targets.merge(
        listings[["listing", "listing_name"]],
        on="listing",
        how="left",
    )

    targets["listing_name"] = targets["listing_name"].fillna(targets["listing"])

    records = []

    for _, row in targets.iterrows():
        records.append({
            "listing": row["listing_name"],
            "year_month": f"{int(row['year'])}-{int(row['month']):02d}",
            "target": _safe_number(row["target"]),
        })

    return records



def generate_ai_insights_html(
    daily_for_report: pd.DataFrame,
    daily_expenses_for_report: pd.DataFrame,
    listings: pd.DataFrame,
    report_group_label: str,
    reservations_for_report: pd.DataFrame | None = None,
    monthly_targets: pd.DataFrame | None = None,
    enabled: bool = False,
) -> str:

    if not enabled:
        return ""

    if OpenAI is None:
        return """
        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px; border:1px solid #e5e7eb;">
          <h3 style="margin-top:0;">AI Insights & Actions</h3>
          <p style="font-size:13px; color:#666;">
          AI insights requested but OpenAI package not installed.
          </p>
        </div>
        """

    if not os.environ.get("OPENAI_API_KEY"):
        return """
        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px; border:1px solid #e5e7eb;">
          <h3 style="margin-top:0;">AI Insights & Actions</h3>
          <p style="font-size:13px; color:#666;">
          OPENAI_API_KEY not set.
          </p>
        </div>
        """

    payload = build_ai_analysis_payload(
        daily_for_report,
        daily_expenses_for_report,
        listings,
        report_group_label,
        reservations_for_report=reservations_for_report,
        monthly_targets=monthly_targets,
    )

    #targets_json = monthly_targets.to_dict(orient="records")

    prompt = f"""
Tu es un analyste pragmatique spécialisé dans la rentabilité de locations courte durée.

Objectif :
Fournir des insights utiles et directement actionnables, pas un résumé.

Règles d’interprétation IMPORTANTES :

- Le coût fixe par nuit augmente mécaniquement quand le taux d’occupation est faible.
- Ne mentionner les coûts fixes qu’UNE SEULE FOIS maximum, sauf anomalie évidente.
- Ne pas surinterpréter les coûts fixes.

- Les coûts fixes sont peu actionnables à court terme.
- Prioriser les leviers réellement actionnables :
  - taux d’occupation
  - niveau de prix
  - durée de séjour
  - coûts variables (ménage, conciergerie, énergie)
  - performance par logement

- Ne PAS utiliser le taux d’occupation sur l’année complète comme indicateur principal.
- Se concentrer sur :
  - le réalisé depuis le début de l’année (vs objectifs)
  - les réservations confirmées à venir

- Ne PAS utiliser de noms de variables, ni d’anglais technique (ADR, YTD, OTB, etc).
- Utiliser uniquement un langage naturel en français.

- comprenez qu'un taux d'occupation de 100% pendant plus de deux mois signifie un autre type de location (de moyenne ou de longue durée). Une reflexion est possible sur la rentabilité de ce type d'activité vs la location courte durée en prenant en compte les frais reduit associés et le taux d'occupation garantie, mais verus des revenues inferieures.
- comprenez qu'un taux d'occupation à 0% pendant plus de 4 mois signifie un bien qui n'est pas commercialisé mais qui contribue quand meme aux frais fixes.

- Distinguer clairement les causes :
  1. manque d’occupation
  2. prix inadapté
  3. coûts variables trop élevés
  4. problème structurel réel

Format attendu :

- Répondre en FRANÇAIS uniquement
- Être concis et concret
- Maximum :
  - 3 points critiques
  - 3 facteurs explicatifs
  - 3 actions recommandées
- Pas de répétition
- Pas de phrases longues
Commencer par le problème le plus important économiquement.

Contexte saisonnier :
- Utiliser une connaissance générale de la saisonnalité des locations courte durée en zone touristique côtière française (bretagne).
- Ne pas inventer de chiffres de marché.
- Tenir compte du fait que l’hiver et le début de printemps peuvent être faibles, tandis que l’été est normalement beaucoup plus porteur.
- Ne pas conclure trop vite qu’un faible taux d’occupation en basse saison est structurel.
Quand tu analyses les réservations confirmées à venir, tiens compte du calendrier : une faible occupation future peut être normale si la période est encore éloignée ou hors saison, mais plus préoccupante si l’on approche d’une période normalement forte.
- Distinguer :
  1. faiblesse normale de basse saison,
  2. faiblesse inquiétante des réservations à venir,
  3. problème spécifique à un logement.
- Si une recommandation dépend de la saison, le dire explicitement.

Retourner UNIQUEMENT du JSON valide :

{{
  "critical_issues": [],
  "key_drivers": [],
  "recommended_actions": [],
  "data_gaps": []
}}

Données :
{json.dumps(payload, ensure_ascii=False, indent=2)}


"""
    
    prompt=f"""
    Analyse PRIORITAIRE :

    1. Identifier les logements réellement sous-performants.
    2. Distinguer un problème structurel d’un simple effet saisonnier.
    3. Comparer le réalisé aux objectifs mensuels (targets) - focus sur le passé et le futur proche (voir ci-dessous)
    4. Noter que les reservations se font au fur et à mesure, il est normal que le calendrier soit vide 2 mois avant (sauf pour la periode estivale)
    4. Identifier les leviers commerciaux réalistes :
    - occupation
    - prix
    - durée de séjour
    - réduction des coûts variables
    5. Ne mentionner les coûts fixes qu’en dernier niveau d’analyse, qui sont en generale non-compressible (eg remboursement de credit immobilier)

    Tu dois commenter les écarts aux objectifs si les données "monthly_vs_target" sont présentes.
Identifier les mois/logements les plus en retard par rapport aux objectifs.
Ne pas simplement mentionner les objectifs : expliquer où l’écart est le plus important et ce que cela implique.

    Adopter une logique d’exploitant :
    - éviter les conclusions alarmistes
    - éviter les répétitions
    - privilégier les recommandations concrètes et réalistes

    Contexte touristique 2026 (France / Bretagne) :

- Le tourisme domestique français reste globalement résilient en 2026.
- Les destinations côtières françaises devraient rester relativement demandées pendant l’été.
- La Bretagne reste un marché très saisonnier, avec une forte concentration de la demande entre juin et août.
- Les réservations hors saison restent naturellement plus faibles.
- Les voyageurs semblent réserver plus tardivement qu’avant, dans un contexte économique plus incertain.
- Les séjours plus longs et les locations de moyenne durée deviennent plus fréquents.
- Un faible niveau de réservation en été est plus préoccupant qu’un faible niveau de réservation en hiver.
- Ne pas considérer automatiquement un faible remplissage futur comme structurel si la période est encore éloignée.

    Retourner UNIQUEMENT du JSON valide :

    {{
    "critical_issues": [],
    "target_gaps": [],
    "key_drivers": [],
    "recommended_actions": [],
    "data_gaps": []
    }}

    IMPORTANT :
    - Toutes les sections doivent contenir uniquement des phrases simples.
    - Ne pas retourner d’objets JSON imbriqués.
    - Pas d’analyse mois par mois.
    - Faire une synthèse de haut niveau.
    - Mentionner les objectifs uniquement lorsqu’un écart est significatif ou révélateur.
    - Éviter les répétitions.
    - repondre en francais et ne pas utiliser des acronymes.

    Données :
    
    {json.dumps(payload, ensure_ascii=False, indent=2)}


"""

    try:
        client = OpenAI()
        response = client.responses.create(
            model=os.environ.get("OPENAI_INSIGHTS_MODEL", "gpt-4.1-mini"),
            input=prompt,
        )

        ai_text = response.output_text.strip()

        # Extract JSON safely
        start = ai_text.find("{")
        end = ai_text.rfind("}")

        if start == -1 or end == -1:
            return f"<pre>{_html_escape(ai_text)}</pre>"

        json_text = ai_text[start:end + 1]

        try:
            sections = json.loads(json_text)
        except json.JSONDecodeError:
            return f"<pre>{_html_escape(ai_text)}</pre>"

        return render_structured_ai_insights(sections)

    except Exception as exc:
        return f"""
        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px; border:1px solid #e5e7eb;">
          <h3>AI Insights & Actions</h3>
          <p style="color:#666;">OpenAI error: {_html_escape(str(exc))}</p>
        </div>
        """
    

def build_monthly_vs_target_payload(
    monthly_listing_performance: list[dict],
    monthly_targets: pd.DataFrame,
    listings: pd.DataFrame,
) -> list[dict]:
    perf = pd.DataFrame(monthly_listing_performance)
    targets = monthly_targets.copy()
    listings = listings.copy()

    if perf.empty or targets.empty:
        return []

    targets["listing_key"] = targets["listing"].apply(normalise_listing_key)
    listings["listing_key"] = listings["listing"].apply(normalise_listing_key)

    targets = targets.merge(
        listings[["listing_key", "listing_name"]],
        on="listing_key",
        how="left",
    )

    targets["listing"] = targets["listing_name"].fillna(targets["listing_key"])

    targets["month"] = (
        targets["year"].astype(int).astype(str)
        + "-"
        + targets["month"].astype(int).astype(str).str.zfill(2)
    )

    targets["monthly_target"] = pd.to_numeric(
        targets["target"],
        errors="coerce",
    )

    comparison = perf.merge(
        targets[["listing", "month", "monthly_target"]],
        on=["listing", "month"],
        how="left",
    )

    comparison["gap_to_target"] = (
        comparison["gross_revenue"] - comparison["monthly_target"]
    )

    comparison["pct_of_target"] = (
        comparison["gross_revenue"]
        / comparison["monthly_target"].replace(0, pd.NA)
    )

    records = []

    for _, row in comparison.iterrows():
        if pd.isna(row.get("monthly_target")):
            continue

        records.append({
            "listing": row["listing"],
            "month": row["month"],
            "period_type": row["period_type"],
            "gross_revenue": _safe_number(row["gross_revenue"]),
            "monthly_target": _safe_number(row["monthly_target"]),
            "gap_to_target": _safe_number(row["gap_to_target"]),
            "pct_of_target": _safe_number(row["pct_of_target"]),
            "occupancy": _safe_number(row["occupancy"]),
            "booked_nights": int(row["booked_nights"]),
        })

    return records
    
def render_structured_ai_insights(sections: dict) -> str:
    section_config = [
        ("critical_issues", "🔴 Points critiques", "#dc2626"),
        ("target_gaps", "🎯 Écarts aux objectifs", "#2563eb"),
        ("key_drivers", "🟡 Facteurs explicatifs", "#f59e0b"),
        ("recommended_actions", "🟢 Actions recommandées", "#16a34a"),
        ("data_gaps", "⚪ Données manquantes", "#6b7280"),
    ]

    html_parts = []

    for key, title, color in section_config:
        items = sections.get(key, [])

        if not items:
            continue

        html_parts.append(
            f'<h4 style="margin:14px 0 6px 0; color:{color};">{title}</h4>'
        )

        html_parts.append(
            '<ul style="margin-top:4px; padding-left:20px; line-height:1.4;">'
        )

        for item in items:
            html_parts.append(
                f"<li>{_html_escape(str(item))}</li>"
            )

        html_parts.append("</ul>")

    if not html_parts:
        return """
        <p style="font-size:13px; color:#666;">
        Aucune analyse structurée n’a été générée.
        </p>
        """

    return "\n".join(html_parts)

def build_finance_email_html(
    gross_revenue_pivot: pd.DataFrame,
    contribution_pivot: pd.DataFrame,
    net_after_fixed_pivot: pd.DataFrame,
    ADR_pivot: pd.DataFrame,
    occupancy_pivot: pd.DataFrame,
    cleaning_fees_pivot: pd.DataFrame,
    concierge_fees_pivot: pd.DataFrame,
    cleaning_pct_net_pivot: pd.DataFrame,
    finance_chart_path: Path,
    listing_net_chart_path: Path,
    occupancy_chart_path: Path,
    finance_kpis: list[list[dict]],
    performance_by_listing_table: pd.DataFrame,
    funnel_table: pd.DataFrame,
    ai_insights_html: str = "",
    report_group_label: str = "All groups",
) -> str:

    chart_src = image_to_base64_data_uri(finance_chart_path)
    listing_net_chart_src = image_to_base64_data_uri(listing_net_chart_path)
    occupancy_chart_src = image_to_base64_data_uri(occupancy_chart_path)
    kpi_cards = render_finance_kpi_cards(finance_kpis)

    performance_by_listing_html = dataframe_to_html_table(
        performance_by_listing_table,
        color_percentages=False,
    )

    funnel_html = dataframe_to_html_table(
        funnel_table,
        color_percentages=False,
    )

    gross_table = dataframe_to_html_table(
        gross_revenue_pivot,
        currency_columns=[col for col in gross_revenue_pivot.columns if col != "listing_name"],
    )

    contribution_table = dataframe_to_html_table(
        contribution_pivot,
        currency_columns=[col for col in contribution_pivot.columns if col != "listing_name"],
    )

    net_table = dataframe_to_html_table(
        net_after_fixed_pivot,
        currency_columns=[col for col in net_after_fixed_pivot.columns if col != "listing_name"],
    )

    ADR_table = dataframe_to_html_table(
        ADR_pivot,
        currency_columns=[col for col in ADR_pivot.columns if col != "listing_name"],
    )

    occupancy_table = dataframe_to_html_table(occupancy_pivot)

    cleaning_table = dataframe_to_html_table(
        cleaning_fees_pivot,
        currency_columns=[col for col in cleaning_fees_pivot.columns if col != "listing_name"],
    )

    concierge_table = dataframe_to_html_table(
        concierge_fees_pivot,
        currency_columns=[col for col in concierge_fees_pivot.columns if col != "listing_name"],
    )

    cleaning_pct_table = dataframe_to_html_table(
        cleaning_pct_net_pivot,
        color_percentages=False,
    )

    html = f"""
    <html>
    <body style="margin:0; padding:0; background:#f6f7f9; font-family:Arial, sans-serif; color:#222;">
      <div style="max-width:1000px; margin:0 auto; padding:24px;">

        <div style="background:#1f2937; color:white; padding:20px; border-radius:8px;">
          <h2 style="margin:0;">Finance Report</h2>
          <p style="margin:6px 0 0 0;">Revenue and profitability by listing — {report_group_label}</p>
        </div>

        {kpi_cards}

        {ai_insights_html}

        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px;">
        <h3 style="margin-top:0;">Cumulative Performance — Actual + On the Books</h3>
        <p style="font-size:12px; color:#666; margin-top:0;">
        Solid lines are actuals. Dashed lines are future on-the-books values. The vertical line marks today.
        </p>
        <img src="{chart_src}" style="width:100%; max-width:900px;">
        </div>

        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px;">
        <h3 style="margin-top:0;">Performance by Listing</h3>
        <p style="font-size:12px; color:#666; margin-top:0;">
        YTD columns show actual performance to yesterday. On-the-books columns show confirmed future bookings.
        </p>
        {performance_by_listing_html}
        </div>

        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px;">
        <h3 style="margin-top:0;">Profitability Funnel by Listing</h3>
        <p style="font-size:12px; color:#666; margin-top:0;">
        Funnel values are YTD actuals. Variable costs per night include cleaning, concierge and energy/usage-based variable expenses.
        </p>
        {funnel_html}
        </div>

        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px;">
        <h3 style="margin-top:0;">Net Revenue by Listing — Actual + On the Books</h3>
        <p style="font-size:12px; color:#666; margin-top:0;">
        Solid lines are actuals. Dashed lines are future on-the-books values. The vertical line marks today.
        </p>
        <img src="{listing_net_chart_src}" style="width:100%; max-width:900px;">
        </div>

        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px;">
        <h3 style="margin-top:0;">Occupancy by Listing — Actual + On the Books</h3>
        <p style="font-size:12px; color:#666; margin-top:0;">
        One line per listing. Solid lines are actual months. Dashed lines are current/future on-the-books months. The vertical line marks today.
        </p>
        <img src="{occupancy_chart_src}" style="width:100%; max-width:900px;">
        </div>

        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px;">
          <h3 style="margin-top:0;">Gross Revenue</h3>
          {gross_table}
        </div>

        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px;">
        <h3 style="margin-top:0;">Cleaning Fees</h3>
        {cleaning_table}
        </div>

        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px;">
        <h3 style="margin-top:0;">Cleaning Fees % of Net Revenue</h3>
        {cleaning_pct_table}
        </div>

        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px;">
        <h3 style="margin-top:0;">Net After Fixed Costs</h3>
        {net_table}
        </div>

        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px;">
        <h3 style="margin-top:0;">Contribution Before Fixed Costs</h3>
        {contribution_table}
        </div>

        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px;">
        <h3 style="margin-top:0;">Average Daily Rate</h3>
        {ADR_table}
        </div>

        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px;">
        <h3 style="margin-top:0;">Occupancy</h3>
        {occupancy_table}
        </div>

        <div style="background:white; padding:16px; margin-top:16px; border-radius:8px;">
        <h3 style="margin-top:0;">Concierge Fees</h3>
        {concierge_table}
        </div>
      </div>
    </body>
    </html>
    """
    return html


def build_cumulative_finance_chart(
    daily: pd.DataFrame,
    daily_targets: pd.DataFrame,
    daily_expenses: pd.DataFrame,
    output_dir: Path,
    listings: pd.DataFrame | None = None,
    filename_suffix: str = "all-groups",
    today: datetime | None = None,
) -> Path:
    """
    Build cumulative performance chart.

    This is not pure YTD:
    - solid line = actuals up to yesterday
    - dashed line = future on-the-books values
    - vertical line = today
    """
    if today is None:
        today = datetime.today()

    today_ts = pd.Timestamp(today).normalize()

    revenue = daily.copy()
    expenses = daily_expenses.copy()
    targets = daily_targets.copy()

    if listings is not None:
        selected_listing_keys = set(listings["listing"].dropna().apply(normalise_listing_key))
        revenue = revenue[revenue["listing"].apply(normalise_listing_key).isin(selected_listing_keys)]
        expenses = expenses[expenses["listing"].apply(normalise_listing_key).isin(selected_listing_keys)]
        targets = targets[targets["listing"].apply(normalise_listing_key).isin(selected_listing_keys)]

    revenue["date"] = pd.to_datetime(revenue["date"])
    expenses["date"] = pd.to_datetime(expenses["date"])
    targets["date"] = pd.to_datetime(targets["date"])

    revenue_daily = (
        revenue.groupby("date", as_index=False)
        .agg(
            gross_revenue=("daily_revenue_total", "sum"),
            cleaning_fees=("daily_cleaning_fees", "sum"),
            concierge_fees=("daily_concierge_commission", "sum"),
        )
    )

    expenses_daily = (
        expenses.groupby(["date", "expense_type"], as_index=False)
        .agg(expense=("daily_expense", "sum"))
        .pivot_table(
            index="date",
            columns="expense_type",
            values="expense",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )

    if "variable" not in expenses_daily.columns:
        expenses_daily["variable"] = 0

    if "fixed" not in expenses_daily.columns:
        expenses_daily["fixed"] = 0

    target_daily = (
        targets.groupby("date", as_index=False)
        .agg(target=("daily_target", "sum"))
    )

    chart_df = revenue_daily.merge(expenses_daily, on="date", how="left")
    chart_df = chart_df.merge(target_daily, on="date", how="left")

    chart_df[["variable", "fixed", "target"]] = chart_df[
        ["variable", "fixed", "target"]
    ].fillna(0)

    chart_df["contribution"] = (
        chart_df["gross_revenue"]
        - chart_df["cleaning_fees"]
        - chart_df["concierge_fees"]
        - chart_df["variable"]
    )

    chart_df["net_after_fixed"] = (
        chart_df["contribution"]
        - chart_df["fixed"]
    )

    chart_df = chart_df.sort_values("date")

    for col in ["gross_revenue", "contribution", "net_after_fixed", "target"]:
        chart_df[f"{col}_cum"] = chart_df[col].cumsum()

    chart_path = output_dir / f"finance_cumulative_actual_otb_{filename_suffix}.png"

    plt.figure(figsize=(10, 5))

    series = [
        ("gross_revenue_cum", "Gross revenue"),
        ("contribution_cum", "Contribution before fixed costs"),
        ("net_after_fixed_cum", "Net after fixed costs"),
        ("target_cum", "Target"),
    ]

    for col, label in series:
        actual = chart_df[chart_df["date"] < today_ts]
        future = chart_df[chart_df["date"] >= today_ts]

        # Actual part: solid.
        if not actual.empty:
            line = plt.plot(actual["date"], actual[col], label=label)
            line_color = line[0].get_color()
        else:
            line = plt.plot([], [], label=label)
            line_color = line[0].get_color()

        # Future part: dashed, starting from the last actual point so the line is continuous.
        if not future.empty:
            if not actual.empty:
                bridge = pd.concat([actual.tail(1), future], ignore_index=True)
            else:
                bridge = future

            plt.plot(
                bridge["date"],
                bridge[col],
                linestyle="--",
                color=line_color,
            )

    plt.axvline(today_ts, linestyle="--", linewidth=1)
    plt.title("Cumulative Performance — Actual + On the Books")
    plt.xlabel("")
    plt.ylabel("€")
    plt.legend()
    plt.grid(True, alpha=0.3)
    plt.tight_layout()

    plt.savefig(chart_path, dpi=150)
    plt.close()

    return chart_path

def image_to_base64_data_uri(image_path: Path) -> str:
    encoded = base64.b64encode(image_path.read_bytes()).decode("utf-8")
    return f"data:image/png;base64,{encoded}"








# =============================================================================
# Main orchestration
# =============================================================================


def main() -> None:
    # --- ARGUMENT PARSING ---
    args = parse_args()
    source = args.source
    sheet_id = args.sheet_id
    input_path = Path(args.input) if args.input else None

    if args.output:
        output_dir = Path(args.output)
        output_dir.mkdir(parents=True, exist_ok=True)
    else:
        output_dir = get_output_dir()

    mode = args.mode
    dry_run = args.dry_run
    selected_groups = parse_group_filter(args.groups)

    print(f"Mode: {mode}")
    print(f"Dry run: {dry_run}")
    print(f"Source: {source}")
    print(f"Output: {output_dir}")
    print(f"Finance groups: {group_label(selected_groups)}")
    print(f"AI insights: {args.ai_insights}")

    if source == "excel":
        if input_path is None:
            raise ValueError("Excel source requires --input")

        if not input_path.exists():
            raise FileNotFoundError(f"Workbook not found: {input_path}")

        print(f"Reading workbook: {input_path}")

    elif source == "google-sheets":
        if not sheet_id:
            raise ValueError("Google Sheets source requires --sheet-id or GOOGLE_SHEET_ID")

        print(f"Reading Google Sheet: {sheet_id}")
        analysis_year = 2026

    # -----------------------------
    # Reservations
    # -----------------------------
    df = load_reservations(source, input_path, sheet_id)

    required_reservation_columns = [
        "Listing",
        "Check in Date",
        "Number of Nights",
        "Revenue Net",
    ]
    validate_required_columns(df, required_reservation_columns)

    df = drop_empty_rows(df)
    df = rename_reservation_columns(df)
    df = add_optional_reservation_columns(df)

    reservation_date_columns = ["booking_date", "checkin_date"]
    reservation_numeric_columns = [
        "listing",
        "nights",
        "total_revenue",
        "cleaning_fees",
        "concierge_commission",
        "revenue_net",
    ]

    df = parse_types(df, reservation_date_columns, reservation_numeric_columns)
    df["listing"] = df["listing"].apply(normalise_listing_key)
    df = add_checkout_date(df)

    df_listings = load_listings_lookup(source, input_path, sheet_id)
    finance_listings = filter_listings_by_group(df_listings, selected_groups)
    df = enrich_with_listing_info(df, df_listings)

    validate_reservations(df)

    overlaps = find_overlaps(df)
    if not overlaps.empty:
        overlaps_path = output_dir / "overlaps.csv"
        overlaps.to_csv(overlaps_path, index=False)
        print("\nOverlapping bookings found:")
        print(overlaps)
        print(f"\nSaved overlap details to: {overlaps_path}")
        raise ValueError("Overlap validation failed")

    df_reservations = df.copy()

    # -----------------------------
    # Calendar + daily occupancy
    # Use the Listings table, not reservations, so listings with no bookings
    # still exist in the calendar and Power BI model.
    # -----------------------------
    calendar = build_calendar(analysis_year, df_listings)
    booking_nights = expand_booking_nights(df)
    daily = build_daily_occupancy(calendar, booking_nights)

    # -----------------------------
    # Reporting group filter
    # -----------------------------
    # The same --groups argument is used for both finance and operational emails.
    # If no group is supplied, finance_listings contains all listings, so these
    # filtered tables are equivalent to the full tables.
    selected_listing_keys = set(
        finance_listings["listing"].dropna().apply(normalise_listing_key)
    )

    daily_for_report = daily[
        daily["listing"].apply(normalise_listing_key).isin(selected_listing_keys)
    ].copy()

    reservations_for_report = df_reservations[
        df_reservations["listing"].apply(normalise_listing_key).isin(selected_listing_keys)
    ].copy()


    # -----------------------------
    # Monthly targets
    # -----------------------------
    required_target_columns = ["Month", "Year"]

    df_targets = load_targets(source, input_path, sheet_id)
    df_targets = df_targets.dropna(how="all")
    df_targets = df_targets[
        ~df_targets["Month"].astype(str).str.strip().str.lower().eq("sum")
    ]

    validate_required_columns(df_targets, required_target_columns)

    df_targets = rename_target_columns(df_targets)

    # Ensure listing columns become strings
    df_targets.columns = [
        str(col) if i >= 2 else col
        for i, col in enumerate(df_targets.columns)
    ]

    target_date_columns = []
    target_numeric_columns = ["year", "month"]
    target_numeric_columns.extend(df_targets.columns[2:].tolist())

    df_targets = parse_types(df_targets, target_date_columns, target_numeric_columns)

    expanded_targets = expand_targets(df_targets)
    daily_targets = build_daily_targets(calendar, expanded_targets)

    # -----------------------------
    # Final joined revenue / target table
    # -----------------------------
    final = daily.merge(
        daily_targets[["date", "year", "month", "listing", "daily_target"]],
        on=["date", "year", "month", "listing"],
        how="left",
    )

    final["daily_target"] = final["daily_target"].fillna(0)
    final["gap"] = final["daily_revenue_net"] - final["daily_target"]

    # -----------------------------
    # Expenses
    # -----------------------------
    df_fixed_expenses = load_fixed_expenses(source, input_path, sheet_id)
    df_variable_expenses = load_variable_expenses(source, input_path, sheet_id)

    df_fixed_expenses_long = expand_fixed_expenses(df_fixed_expenses)
    df_variable_expenses_long = expand_variable_expenses(df_variable_expenses)

    df_expanded_expenses = combine_monthly_expenses(
        df_fixed_expenses_long,
        df_variable_expenses_long
    )
    df_expanded_daily_expenses = expand_daily_expenses(df_expanded_expenses)

    daily_expenses_for_report = df_expanded_daily_expenses[
        df_expanded_daily_expenses["listing"].apply(normalise_listing_key).isin(selected_listing_keys)
    ].copy()


    monthly_revenue = booking_nights.copy()
    monthly_revenue["year_month"] = monthly_revenue["date"].dt.strftime("%Y-%m")

    monthly_revenue = monthly_revenue.merge(
        df_listings[["listing", "listing_name"]],
        on="listing",
        how="left"
    )

    monthly_finance = build_monthly_finance_base(
        booking_nights,
        df_expanded_expenses,
        finance_listings,
    )

    gross_revenue_pivot = build_monthly_listing_pivot(
        monthly_finance,
        value_col="Gross_Revenue",
    )

    contribution_pivot = build_monthly_listing_pivot(
        monthly_finance,
        value_col="Contribution_Before_Fixed_Costs",
    )

    net_after_fixed_pivot = build_monthly_listing_pivot(
        monthly_finance,
        value_col="Net_After_Fixed_Costs",
    )

    ADR = build_monthly_listing_pivot(
        monthly_finance,
        value_col="ADR",
    )

    ADR = ADR.drop(columns=["Total"], errors="ignore")
    ADR = ADR[ADR["listing_name"] != "TOTAL"]

    monthly_occupancy = build_monthly_occupancy_base(
        daily_for_report,
        finance_listings,
    )

    occupancy_pivot = build_monthly_listing_pivot(
        monthly_occupancy,
        value_col="Occupancy_Rate",
    )

    # Remove misleading additive totals
    occupancy_pivot = occupancy_pivot.drop(columns=["Total"], errors="ignore")
    occupancy_pivot = occupancy_pivot[occupancy_pivot["listing_name"] != "TOTAL"]

    cleaning_fees_pivot = build_monthly_listing_pivot(
        monthly_finance,
        value_col="Cleaning_Fees",
    )

    concierge_fees_pivot = build_monthly_listing_pivot(
        monthly_finance,
        value_col="Concierge_Fees",
    )

    cleaning_pct_net_pivot = build_cleaning_pct_net_pivot(monthly_finance)


    month_cols = [
        col for col in occupancy_pivot.columns
        if col != "listing_name"
    ]

    for col in month_cols:
        occupancy_pivot[col] = (
            occupancy_pivot[col] * 100
        ).round(0).astype("Int64").astype(str) + "%"

    finance_kpis = build_finance_kpis(
        daily_for_report,
        daily_expenses_for_report,
    )

    performance_by_listing_table, funnel_table = build_apartment_kpi_table(
        daily_for_report,
        daily_expenses_for_report,
        finance_listings,
    )

    # -----------------------------
    # Save outputs
    # -----------------------------
    df.to_csv(output_dir / "reservations_clean.csv", index=False)
    df_listings.to_csv(output_dir / "listings.csv", index=False)
    finance_listings.to_csv(output_dir / "finance_listings.csv", index=False)

    calendar.to_csv(output_dir / "calendar.csv", index=False)
    booking_nights.to_csv(output_dir / "booking_nights.csv", index=False)
    daily.to_csv(output_dir / "daily_occupancy.csv", index=False)

    df_targets.to_csv(output_dir / "targets_wide_clean.csv", index=False)
    expanded_targets.to_csv(output_dir / "targets_long.csv", index=False)
    daily_targets.to_csv(output_dir / "daily_targets.csv", index=False)

    df_fixed_expenses_long.to_csv(output_dir / "fixed_expenses_monthly_long.csv", index=False)
    df_variable_expenses_long.to_csv(output_dir / "variable_expenses_monthly_long.csv", index=False)
    df_expanded_expenses.to_csv(output_dir / "expenses_monthly_long.csv", index=False)
    df_expanded_daily_expenses.to_csv(output_dir / "expenses_long.csv", index=False)

    final.to_csv(output_dir / "daily_performance.csv", index=False)
    performance_by_listing_table.to_csv(output_dir / f"performance_by_listing_{group_slug(selected_groups)}.csv", index=False)
    funnel_table.to_csv(output_dir / f"funnel_by_listing_{group_slug(selected_groups)}.csv", index=False)

    print("Reservations loaded and validated successfully.")
    print(f"Outputs saved to: {output_dir}")


    weekly_occupancy = build_weekly_occupancy(daily_for_report)

    weekly_occupancy = weekly_occupancy.merge(
        df_listings[["listing", "listing_name"]],
        on="listing",
        how="left"
        )
    
    weekly_occupancy["listing"] = weekly_occupancy["listing"].astype(str)
    df_listings["listing"] = df_listings["listing"].astype(str)

   # weekly = weekly.rename(columns={"Name": "listing_name"})
    weekly_occupancy_pivot = pivot_weekly_occupancy(weekly_occupancy)

    weekly_occupancy.to_csv(output_dir / "weekly_occupancy_long.csv", index=False)
    weekly_occupancy_pivot.to_csv(output_dir / "weekly_occupancy_pivot.csv", index=False)

    revenue_this_week = build_revenue_this_week(
        reservations_for_report,
        daily_expenses_for_report
    )


    # -----------------------------
    # Optional weekly email
    # -----------------------------
    if mode == "weekly-email":
        summary = build_weekly_arrivals_departures_summary(reservations_for_report)
        html = build_weekly_email_html(
            summary,
            weekly_occupancy_pivot,
            revenue_this_week,
            report_group_label=group_label(selected_groups),
        )

        subject = (
            f"Weekly Reservations Summary — {group_label(selected_groups)} "
            f"{summary['start_of_week'].strftime('%d %b')} – "
            f"{summary['end_of_week'].strftime('%d %b')}"
        )

        if dry_run:
            preview_path = output_dir / f"email_preview_{group_slug(selected_groups)}.html"
            preview_path.write_text(html, encoding="utf-8")
            print(f"Dry run: email preview saved to {preview_path}")
        else:
            send_html_email(subject, html)
            print("Weekly email sent.")

    if mode == "finance-email":

        finance_chart_path = build_cumulative_finance_chart(
            daily,
            daily_targets,
            df_expanded_daily_expenses,
            output_dir,
            listings=finance_listings,
            filename_suffix=group_slug(selected_groups),
        )

        listing_net_chart_path = build_listing_ytd_net_chart(
            daily_for_report,
            daily_expenses_for_report,
            finance_listings,
            output_dir,
            filename_suffix=group_slug(selected_groups),
        )

        occupancy_chart_path = build_monthly_occupancy_chart(
            monthly_occupancy,
            output_dir,
            filename_suffix=group_slug(selected_groups),
        )

        ai_insights_html = generate_ai_insights_html(
            daily_for_report,
            daily_expenses_for_report,
            finance_listings,
            report_group_label=group_label(selected_groups),
            reservations_for_report=reservations_for_report,
            monthly_targets=expanded_targets,
            enabled=args.ai_insights,
        )

        html = build_finance_email_html(
            gross_revenue_pivot,
            contribution_pivot,
            net_after_fixed_pivot,
            ADR,
            occupancy_pivot,
            cleaning_fees_pivot,
            concierge_fees_pivot,
            cleaning_pct_net_pivot,
            finance_chart_path,
            listing_net_chart_path,
            occupancy_chart_path,
            finance_kpis,
            performance_by_listing_table,
            funnel_table,
            ai_insights_html=ai_insights_html,
            report_group_label=group_label(selected_groups),
        )

        subject = f"Finance Report — {group_label(selected_groups)}"

        if dry_run:
            preview_path = output_dir / f"finance_email_preview_{group_slug(selected_groups)}.html"
            preview_path.write_text(html, encoding="utf-8")
            print(f"Dry run: finance email preview saved to {preview_path}")
        else:
            send_html_email(subject, html)
            print("Finance email sent.")


if __name__ == "__main__":
    main()
