from pathlib import Path
import pandas as pd
import datetime
import calendar
import sys 

import sys
print("PYTHON USED:", sys.executable)

from dotenv import load_dotenv
import os

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")


from pathlib import Path


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

# -----------------------------
# 1. LOAD RESERVATIONS SHEET
# -----------------------------
def load_reservations(path: Path, sheet_name: str = "Reservations") -> pd.DataFrame:
    """
    Load the Reservations sheet from the Excel workbook.

    Why read_excel?
    - The source file is an Excel workbook, not CSV
    - We can target a specific sheet by name
    """
    df = pd.read_excel(path, sheet_name=sheet_name)
    return df


def load_expenses(path: Path, sheet_name: str = "Expenses") -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet_name)
    return df

def load_listings_lookup(input_path: Path) -> pd.DataFrame:
    df = pd.read_excel(input_path, sheet_name="Listings")

    df.columns = df.columns.str.strip()  # clean headers

    return df

def enrich_with_listing_info(df: pd.DataFrame, listings: pd.DataFrame) -> pd.DataFrame:
    df = df.merge(
        listings.rename(columns={"Listing": "listing"}),
        on="listing",
        how="left"
    )

    return df

# -----------------------------
# 2. VALIDATE REQUIRED COLUMNS
# -----------------------------
def validate_required_columns(df: pd.DataFrame, required_columns: list[str]) -> None:
    """
    Check that the expected Excel columns exist before we do any transformations.

    This is a schema check:
    - are the columns we depend on actually present?
    """
    missing = [col for col in required_columns if col not in df.columns]

    if missing:
        raise ValueError(f"Missing required columns: {missing}")


# -----------------------------
# 3. RENAME COLUMNS
# -----------------------------
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
    }

    df = df.rename(columns=column_map)
    return df


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

# -----------------------------
# 4. PARSE TYPES
# -----------------------------
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


# -----------------------------
# 5. DROP FULLY EMPTY ROWS
# -----------------------------
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


# -----------------------------
# 6. ADD CHECKOUT DATE
# -----------------------------
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


# -----------------------------
# 7. VALIDATE ROW-LEVEL LOGIC
# -----------------------------
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


# -----------------------------
# 8. FIND OVERLAPS
# -----------------------------
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


def expand_booking_nights(df: pd.dataFrame) -> pd.DataFrame:
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
                "listing": row["listing"],
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
    df["listing"] = df["listing"].astype(str)
    return expanded_df

def build_calendar(year, df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    df = df.copy()
    start_date = datetime(int(year), 1, 1)
    end_date = datetime(int(year), 12, 31)
    dates = pd.date_range(start=start_date, end=end_date)

    for listing in df["listing"].dropna().unique():

        for date in dates:
            rows.append({
                "date": date,
                "year": date.year,
                "month": date.month,
                "listing": listing,
            })

    calendar = pd.DataFrame(rows)
    calendar = calendar.sort_values("listing")
    calendar["listing"] = calendar["listing"].astype(str)
    return calendar

def build_daily_occupancy(calendar: pd.DataFrame, booking_nights: pd.DataFrame) -> pd.DataFrame:
    """
    Merge full calendar with booking nights to create a complete daily occupancy table.

    Output:
    - one row per date + listing
    - includes both occupied and unoccupied nights
    """
    booking_nights["listing"] = booking_nights["listing"].astype(str)
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

def build_daily_targets(df_calendar: pd.DataFrame, monthly_targets: pd.DataFrame) -> pd.DataFrame:
    """
    Merge full calendar with targets to create a complete daily targets table.

    Output:
    - one row per date + listing
    """
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

# -----------------------------
# 1. LOAD TARGETS SHEET
# -----------------------------
def load_targets(path: Path, sheet_name: str = "Monthly Targets") -> pd.DataFrame:
    """
    Load the Targets sheet from the Excel workbook.

    """
    df = pd.read_excel(path, sheet_name=sheet_name)
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

    melted = df.melt(id_vars=["year", "month"], var_name="listing", value_name="target")
    melted["listing"] = melted["listing"].astype(str)
    return melted



def expand_expenses(df: pd.DataFrame) -> pd.DataFrame:

    # clean column names if needed
    df = df.copy()

    month_cols = [col for col in df.columns if col.startswith("2026-")]
    # melt monthly columns
    df_long = df.melt(
        id_vars=["Description", "Listing"],
        value_vars=month_cols,
        var_name="year_month",
        value_name="monthly_expense"
    )
    df_long["year"] = df_long["year_month"].str[:4].astype(int)
    df_long["month"] = df_long["year_month"].str[5:7].astype(int)

    df_long["days_in_month"] = df_long.apply(
        lambda row: calendar.monthrange(row["year"], row["month"])[1],
        axis=1
    )

    df_long["daily_expense"] = df_long["monthly_expense"] / df_long["days_in_month"]

    return df_long


def expand_daily_expenses(df_long: pd.DataFrame) -> pd.DataFrame:
    rows = []

    for _, row in df_long.iterrows():

        start_date = datetime(row["year"], row["month"], 1)
        end_date = datetime(row["year"], row["month"],
                            calendar.monthrange(row["year"], row["month"])[1])

        dates = pd.date_range(start=start_date, end=end_date)

        for date in dates:
            rows.append({
                "date": date,
                "listing": str(row["Listing"]),
                "description": row["Description"],
                "daily_expense": row["daily_expense"]
            })

    return pd.DataFrame(rows)



def build_weekly_summary(df) -> dict:
    today = datetime.today()

    # Monday to Sunday
    start_of_week = today - timedelta(days=today.weekday())
    end_of_week = start_of_week +timedelta(days=6)


    arrivals = df[
        (df["checkin_date"] >= start_of_week) &
        (df["checkin_date"] <= end_of_week)
    ]

    departures = df[
        (df["checkout_date"] >= start_of_week) &
        (df["checkout_date"] <= end_of_week)
    ]

    return {
        "arrivals": arrivals,
        "departures": departures,
        "num_arrivals": len(arrivals),
        "num_departures": len(departures)
    }


import os
import smtplib
from email.message import EmailMessage
from datetime import datetime, timedelta
import pandas as pd


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


def dataframe_to_html_table(df: pd.DataFrame, columns: list[str]) -> str:
    """
    Convert selected dataframe columns into a simple HTML table.

    Keeping this small and explicit makes the email predictable.
    """

    if df.empty:
        return "<p><em>None</em></p>"

    html = """
    <table border="1" cellspacing="0" cellpadding="6" style="border-collapse: collapse;">
        <thead>
            <tr>
    """

    for col in columns:
        html += f"<th>{col}</th>"

    html += """
            </tr>
        </thead>
        <tbody>
    """

    for _, row in df.iterrows():
        html += "<tr>"
        for col in columns:
            value = row.get(col, "")

            if isinstance(value, pd.Timestamp):
                value = value.strftime("%a %d %b")

            html += f"<td>{value}</td>"
        html += "</tr>"

    html += """
        </tbody>
    </table>
    """

    return html


def build_weekly_email_html(summary: dict) -> str:
    """
    Build the HTML body for the weekly operational email.
    """

    start = summary["start_of_week"].strftime("%d %b")
    end = summary["end_of_week"].strftime("%d %b %Y")

    arrivals_table = dataframe_to_html_table(
        summary["arrivals"],
        ["checkin_date", "guest_name", "listing_name", "nights", "booking_source"]
    )

    departures_table = dataframe_to_html_table(
        summary["departures"],
        ["checkout_date", "guest_name", "listing", "nights", "booking_source"]
    )

    html = f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #222;">
        <h2>Weekly Reservations Summary</h2>

        <p><strong>Period:</strong> {start} – {end}</p>

        <p>
            <strong>Arrivals:</strong> {summary["num_arrivals"]}<br>
            <strong>Departures:</strong> {summary["num_departures"]}
        </p>

        <h3>Arrivals</h3>
        {arrivals_table}

        <h3>Departures</h3>
        {departures_table}

        <p style="font-size: 12px; color: #666; margin-top: 24px;">
            Generated automatically from the reservations workbook.
        </p>
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





# -----------------------------
# 9. MAIN
# -----------------------------
def main() -> None:
 # --- ARGUMENT PARSING ---
    if len(sys.argv) >= 2:
        input_path = Path(sys.argv[1])
        print("Mode: external workbook path provided")
    else:
        input_path = Path("Reservations.xlsm")
        print("Mode: debug/default workbook path")

    # NEW: optional mode
    if len(sys.argv) >= 3:
        mode = sys.argv[2]
    else:
        mode = "update"   # default

    mode = "weekly-email"
    print(f"Run mode: {mode}")

    # --- VALIDATION ---
    if not input_path.exists():
        raise FileNotFoundError(f"Workbook not found: {input_path}")

    print(f"Reading workbook: {input_path}")  
   
    output_dir = get_output_dir()

    required_reservation_columns = [
        "Listing",
        "Check in Date",
        "Number of Nights",
        "Revenue Net",
    ]

    # -----------------------------
    # Reservations
    # -----------------------------
    df = load_reservations(input_path, sheet_name="Reservations")
    df_listings = load_listings_lookup(input_path)

    required_reservation_columns = [
        "Listing",
        "Check in Date",
        "Number of Nights",
        "Revenue Net",
    ]

    validate_required_columns(df, required_reservation_columns)

    df = drop_empty_rows(df)
    df = rename_reservation_columns(df)


    df = enrich_with_listing_info(df, df_listings)

    
    df = load_reservations(input_path, sheet_name="Reservations")
    df_listings = load_listings_lookup(input_path)


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
    df = add_checkout_date(df)
    validate_reservations(df)

    overlaps = find_overlaps(df)
    if not overlaps.empty:
        overlaps_path = output_dir / "overlaps.csv"
        overlaps.to_csv(overlaps_path, index=False)
        print("\nOverlapping bookings found:")
        print(overlaps)
        print(f"\nSaved overlap details to: {overlaps_path}")
        raise ValueError("Overlap validation failed")

    # -----------------------------
    # Calendar + daily occupancy
    # -----------------------------
    analysis_year = 2026


    df_reservations = df.copy()

    calendar = build_calendar(analysis_year, df)
    booking_nights = expand_booking_nights(df)
    daily = build_daily_occupancy(calendar, booking_nights)

    # -----------------------------
    # Monthly targets
    # -----------------------------
    required_target_columns = [
        "Month",
        "Year",
    ]

    df_targets = load_targets(input_path, sheet_name="Monthly Targets")
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
    # Final joined table
    # -----------------------------
    final = daily.merge(
        daily_targets[["date", "year", "month", "listing", "daily_target"]],
        on=["date", "year", "month", "listing"],
        how="left",
    )

    final["daily_target"] = final["daily_target"].fillna(0)
    final["gap"] = final["daily_revenue_net"] - final["daily_target"]

    # -----------------------------
    # Save outputs
    # -----------------------------
    df.to_csv(output_dir / "reservations_clean.csv", index=False)
    calendar.to_csv(output_dir / "calendar.csv", index=False)
    booking_nights.to_csv(output_dir / "booking_nights.csv", index=False)
    daily.to_csv(output_dir / "daily_occupancy.csv", index=False)

    df_targets.to_csv(output_dir / "targets_wide_clean.csv", index=False)
    expanded_targets.to_csv(output_dir / "targets_long.csv", index=False)
    daily_targets.to_csv(output_dir / "daily_targets.csv", index=False)

    final.to_csv(output_dir / "daily_performance.csv", index=False)

    print("Reservations loaded and validated successfully.")
    print(f"Outputs saved to: {output_dir}")

    df_expenses = load_expenses(input_path, sheet_name="Expenses")
    df_expanded_expenses = expand_expenses(df_expenses)
    df_expanded_daily_expenses = expand_daily_expenses(df_expanded_expenses)
    df_expanded_daily_expenses.to_csv(output_dir / "expenses_long.csv", index=False)
    
    summary = build_weekly_summary(df)
    print(summary["num_arrivals"])


    if mode == "weekly-email":
        summary = build_weekly_arrivals_departures_summary(df_reservations)
        html = build_weekly_email_html(summary)

        subject = (
            f"Weekly Reservations Summary "
            f"{summary['start_of_week'].strftime('%d %b')} – "
            f"{summary['end_of_week'].strftime('%d %b')}"
        )

        send_html_email(subject, html)
        print("Weekly email sent.")

if __name__ == "__main__":
    main()