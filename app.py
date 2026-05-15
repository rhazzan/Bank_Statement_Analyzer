# ── Standard library ──────────────────────────────────────────────────────────
import io
import tempfile
from typing import Optional

# ── Third-party ───────────────────────────────────────────────────────────────
import pandas as pd
import streamlit as st

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════

MAIN_HEADER_ROW: int = 6          # 0-indexed row where column headers live
SAVINGS_SHEET: str = "Savings Account Transactions"

# Known possible names for the main transactions sheet (OPay uses "Wallet Account Transactions")
MAIN_SHEET_CANDIDATES: list[str] = [
    "Wallet Account Transactions",
    "Current Account Transactions",
    "Account Transactions",
    "Transactions",
]
REQUIRED_MAIN_COLS: list[str] = [
    "Trans. Date", "Description", "Debit(₦)", "Credit(₦)",
    "Balance After(₦)", "Channel", "Transaction Reference",
]

REQUIRED_SAVINGS_COLS: list[str] = [
    "Trans. Date", "Description", "Debit(₦)", "Credit(₦)",
    "Balance After(₦)", "Channel", "Transaction Reference",
]

MONTH_ORDER: list[str] = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

# ══════════════════════════════════════════════════════════════════════════════
# HELPER UTILITIES
# ══════════════════════════════════════════════════════════════════════════════

def safe_numeric(series: pd.Series) -> pd.Series:
    """
    Safely convert a Series to float.
    Replaces '--', commas, then coerces; NaN → 0.
    """
    return (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("--", "0", regex=False)
        .pipe(pd.to_numeric, errors="coerce")
        .fillna(0)
    )


def extract_name(desc: str) -> str:
    """
    Pull a human-readable name from a pipe-delimited bank description.
    Falls back gracefully when the expected delimiters are absent.
    """
    desc = str(desc)
    lower = desc.lower()
    try:
        if "from" in lower:
            name = lower.split("from")[1].split("|")[0].strip()
        elif "to" in lower:
            name = lower.split("to")[1].split("|")[0].strip()
        else:
            name = desc.split("|")[0].strip()
        return name.title() if name else desc.split("|")[0].strip().title()
    except Exception:
        return desc.split("|")[0].strip().title()


def fix_platform_swap(row: pd.Series) -> pd.Series:
    """
    Some rows have Platform and Account/Phone columns swapped.
    Detect by checking if Platform is all digits (should be Account/Phone).
    """
    platform = str(row.get("Platform", "")).strip()
    account  = str(row.get("Account/Phone", "")).strip()
    if platform.replace(" ", "").isdigit() and any(c.isalpha() for c in account):
        row["Platform"], row["Account/Phone"] = account, platform
    return row


def add_percentage_to_amount_table(df: pd.DataFrame, amount_col: str = "Amount") -> pd.DataFrame:
    """Append a '% of Total' column to any single-amount summary table."""
    out = df.copy()
    total = out[amount_col].sum()
    out["% of Total"] = (out[amount_col] / total * 100).round(2) if total else 0
    return out


def add_percentage_columns(
    pivot_df: pd.DataFrame,
    debit_col: str = "Debit(₦)",
    credit_col: str = "Credit(₦)",
) -> pd.DataFrame:
    """Add % Debit, % Credit, % Total Flow columns to a pivot table."""
    out = pivot_df.copy()
    if debit_col not in out.columns:   # ✅ correct DataFrame way
        out[debit_col] = 0
    if credit_col not in out.columns:
        out[credit_col] = 0

    total_debit  = out[debit_col].sum()
    total_credit = out[credit_col].sum()
    total_flow   = total_debit + total_credit

    out["% Debit"]      = (out[debit_col]  / total_debit  * 100).round(2) if total_debit  else 0
    out["% Credit"]     = (out[credit_col] / total_credit * 100).round(2) if total_credit else 0
    out["% Total Flow"] = ((out[debit_col] + out[credit_col]) / total_flow * 100).round(2) if total_flow else 0
    return out


# ══════════════════════════════════════════════════════════════════════════════
# FILE LOADING  (cached so re-runs from widget interaction are free)
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def load_file(file_bytes: bytes) -> tuple[pd.DataFrame, bool, Optional[pd.DataFrame]]:
    buf = io.BytesIO(file_bytes)
    xls = pd.ExcelFile(buf)

    # ── Detect main sheet ─────────────────────────────────────────────────────
    # OPay exports "Wallet Account Transactions"; other banks may differ.
    # Try known candidates, then fall back to the first non-savings sheet.
    main_sheet = None
    for candidate in MAIN_SHEET_CANDIDATES:
        if candidate in xls.sheet_names:
            main_sheet = candidate
            break
    if main_sheet is None:
        # Fall back: first sheet that is NOT the savings sheet
        non_savings = [s for s in xls.sheet_names if s != SAVINGS_SHEET]
        if not non_savings:
            raise ValueError(
                f"Cannot find a main transactions sheet. "
                f"Sheets found: {xls.sheet_names}"
            )
        main_sheet = non_savings[0]

    buf.seek(0)
    main_df = pd.read_excel(buf, sheet_name=main_sheet, header=MAIN_HEADER_ROW)

    # ── Optional savings sheet ────────────────────────────────────────────────
    savings_df = None
    savings_exists = SAVINGS_SHEET in xls.sheet_names
    if savings_exists:
        buf.seek(0)
        try:
            savings_df = pd.read_excel(buf, sheet_name=SAVINGS_SHEET, header=MAIN_HEADER_ROW)
        except Exception:
            savings_exists = False

    return main_df, savings_exists, savings_df


# ══════════════════════════════════════════════════════════════════════════════
# DATA CLEANING
# ══════════════════════════════════════════════════════════════════════════════

def clean_transactions(raw_df: pd.DataFrame) -> pd.DataFrame:
    """
    Full cleaning pipeline for the main transactions sheet.
    Every step is wrapped defensively so a single bad value never crashes
    the whole pipeline.

    FIX: date conversion now uses errors='coerce' so malformed dates
    become NaT rather than raising an exception.
    """
    df = raw_df.copy()

    # ── Validate required columns ─────────────────────────────────────────────
    missing = [c for c in REQUIRED_MAIN_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Main sheet is missing columns: {missing}")

    # ── Date parsing ──────────────────────────────────────────────────────────
    # FIX: was `format='%d %b %Y %H:%M:%S'` which crashes on ANY mismatch.
    # Now uses errors='coerce' so bad rows become NaT instead of exploding.
    df["Trans. Date1"] = pd.to_datetime(df["Trans. Date"], errors="coerce")
    bad_dates = df["Trans. Date1"].isna().sum()
    if bad_dates:
        st.warning(f"⚠️ {bad_dates} row(s) had unparseable dates and will be excluded.")
    df = df[df["Trans. Date1"].notna()].copy()

    df["Trans. Date"] = df["Trans. Date1"].dt.date
    df["Time"]        = df["Trans. Date1"].dt.time

    # ── Name extraction ───────────────────────────────────────────────────────
    df["Transaction Name"] = df["Description"].apply(extract_name)

    # ── Numeric columns ───────────────────────────────────────────────────────
    df["Debit(₦)"]  = safe_numeric(df["Debit(₦)"])
    df["Credit(₦)"] = safe_numeric(df["Credit(₦)"])

    # ── Transaction type + unified amount ─────────────────────────────────────
    df["Transaction Type"] = df.apply(
        lambda r: "Debit(₦)" if r["Debit(₦)"] > 0 else "Credit(₦)", axis=1
    )
    df["Amount"] = df.apply(
        lambda r: r["Debit(₦)"] if r["Debit(₦)"] > 0 else r["Credit(₦)"], axis=1
    )
    df = df.drop(columns=["Debit(₦)", "Credit(₦)"], errors="ignore")

    # ── Description split ─────────────────────────────────────────────────────
    desc_splits = df["Description"].str.split("|", expand=True)
    desc_cols   = ["Transaction To/From", "Platform", "Account/Phone", "Extra Info"]
    # Only assign as many column names as we actually have
    desc_splits.columns = desc_cols[: desc_splits.shape[1]]
    # Fill any missing expected columns with empty string
    for col in desc_cols:
        if col not in desc_splits.columns:
            desc_splits[col] = ""

    desc_splits = desc_splits.apply(fix_platform_swap, axis=1)
    df = pd.concat([df.reset_index(drop=True), desc_splits.reset_index(drop=True)], axis=1)

    # ── Drop Value Date if it exists ──────────────────────────────────────────
    df = df.drop(columns=["Value Date", "Trans. Date1"], errors="ignore")

    # ── Reorder columns (only keep cols that actually exist) ──────────────────
    desired_order = [
        "Transaction Reference", "Trans. Date", "Time",
        "Transaction Type", "Transaction To/From", "Transaction Name",
        "Account/Phone", "Platform", "Channel", "Extra Info",
        "Amount", "Balance After(₦)",
    ]
    existing_cols = [c for c in desired_order if c in df.columns]
    df = df[existing_cols]

  # ── Final numeric safety pass ─────────────────────────────────────────────
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
    # Balance column arrives as comma-formatted strings e.g. "48,144.63"
    if "Balance After(₦)" in df.columns:
        df["Balance After(₦)"] = safe_numeric(df["Balance After(₦)"])
    return df


def process_savings_sheet(raw_savings: pd.DataFrame) -> Optional[dict]:
    """
    Clean the savings sheet and compute summary tables.

    Returns a dict with DataFrames ready for Excel output, or None if
    the sheet is empty or missing required columns.
    """
    df_s = raw_savings.copy()
    df_s.columns = df_s.columns.str.strip()

    # Drop Value Date if present
    df_s = df_s.drop(columns=["Value Date"], errors="ignore")

    missing = [c for c in REQUIRED_SAVINGS_COLS if c not in df_s.columns]
    if missing:
        st.warning(f"Savings sheet missing columns {missing} — savings analysis skipped.")
        return None

    # Date
    df_s["Trans. Date"] = pd.to_datetime(df_s["Trans. Date"], errors="coerce")
    df_s = df_s[df_s["Trans. Date"].notna()].copy()

    if df_s.empty:
        st.info("Savings sheet exists but has no valid rows after date cleaning.")
        return None

    # Numeric
    for col in ["Debit(₦)", "Credit(₦)", "Balance After(₦)"]:
        df_s[col] = safe_numeric(df_s[col])

    # Interest rows
    interest_mask = df_s["Description"].str.contains("Interest", case=False, na=False)
    interest_df   = df_s[interest_mask]
    total_interest = interest_df["Credit(₦)"].sum()

    latest_balance = df_s.sort_values("Trans. Date").iloc[-1]["Balance After(₦)"]

    summary_df = pd.DataFrame({
        "Metric": ["Total Interest Earned"],
        "Value":  [round(total_interest, 2)],
    })
    balance_df = pd.DataFrame({
        "Metric": ["Latest Savings Balance"],
        "Value":  [round(latest_balance, 2)],
    })
    interest_by_type = (
        interest_df.groupby("Description")["Credit(₦)"].sum().reset_index()
    )
    balance_by_type = (
        df_s.groupby("Description")["Balance After(₦)"].max().reset_index()
    )

    return {
        "summary":          summary_df,
        "balance":          balance_df,
        "interest_by_type": interest_by_type,
        "balance_by_type":  balance_by_type,
    }


# ══════════════════════════════════════════════════════════════════════════════
# ANALYSIS GENERATION
# ══════════════════════════════════════════════════════════════════════════════

def generate_analysis(df: pd.DataFrame, top_n: int = 10) -> dict:
    """
    Derive all analytical summary tables from the cleaned dataframe.
    Pure function — no Streamlit calls, no file I/O.
    """
    df = df.copy()
    df["Trans. Date"] = pd.to_datetime(df["Trans. Date"], errors="coerce")
    df["Month"]       = df["Trans. Date"].dt.month_name()
    df["Date_Only"]   = df["Trans. Date"].dt.date

    is_debit  = df["Transaction Type"] == "Debit(₦)"
    is_credit = df["Transaction Type"] == "Credit(₦)"

    total_debit        = df.loc[is_debit,  "Amount"].sum()
    total_credit       = df.loc[is_credit, "Amount"].sum()
    total_debit_count  = is_debit.sum()
    total_credit_count = is_credit.sum()
    latest_balance     = df.sort_values("Trans. Date").iloc[-1]["Balance After(₦)"]

    summary_df = pd.DataFrame({
        "Metric": [
            "Total Debit Amount (₦)",
            "Total Credit Amount (₦)",
            "Number of Debit Transactions",
            "Number of Credit Transactions",
            "Current Balance (₦)",
        ],
        "Value": [
            total_debit, total_credit,
            total_debit_count, total_credit_count,
            latest_balance,
        ],
    })

    # Monthly
    monthly = df.pivot_table(
        index="Month", columns="Transaction Type",
        values="Amount", aggfunc="sum", fill_value=0,
    )
    monthly = add_percentage_columns(monthly)
    monthly = monthly.reindex(MONTH_ORDER).dropna(how="all")

    # Platform
    platform = df.pivot_table(
        index="Platform", columns="Transaction Type",
        values="Amount", aggfunc="sum", fill_value=0,
    )
    platform = add_percentage_columns(platform)

    # Daily
    daily = df.pivot_table(
        index="Date_Only", columns="Transaction Type",
        values="Amount", aggfunc="sum", fill_value=0,
    ).sort_index()

    # Top spenders / income
    top_spending = (
        df[is_debit]
        .groupby("Transaction To/From")["Amount"].sum()
        .sort_values(ascending=False)
        .head(top_n)
        .reset_index()
    )
    top_spending = add_percentage_to_amount_table(top_spending)

    top_income = (
        df[is_credit]
        .groupby("Transaction To/From")["Amount"].sum()
        .sort_values(ascending=False)
        .head(top_n)
        .reset_index()
    )
    top_income = add_percentage_to_amount_table(top_income)

    return {
        "summary":        summary_df,
        "monthly":        monthly,
        "platform":       platform,
        "daily":          daily,
        "top_spending":   top_spending,
        "top_income":     top_income,
        "total_debit":    total_debit,
        "total_credit":   total_credit,
        "latest_balance": latest_balance,
        "debit_count":    total_debit_count,
        "credit_count":   total_credit_count,
    }


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL REPORT WRITER
# ══════════════════════════════════════════════════════════════════════════════

def write_excel_report(
    df: pd.DataFrame,
    analysis: dict,
    savings_data: Optional[dict],
) -> bytes:
    """
    Write the cleaned data + all analysis tables to an Excel workbook
    held entirely in memory (BytesIO).  Returns raw bytes.

    FIX: was writing to `uploaded_file` directly (invalid), and used
    mode='a' before the file even existed.  Now uses BytesIO throughout.
    """
    buf = io.BytesIO()

    try:
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:

            # ── Cleaned data ──────────────────────────────────────────────────
            df.to_excel(writer, sheet_name="Cleaned_Data", index=False)

            # ── Analysis sheet ────────────────────────────────────────────────
            row = 0

            def write_section(title: str, data: pd.DataFrame, reset_index: bool = False) -> int:
                nonlocal row
                pd.DataFrame({title: []}).to_excel(
                    writer, sheet_name="Analysis", startrow=row, index=False
                )
                row += 1
                data.to_excel(writer, sheet_name="Analysis", startrow=row, index=reset_index)
                row += len(data) + 3
                return row

            write_section("OVERALL FINANCIAL SUMMARY",   analysis["summary"],      reset_index=False)
            write_section("MONTHLY CASH FLOW SUMMARY",   analysis["monthly"],      reset_index=True)
            write_section("PLATFORM PERFORMANCE SUMMARY",analysis["platform"],     reset_index=True)
            write_section("DAILY TRANSACTION TREND",     analysis["daily"],        reset_index=True)
            write_section("TOP SPENDING RECIPIENTS",     analysis["top_spending"], reset_index=False)
            write_section("TOP INCOME SOURCES",          analysis["top_income"],   reset_index=False)

            # ── Savings sheet (optional) ──────────────────────────────────────
            if savings_data:
                sd = savings_data
                sd["summary"].insert(0, "Section", "TOTAL INTEREST")
                sd["balance"].insert(0, "Section", "LATEST SAVINGS BALANCE")
                sd["interest_by_type"].insert(0, "Section", "INTEREST BY SAVINGS TYPE")
                sd["balance_by_type"].insert(0, "Section", "BALANCE BY SAVINGS TYPE")

                final = pd.concat(
                    [
                        sd["summary"],
                        pd.DataFrame([[]]),
                        sd["balance"],
                        pd.DataFrame([[]]),
                        sd["interest_by_type"],
                        pd.DataFrame([[]]),
                        sd["balance_by_type"],
                    ],
                    ignore_index=True,
                )
                final.to_excel(writer, sheet_name="Savings_Analysis", index=False)

    except Exception as exc:
        st.error(f"❌ Excel report generation failed: {exc}")
        return b""

    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# STREAMLIT TAB RENDERERS
# ══════════════════════════════════════════════════════════════════════════════

def show_dashboard(df: pd.DataFrame, analysis: dict, top_n: int) -> None:
    """Render all charts, KPIs, and tables inside the Dashboard tab."""

    st.header("📊 Financial Dashboard")

    # ── KPI row ───────────────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("💸 Total Debits",   f"₦{analysis['total_debit']:,.2f}")
    c2.metric("💰 Total Credits",  f"₦{analysis['total_credit']:,.2f}")
    c3.metric("🏦 Balance",        f"₦{analysis['latest_balance']:,.2f}")
    c4.metric("🔢 Transactions",   analysis["debit_count"] + analysis["credit_count"])

    st.divider()

    # ── Monthly cash flow ─────────────────────────────────────────────────────
    st.subheader("📅 Monthly Cash Flow")
    monthly = analysis["monthly"].copy()
    chart_cols = [c for c in ["Debit(₦)", "Credit(₦)"] if c in monthly.columns]
    if chart_cols:
        st.bar_chart(monthly[chart_cols])
    st.dataframe(monthly, use_container_width=True)

    st.divider()

    # ── Platform performance ──────────────────────────────────────────────────
    st.subheader("🏦 Platform Performance")
    st.dataframe(analysis["platform"], use_container_width=True)

    st.divider()

    # ── Daily trend ───────────────────────────────────────────────────────────
    st.subheader("📈 Daily Transaction Trend")
    daily = analysis["daily"].copy()
    daily_chart_cols = [c for c in ["Debit(₦)", "Credit(₦)"] if c in daily.columns]
    if daily_chart_cols:
        st.line_chart(daily[daily_chart_cols])

    st.divider()

    # ── Top spenders / income ─────────────────────────────────────────────────
    col_left, col_right = st.columns(2)
    with col_left:
        st.subheader(f"🏆 Top {top_n} Spending Recipients")
        st.dataframe(analysis["top_spending"], use_container_width=True)
    with col_right:
        st.subheader(f"🏆 Top {top_n} Income Sources")
        st.dataframe(analysis["top_income"], use_container_width=True)

    st.divider()

    # ── Overall summary ───────────────────────────────────────────────────────
    st.subheader("📋 Overall Financial Summary")
    st.dataframe(analysis["summary"], use_container_width=True)


def show_dataset(df: pd.DataFrame) -> None:
    """
    Render the Dataset tab.

    FIX: Previously `df` was defined inside `with tab1:` so it was out of
    scope here.  Now the cleaned df is passed in explicitly (and also stored
    in st.session_state), so this tab always has access to it.
    """
    st.header("🗂️ Cleaned Dataset")

    if df is None or df.empty:
        st.warning("No data to display. Please upload and process a file first.")
        return

    st.success(f"✅ {len(df):,} transactions loaded")

    # Quick filters in the sidebar (already set before this call, passed in)
    st.dataframe(df, use_container_width=True, height=600)

    st.divider()
    st.subheader("📥 Download Cleaned Data")
    csv_bytes = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="⬇️ Download as CSV",
        data=csv_bytes,
        file_name="cleaned_transactions.csv",
        mime="text/csv",
    )


# ══════════════════════════════════════════════════════════════════════════════
# MAIN APP ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    """
    Application entry point.

    Architecture
    ------------
    1. File uploaded  → bytes cached in session_state (avoids repeated reads)
    2. "Process Data" pressed → load_file + clean_transactions run once;
       result stored in session_state["df"]
    3. All tabs read ONLY from session_state → zero scope issues
    """

    # ── Page config ───────────────────────────────────────────────────────────
    st.set_page_config(
        page_title="Nigerian Bank Analyser",
        page_icon="🏦",
        layout="wide",
    )

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.title("🏦 Bank Statement Analyser")
        st.markdown("Upload your Nigerian bank statement Excel file to get started.")
        st.divider()
        top_n = st.number_input(
            "Top N recipients / sources", min_value=1, max_value=50, value=10, step=1
        )
        st.divider()
        st.caption("Supports:OPay Only")

    # ── Header ────────────────────────────────────────────────────────────────
    st.title("🏦 Nigerian Bank Statement Analyser")

    # ── File uploader ─────────────────────────────────────────────────────────
    uploaded_file = st.file_uploader(
        "Upload your bank statement (.xlsx)",
        type=["xlsx", "xls"],
        help="Excel file exported from your bank's internet banking portal.",
    )

    # ── Process button ────────────────────────────────────────────────────────
    process_disabled = uploaded_file is None
    if st.button("⚙️ Process Data", disabled=process_disabled, type="primary"):
        if uploaded_file is not None:
            # Read bytes once; re-use for everything downstream
            file_bytes = uploaded_file.read()

            with st.spinner("Loading workbook…"):
                try:
                    raw_df, savings_exists, raw_savings = load_file(file_bytes)
                except Exception as exc:
                    st.error(f"❌ Could not open the file: {exc}")
                    st.stop()

            with st.spinner("Cleaning transactions…"):
                try:
                    df = clean_transactions(raw_df)
                except ValueError as exc:
                    st.error(f"❌ Data validation error: {exc}")
                    st.stop()
                except Exception as exc:
                    st.error(f"❌ Unexpected cleaning error: {exc}")
                    st.stop()

            # Savings
            savings_data: Optional[dict] = None
            if savings_exists and raw_savings is not None:
                with st.spinner("Processing savings sheet…"):
                    try:
                        savings_data = process_savings_sheet(raw_savings)
                    except Exception as exc:
                        st.warning(f"⚠️ Savings sheet skipped: {exc}")

            # Analysis
            with st.spinner("Generating analysis…"):
                try:
                    analysis = generate_analysis(df, top_n=int(top_n))
                except Exception as exc:
                    st.error(f"❌ Analysis failed: {exc}")
                    st.stop()

            # Excel report
            with st.spinner("Writing Excel report…"):
                report_bytes = write_excel_report(df, analysis, savings_data)

            # ── Store everything in session_state ─────────────────────────────
            # FIX: this is the key change — df is now globally accessible
            # to any tab via st.session_state, not buried in a local scope.
            st.session_state["df"]           = df
            st.session_state["analysis"]     = analysis
            st.session_state["savings_data"] = savings_data
            st.session_state["report_bytes"] = report_bytes
            st.session_state["processed"]    = True

            st.success("✅ File processed successfully!")

            # Download button for the report (shown immediately after processing)
            if report_bytes:
                st.download_button(
                    label="📥 Download Full Excel Report",
                    data=report_bytes,
                    file_name="bank_analysis_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    # ── Tabs (created ONCE, unconditionally) ──────────────────────────────────
    # FIX: previously tabs were created twice (before and after the button),
    # producing two separate widget trees.  Now there is exactly one set.
    tab_dashboard, tab_dataset = st.tabs(["📊 Dashboard", "🗂️ Dataset"])

    # Pull data from session_state (safe even before first processing)
    df_ready       = st.session_state.get("processed", False)
    df             = st.session_state.get("df")
    analysis       = st.session_state.get("analysis")

    with tab_dashboard:
        if not df_ready:
            st.info("👆 Upload a file and click **Process Data** to see your dashboard.")
        else:
            try:
                show_dashboard(df, analysis, top_n=int(top_n))
            except Exception as exc:
                st.error(f"❌ Dashboard render error: {exc}")

    with tab_dataset:
        if not df_ready:
            st.info("👆 Upload a file and click **Process Data** to see the dataset.")
        else:
            try:
                # FIX: df comes from session_state — always in scope here
                show_dataset(df)
            except Exception as exc:
                st.error(f"❌ Dataset render error: {exc}")


# ── Run ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    main()
