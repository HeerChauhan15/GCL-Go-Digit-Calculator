import streamlit as st
import pandas as pd
import numpy as np
import os
import re

st.set_page_config(
    page_title="GCL Go Digit Calculator",
    page_icon="💰",
    layout="wide"
)

st.title("💰 GCL Go Digit Calculator")
st.markdown("Select plan details below")

GST_RATE = 0.18

# ============================================
# AGE LIMITS (ALL SEGMENTS)
# ============================================
AGE_MIN = 18
AGE_MAX = 60

# ============================================
# FILE MAP: (Segment, Type of Cover) -> filename
# ============================================
FILE_MAP = {
    ("Home Loan", "Level"): "Home Loan Level.xlsx",
    ("Home Loan", "Reducing"): "Home Loan Reducing.xlsx",
    ("Loan Against Property", "Level"): "LAP Level.xlsx",
    ("Loan Against Property", "Reducing"): "LAP Reducing.xlsx",
    ("Micro Loan", "Level"): "Micro Loan Level.xlsx",
    ("Micro Loan", "Reducing"): "Micro Loan Reducing.xlsx",
    ("Personal Loan", "Level"): "Personal Loan Level.xlsx",
    ("Personal Loan", "Reducing"): "Personal Loan Reducing.xlsx",
    ("Secured Loan", "Level"): "Secured Loan Business Level.xlsx",
    ("Secured Loan", "Reducing"): "Secured Loan Business Reducing.xlsx",
    ("Unsecured Loan", "Level"): "Unsecured Loan Business Level.xlsx",
    ("Unsecured Loan", "Reducing"): "Unsecured Loan Business Reducing.xlsx",
}

# ============================================
# SUM ASSURED & TENURE LIMITS PER SEGMENT
# ============================================
SEGMENT_LIMITS = {
    "Home Loan": {"sa_min": 500000, "sa_max": 6000000, "t_min": 5, "t_max": 25},
    "Loan Against Property": {"sa_min": 100000, "sa_max": 4000000, "t_min": 1, "t_max": 10},
    "Micro Loan": {"sa_min": 10000, "sa_max": 200000, "t_min": 1, "t_max": 3},
    "Personal Loan": {"sa_min": 100000, "sa_max": 1500000, "t_min": 1, "t_max": 10},
    "Secured Loan": {"sa_min": 100000, "sa_max": 1000000, "t_min": 1, "t_max": 10},
    "Unsecured Loan": {"sa_min": 100000, "sa_max": 1500000, "t_min": 1, "t_max": 10},
}

SEGMENT_OPTIONS = list(SEGMENT_LIMITS.keys())
COVER_OPTIONS = ["Level", "Reducing"]

# ============================================
# RATE TABLE LOADER
# ============================================
def load_rate_table(segment, cover_type):
    fname = FILE_MAP[(segment, cover_type)]
    if not os.path.exists(fname):
        raise FileNotFoundError(
            f"File not found: '{fname}' — Please make sure this file is in the GitHub repo."
        )

    raw = pd.read_excel(fname, sheet_name="Sheet1", header=None)

    header_row = None
    for i, row in raw.iterrows():
        for val in row.values:
            if isinstance(val, str) and "AGE" in val.upper():
                header_row = i
                break
        if header_row is not None:
            break

    if header_row is None:
        raise ValueError("Could not find AGE/TERM header row in the file.")

    df = pd.read_excel(fname, sheet_name="Sheet1", header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    age_col = df.columns[0]
    df = df.dropna(subset=[age_col])
    df[age_col] = pd.to_numeric(df[age_col], errors='coerce')
    df = df.dropna(subset=[age_col])
    df[age_col] = df[age_col].astype(int)
    df = df.set_index(age_col)

    tenure_map = {}
    for col in df.columns:
        try:
            tenure_map[int(float(col))] = col
        except Exception:
            pass

    return df, tenure_map


def get_rate(df, tenure_map, age, tenure_years):
    """User enters tenure in YEARS, but the rate table's columns are in MONTHS.
    Convert years -> months before looking up the column."""
    if age not in df.index:
        raise ValueError(f"Age {age} not found in rate table.")
    tenure_months = int(round(tenure_years * 12))
    if tenure_months not in tenure_map:
        raise ValueError(
            f"Tenure {tenure_years} yrs ({tenure_months} months) not found in rate table."
        )
    return float(df.loc[age, tenure_map[tenure_months]])


def compute_premium_breakup(rate, sum_assured):
    """Rate table stores GROSS premium rate (per Rs 1,00,000 SA).
    Returns (net_premium, gst_amount, gross_premium)."""
    gross_premium = rate * (sum_assured / 100000)
    net_premium = gross_premium / (1 + GST_RATE)
    gst_amount = gross_premium - net_premium
    return round(net_premium, 2), round(gst_amount, 2), round(gross_premium, 2)


def find_column(df, target):
    """Exact (case/space-insensitive) match."""
    target_norm = target.strip().lower().replace(" ", "")
    for col in df.columns:
        col_norm = str(col).strip().lower().replace(" ", "")
        if col_norm == target_norm:
            return col
    return None


def find_flexible_column(df, keywords):
    """Flexible detection — matches if any keyword (normalized) appears in the column name."""
    for col in df.columns:
        norm = re.sub(r'[\s_-]+', '', str(col).lower())
        for kw in keywords:
            if kw in norm:
                return col
    return None


# ============================================
# DROPDOWNS
# ============================================
col1, col2 = st.columns(2)
with col1:
    segment = st.selectbox("Segment", SEGMENT_OPTIONS)
with col2:
    cover_type = st.selectbox("Type of Cover", COVER_OPTIONS)

limits = SEGMENT_LIMITS[segment]
sa_min, sa_max = limits["sa_min"], limits["sa_max"]
min_tenure, max_tenure = limits["t_min"], limits["t_max"]

# ============================================
# SUM ASSURED (mandatory)
# ============================================
sum_assured_input = st.number_input(
    "Select Sum Assured (₹)",
    min_value=0,
    value=sa_min,
    step=10000,
    help=f"For {segment}, Sum Assured must be between ₹{sa_min:,} and ₹{sa_max:,}."
)

if sum_assured_input < sa_min:
    st.warning(
        f"⚠️ Minimum Sum Assured for {segment} is ₹{sa_min:,}. "
        f"Value adjusted to ₹{sa_min:,}."
    )
    sum_assured = sa_min
elif sum_assured_input > sa_max:
    st.warning(
        f"⚠️ Maximum Sum Assured for {segment} is ₹{sa_max:,}. "
        f"Value adjusted to ₹{sa_max:,}."
    )
    sum_assured = sa_max
else:
    sum_assured = sum_assured_input

st.divider()

# ============================================
# MANUAL SECTION
# ============================================
col3, col4 = st.columns(2)
with col3:
    age_input = st.number_input(
        "Enter Age",
        min_value=18,
        value=30,
        step=1
    )

    if age_input < AGE_MIN:
        st.warning(
            f"⚠️ Minimum Age for all segments is **{AGE_MIN} years**. "
            f"Value adjusted to **{AGE_MIN} years**."
        )
        age = AGE_MIN
    elif age_input > AGE_MAX:
        st.warning(
            f"⚠️ Maximum Age for all segments is **{AGE_MAX} years**. "
            f"Value adjusted to **{AGE_MAX} years**."
        )
        age = AGE_MAX
    else:
        age = age_input

with col4:
    tenure_input = st.number_input(
        "Enter Tenure",
        min_value=0,
        value=min_tenure,
        step=1
    )
    if tenure_input < min_tenure:
        st.warning(
            f"⚠️ Minimum Tenure for {segment} is {min_tenure} yrs. "
            f"Value adjusted to {min_tenure} yrs."
        )
        tenure = min_tenure
    elif tenure_input > max_tenure:
        st.warning(
            f"⚠️ Maximum Tenure for {segment} is {max_tenure} yrs. "
            f"Value adjusted to {max_tenure} yrs."
        )
        tenure = max_tenure
    else:
        tenure = tenure_input
    st.caption("📅 Tenure is in Years")

st.write("")
if st.button("Get Rate", type="primary", use_container_width=True):
    try:
        df_rates, tenure_map = load_rate_table(segment, cover_type)
        rate = get_rate(df_rates, tenure_map, age, tenure)
        net_premium, gst_amount, gross_premium = compute_premium_breakup(rate, sum_assured)

        st.success(
            f"✅ {segment} | {cover_type} | Age {age} | Tenure {tenure} yrs | "
            f"Sum Assured ₹{sum_assured:,}"
        )
        col_b, col_c, col_d = st.columns(3)
        with col_b:
            st.metric("Net Premium (excl. GST)", f"₹ {net_premium:,.2f}")
        with col_c:
            st.metric("GST (18%)", f"₹ {gst_amount:,.2f}")
        with col_d:
            st.metric("Gross Premium", f"₹ {gross_premium:,.2f}")
    except Exception as e:
        st.error(f"Error: {e}")

st.divider()

# ============================================
# EXCEL UPLOAD SECTION
# ============================================
st.subheader("📂 Upload Member Data for Bulk Rate Lookup")

st.markdown(
    "Your Excel must have at least: Sum Assured, Age, Tenure (in years). "
    "Any other columns you include will be carried through to the output unchanged."
)

st.warning("⚠️ Please make sure you have selected Segment and Type of Cover above before uploading your Excel file.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]

        st.subheader("Uploaded Data Preview")
        st.dataframe(df.head())

        df_rates, tenure_map = load_rate_table(segment, cover_type)

        age_col = find_column(df, "Age") or find_flexible_column(df, ["age"])
        tenure_col = find_column(df, "Tenure") or find_flexible_column(df, ["tenure"])
        sa_col = (
            find_column(df, "Sum Assured")
            or find_flexible_column(df, ["sumassured", "suminsured", "loanamount"])
        )

        missing = []
        if not age_col:
            missing.append("Age")
        if not tenure_col:
            missing.append("Tenure")
        if not sa_col:
            missing.append("Sum Assured")
        if missing:
            raise ValueError("Excel must contain mandatory columns: " + ", ".join(missing))

        df[age_col] = pd.to_numeric(df[age_col], errors='coerce')
        df[tenure_col] = pd.to_numeric(df[tenure_col], errors='coerce')
        df[sa_col] = pd.to_numeric(df[sa_col], errors='coerce')

        if df[tenure_col].dropna().median() > 30:
            st.info("ℹ️ Tenure values look like months — auto-converting to years.")
            df[tenure_col] = (df[tenure_col] / 12).round(0).astype('Int64')
        else:
            df[tenure_col] = df[tenure_col].round(0).astype('Int64')

        df[age_col] = df[age_col].round(0).astype('Int64')

        # ---- AGE VALIDATION (18-60) — NO CLIPPING, ROWS ARE LEFT BLANK ----
        age_invalid_mask = (df[age_col] < AGE_MIN) | (df[age_col] > AGE_MAX) | df[age_col].isna()
        age_invalid_count = int(age_invalid_mask.sum())
        if age_invalid_count > 0:
            st.warning(
                f"⚠️ {age_invalid_count} row(s) have Age outside the allowed range "
                f"({AGE_MIN}-{AGE_MAX} years). Premiums will not be calculated for these rows."
            )

        # ---- TENURE VALIDATION — NO CLIPPING, ROWS ARE LEFT BLANK ----
        tenure_invalid_mask = (df[tenure_col] < min_tenure) | (df[tenure_col] > max_tenure) | df[tenure_col].isna()
        tenure_invalid_count = int(tenure_invalid_mask.sum())
        if tenure_invalid_count > 0:
            st.warning(
                f"⚠️ {tenure_invalid_count} row(s) have Tenure outside the allowed range "
                f"({min_tenure}-{max_tenure} yrs for {segment}). Premiums will not be calculated for these rows."
            )

        df[sa_col] = df[sa_col].fillna(sum_assured)
        sa_out_of_range = ((df[sa_col] < sa_min) | (df[sa_col] > sa_max)).sum()
        if sa_out_of_range > 0:
            st.warning(
                f"⚠️ {sa_out_of_range} row(s) had Sum Assured outside the allowed range "
                f"(₹{sa_min:,}-₹{sa_max:,} for {segment}) and were adjusted to the nearest limit."
            )
        df[sa_col] = df[sa_col].clip(lower=sa_min, upper=sa_max)

        net_list, gst_list, gross_list, status_list = [], [], [], []
        for idx, row in df.iterrows():
            try:
                r_age_raw = row[age_col]
                r_tenure_raw = row[tenure_col]

                # Blank out rows with missing / out-of-range age
                if pd.isna(r_age_raw) or r_age_raw < AGE_MIN or r_age_raw > AGE_MAX:
                    net_list.append(None)
                    gst_list.append(None)
                    gross_list.append(None)
                    status_list.append(f"❌ Age must be between {AGE_MIN} and {AGE_MAX}")
                    continue

                # Blank out rows with missing / out-of-range tenure
                if pd.isna(r_tenure_raw) or r_tenure_raw < min_tenure or r_tenure_raw > max_tenure:
                    net_list.append(None)
                    gst_list.append(None)
                    gross_list.append(None)
                    status_list.append(f"❌ Tenure must be between {min_tenure} and {max_tenure} yrs")
                    continue

                r_age = int(r_age_raw)
                r_tenure = int(r_tenure_raw)
                r_sa = float(row[sa_col])

                rate = get_rate(df_rates, tenure_map, r_age, r_tenure)
                net_p, gst_a, gross_p = compute_premium_breakup(rate, r_sa)

                net_list.append(net_p)
                gst_list.append(gst_a)
                gross_list.append(gross_p)
                status_list.append("✅")

            except Exception as e:
                net_list.append(None)
                gst_list.append(None)
                gross_list.append(None)
                status_list.append(f"❌ {e}")

        df["Net Premium"] = net_list
        df["GST Amount"] = gst_list
        df["Gross Premium"] = gross_list
        df["Status"] = status_list

        core_cols = [sa_col, age_col, tenure_col, "Net Premium", "GST Amount", "Gross Premium", "Status"]
        extra_cols = [c for c in df.columns if c not in core_cols]
        df_out = df[core_cols + extra_cols]

        total_net = pd.to_numeric(pd.Series(net_list), errors='coerce').sum()
        total_gst = pd.to_numeric(pd.Series(gst_list), errors='coerce').sum()
        total_gross = pd.to_numeric(pd.Series(gross_list), errors='coerce').sum()

        col_x, col_y, col_z = st.columns(3)
        with col_x:
            st.metric("💰 Total Net Premium", f"₹ {total_net:,.2f}")
        with col_y:
            st.metric("💰 Total GST", f"₹ {total_gst:,.2f}")
        with col_z:
            st.metric("💰 Total Gross Premium", f"₹ {total_gross:,.2f}")

        st.subheader("Rate Lookup Output")
        st.dataframe(df_out, use_container_width=True)

        total_row = {c: "" for c in df_out.columns}
        total_row[sa_col] = "TOTAL"
        total_row["Net Premium"] = round(total_net, 2)
        total_row["GST Amount"] = round(total_gst, 2)
        total_row["Gross Premium"] = round(total_gross, 2)
        df_final = pd.concat([df_out, pd.DataFrame([total_row])], ignore_index=True)

        output_file = "Rate_Output.xlsx"
        df_final.to_excel(output_file, index=False)

        with open(output_file, "rb") as file:
            st.download_button(
                label="⬇ Download Output Excel",
                data=file,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
