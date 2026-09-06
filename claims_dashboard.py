"""Local Streamlit prototype. Run with --server.address 127.0.0.1.

Install: python -m pip install streamlit pandas openpyxl
This version has no hosted authentication. Use fictional data only.
"""
import io
import pandas as pd
import streamlit as st

TITLE = "Claims Risk Management Intelligent Dashboard"
AMOUNTS = ["CLAIM_AMT", "AMOUNT_AGREED", "REJECTED AMOUNT", "APPLICABLE_BENEFIT_LIMIT"]
REQUIRED = AMOUNTS + ["CLIENTS", "SCHEME_NAME", "CLAIM_TYPE", "CLAIM_NUMBER", "CLAIM_CURRECY", "TREATMENT_DATE_1", "CLAIM_PROCESSED_DATE"]


def clean_amounts(values):
    # Python split handles Unicode whitespace without backend-specific regex.
    text = values.map(lambda value: "".join(str(value).split()).replace(",", "."))
    return pd.to_numeric(text, errors="coerce")


def prepare(raw):
    data = raw.copy()
    data.columns = data.columns.astype(str).str.strip()
    missing = sorted(set(REQUIRED) - set(data.columns))
    if missing:
        raise ValueError("Missing columns: " + ", ".join(missing))
    data = data.dropna(how="all")
    for col in AMOUNTS:
        data[col] = clean_amounts(data[col])
    for col in ["TREATMENT_DATE_1", "CLAIM_PROCESSED_DATE"]:
        data[col] = pd.to_datetime(data[col], errors="coerce")
    for col in ["CLIENTS", "SCHEME_NAME", "CLAIM_TYPE", "CLAIM_CURRECY"]:
        data[col] = data[col].fillna("Unknown").astype(str).str.strip().replace("", "Unknown")
    data["Treatment month"] = data["TREATMENT_DATE_1"].dt.strftime("%Y-%m").fillna("Unknown")
    data["Processing lag (days)"] = (data["CLAIM_PROCESSED_DATE"] - data["TREATMENT_DATE_1"]).dt.days
    data["Negative payment"] = data["AMOUNT_AGREED"] < 0
    data["Above invoice"] = data["AMOUNT_AGREED"] > data["CLAIM_AMT"] + 0.01
    data["Above per-claim limit"] = data["AMOUNT_AGREED"] > data["APPLICABLE_BENEFIT_LIMIT"] + 0.01
    data["Amounts do not reconcile"] = (data["CLAIM_AMT"] - data["AMOUNT_AGREED"] - data["REJECTED AMOUNT"]).abs() > 0.01
    data["Invalid date sequence"] = data["Processing lag (days)"] < 0
    data["Missing or invalid value"] = data[AMOUNTS + ["TREATMENT_DATE_1", "CLAIM_PROCESSED_DATE", "CLAIM_NUMBER"]].isna().any(axis=1)
    rules = ["Negative payment", "Above invoice", "Above per-claim limit", "Amounts do not reconcile", "Invalid date sequence", "Missing or invalid value"]
    data["Checks"] = data[rules].apply(lambda row: "; ".join(row.index[row]), axis=1)
    return data, rules


PREMIUM_KEYS = ["Client", "Product", "Month", "Currency"]


def prepare_premiums(raw):
    result = raw.dropna(how="all").copy()
    result.columns = result.columns.astype(str).str.strip()
    required = PREMIUM_KEYS + ["Earned Premium", "Data Type"]
    missing = sorted(set(required) - set(result.columns))
    if missing:
        raise ValueError("Premiums sheet is missing: " + ", ".join(missing))
    for col in ["Client", "Product", "Currency", "Data Type"]:
        if result[col].isna().any():
            raise ValueError("Premiums: blank " + col)
        result[col] = result[col].astype(str).str.strip()
        if result[col].eq("").any():
            raise ValueError("Premiums: blank " + col)
    result["Currency"] = result["Currency"].str.upper()
    if not result["Data Type"].str.lower().eq("fictional").all():
        raise ValueError("This prototype requires Data Type = Fictional on every premium row.")
    def month(value):
        if pd.isna(value) or isinstance(value, (int, float)):
            raise ValueError("Use YYYY-MM or an Excel date for Month, not a number or blank.")
        return str(pd.Timestamp(value).to_period("M"))
    result["Month"] = result["Month"].map(month)
    result["Earned Premium"] = clean_amounts(result["Earned Premium"])
    if result["Earned Premium"].isna().any() or not result["Earned Premium"].map(lambda x: 0 <= x < float("inf")).all():
        raise ValueError("Earned Premium must contain valid, nonnegative amounts.")
    if result.duplicated(PREMIUM_KEYS).any():
        raise ValueError("Duplicate premium keys: use one row per client, product, month and currency.")
    if result.empty:
        raise ValueError("The Premiums sheet has no data rows.")
    return result[required]


def premiums_from_membership(members, first, last, currency):
    m = members.copy()
    m.columns = m.columns.astype(str).str.strip()
    required = ["4037", "Company name", "Product name", "annual premium", "Joining Date", "Finish Policy Date"]
    if not set(required).issubset(m.columns):
        raise ValueError("Annual premium conversion needs member ID, client, product, annual premium, Joining Date and Finish Policy Date.")
    ids = m["4037"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    if m["4037"].isna().any() or ids.eq("").any() or ids.duplicated().any():
        raise ValueError("Membership IDs must be present and unique to avoid counting premiums twice.")
    for col in ["Company name", "Product name"]:
        if m[col].isna().any():
            raise ValueError("Missing membership client or product.")
        m[col] = m[col].astype(str).str.strip()
        if m[col].eq("").any():
            raise ValueError("Blank membership client or product.")
    m["annual premium"] = clean_amounts(m["annual premium"])
    if not m["annual premium"].map(lambda v: pd.notna(v) and 0 <= v < float("inf")).all():
        raise ValueError("Annual premiums must be valid nonnegative amounts.")
    joined = pd.to_datetime(m["Joining Date"], errors="coerce").dt.normalize()
    finished = pd.to_datetime(m["Finish Policy Date"], errors="coerce").dt.normalize()
    if joined.isna().any() or finished.isna().any() or (finished < joined).any():
        raise ValueError("Membership policy dates are missing, invalid or reversed.")
    frames = []
    for month in pd.period_range(first, last, freq="M"):
        lower = joined.clip(lower=month.start_time)
        upper = finished.clip(upper=month.end_time.normalize())
        days = ((upper - lower).dt.days + 1).clip(lower=0, upper=month.days_in_month)
        values = m[["Company name", "Product name"]].copy()
        values["Earned Premium"] = m["annual premium"] / 12 * days / month.days_in_month
        values = values.groupby(["Company name", "Product name"], as_index=False)["Earned Premium"].sum()
        values = values.rename(columns={"Company name": "Client", "Product name": "Product"})
        values["Month"] = str(month)
        values["Currency"] = currency
        values["Data Type"] = "Fictional"
        frames.append(values)
    if not frames:
        raise ValueError("No valid months for annual premium conversion.")
    return prepare_premiums(pd.concat(frames, ignore_index=True))


def reserve_scenario(data, premiums, clients, products, start_month, valuation_month, currency, completion):
    start = pd.Period(start_month, freq="M").start_time
    valuation = pd.Period(valuation_month, freq="M").end_time
    if start > valuation:
        raise ValueError("The start month must not be after the valuation month.")
    if not 0 < completion <= 100:
        raise ValueError("Paid completion must be above 0% and at most 100%.")
    claims = data.rename(columns={"CLIENTS": "Client", "SCHEME_NAME": "Product", "Treatment month": "Month", "CLAIM_CURRECY": "Currency"}).copy()
    claims["Currency"] = claims["Currency"].str.upper()
    base = claims[claims["Client"].isin(clients) & claims["Product"].isin(products) & claims["Currency"].eq(currency)]
    scoped = base[base["TREATMENT_DATE_1"].between(start, valuation) & base["CLAIM_PROCESSED_DATE"].le(valuation)].copy()
    if scoped["AMOUNT_AGREED"].isna().any():
        raise ValueError("Correct invalid paid amounts in the selected scope first.")
    pre = premiums[premiums["Client"].isin(clients) & premiums["Product"].isin(products) & premiums["Currency"].eq(currency)]
    pairs = pd.concat([base[["Client", "Product"]], pre[["Client", "Product"]]]).drop_duplicates()
    if pairs.empty:
        raise ValueError("No client/product combinations match this selection.")
    months = pd.DataFrame({"Month": pd.period_range(start_month, valuation_month, freq="M").astype(str)})
    grid = pairs.merge(months, how="cross")
    grid["Currency"] = currency
    paid = scoped.groupby(PREMIUM_KEYS, as_index=False)["AMOUNT_AGREED"].sum().rename(columns={"AMOUNT_AGREED": "Paid claims"})
    result = grid.merge(pre, on=PREMIUM_KEYS, how="left", validate="one_to_one").merge(paid, on=PREMIUM_KEYS, how="left", validate="one_to_one")
    result["Paid claims"] = result["Paid claims"].fillna(0)
    if result["Paid claims"].lt(0).any():
        raise ValueError("Resolve negative net client/product/month payments before reserving.")
    result["Premium missing"] = result["Earned Premium"].isna()
    result["Illustrative ultimate claims"] = result["Paid claims"] / (completion / 100)
    result["Illustrative unpaid reserve"] = result["Illustrative ultimate claims"] - result["Paid claims"]
    denominator = result["Earned Premium"].replace(0, float("nan"))
    result["Paid loss ratio (%)"] = result["Paid claims"] / denominator * 100
    result["Ultimate loss ratio (%)"] = result["Illustrative ultimate claims"] / denominator * 100
    return result, scoped


def render_reserves(data, premium_uploads, members=None):
    st.subheader("Reserves & loss ratio · Excel premiums")
    st.warning("FICTIONAL PREMIUMS AND ASSUMED RESERVES — prototype scenario, not a booked reserve or calibrated actuarial estimate.")
    st.caption("Use this tab's client, product, currency and period controls. Sidebar filters apply to other views. Loss ratios include all claim types within the scope selected here.")
    sources = []
    for label, upload in premium_uploads:
        if upload is not None:
            try:
                workbook = pd.ExcelFile(io.BytesIO(upload.getvalue()))
                if "Premiums" in workbook.sheet_names:
                    sources.append((label, upload))
            except Exception:
                st.error("Could not inspect " + label + " for a Premiums sheet.")
                return
    annual_available = members is not None and "annual premium" in members.columns
    labels = [x[0] for x in sources]
    if annual_available:
        labels.append("Membership annual premiums")
    if not labels:
        st.info("Add a Premiums sheet to a workbook, or an annual premium column to the membership workbook, then select the updated file again.")
        st.code("Client | Product | Month | Earned Premium | Currency | Data Type", language=None)
        st.write("For a Premiums sheet, use one row per client/product/month/currency and set Data Type to Fictional. Missing premium rows are not assumed to be zero.")
        return
    source = st.selectbox("Premium source", labels, key="premium_source")
    try:
        if source == "Membership annual premiums":
            currencies = sorted(set(data["CLAIM_CURRECY"].str.upper()))
            assumed_currency = st.selectbox("Annual premium currency · assumption", currencies, index=currencies.index("USD") if "USD" in currencies else 0)
            st.info("Fictional assumption: annual premium is a per-member annual rate in the selected currency. Monthly earned premium = annual premium ÷ 12 × covered days ÷ days in month. Joining and policy-end dates are inclusive. No rate changes, refunds or additional terminations are modelled.")
            first_month = data["TREATMENT_DATE_1"].min().to_period("M")
            last_month = data["CLAIM_PROCESSED_DATE"].max().to_period("M")
            premiums = premiums_from_membership(members, first_month, last_month, assumed_currency)
        else:
            upload = dict(sources)[source]
            premiums = prepare_premiums(pd.read_excel(io.BytesIO(upload.getvalue()), sheet_name="Premiums"))
    except Exception as exc:
        st.error("Premiums could not be used: " + str(exc))
        return
    dates = data["TREATMENT_DATE_1"].dropna()
    processed = data["CLAIM_PROCESSED_DATE"].dropna()
    if dates.empty or processed.empty:
        st.error("Valid treatment and processing dates are required.")
        return
    first = min(dates.min().to_period("M"), pd.Period(premiums["Month"].min(), freq="M"))
    last = processed.max().to_period("M")
    if first > last:
        st.error("No valid scenario period is available.")
        return
    months = pd.period_range(first, last, freq="M").astype(str).tolist()
    a, b, c = st.columns(3)
    start = a.selectbox("Treatment period starts", months, key="reserve_start")
    end = b.selectbox("Valuation month (month-end)", months, index=len(months)-1, key="reserve_end")
    currencies = sorted(set(data["CLAIM_CURRECY"].str.upper()) | set(premiums["Currency"]))
    currency = c.selectbox("Scenario currency", currencies, key="reserve_currency")
    d = data[data["CLAIM_CURRECY"].str.upper().eq(currency)]
    pr = premiums[premiums["Currency"].eq(currency)]
    choices = sorted(set(d["CLIENTS"]) | set(pr["Client"]))
    clients = st.multiselect("Scenario clients", choices, default=choices, key="reserve_clients")
    choices_p = sorted(set(d.loc[d["CLIENTS"].isin(clients), "SCHEME_NAME"]) | set(pr.loc[pr["Client"].isin(clients), "Product"]))
    products = st.multiselect("Scenario products", choices_p, default=choices_p, key="reserve_products")
    completion = st.number_input("Assumed paid completion (%) · fictional reserve assumption", min_value=1.0, max_value=100.0, value=85.0, step=1.0)
    st.caption("Premium amounts come only from Excel. Paid completion remains an editable dashboard assumption; 85% is an arbitrary starting value.")
    if not clients or not products:
        st.info("Select at least one client and product.")
        return
    try:
        result, scoped = reserve_scenario(data, premiums, clients, products, start, end, currency, completion)
    except ValueError as exc:
        st.error(str(exc))
        return
    missing = result["Premium missing"].any()
    premium = result["Earned Premium"].sum() if not missing else float("nan")
    paid = result["Paid claims"].sum()
    ultimate = result["Illustrative ultimate claims"].sum()
    reserve = result["Illustrative unpaid reserve"].sum()
    if missing:
        st.error("Premium rows are missing for the scope below. Overall premium and loss ratios are withheld. Add rows in Excel, or narrow the period/client/product selection. Missing values are not assumed to be zero.")
        st.dataframe(result.loc[result["Premium missing"], PREMIUM_KEYS], hide_index=True)
    x, y, z = st.columns(3)
    x.metric("Fictional earned premium · " + currency, "Incomplete" if missing else f"{premium:,.2f}")
    y.metric("Illustrative unpaid reserve · " + currency, f"{reserve:,.2f}")
    z.metric("Illustrative ultimate claims · " + currency, f"{ultimate:,.2f}")
    x, y = st.columns(2)
    available = not missing and premium > 0
    x.metric("Paid loss ratio · fictional premium", f"{paid/premium:.1%}" if available else "N/A")
    y.metric("Ultimate loss ratio · scenario", f"{ultimate/premium:.1%}" if available else "N/A")
    st.caption(f"Treatment months {start} to {end}; payments processed after month-end {end} are excluded. Premiums are summed once per client/product/month/currency, including rows with no paid claims. Zero total premium produces N/A ratios.")
    values = ["Earned Premium", "Paid claims", "Illustrative unpaid reserve", "Illustrative ultimate claims"]
    def aggregate(key):
        totals = result.groupby(key)[values].sum()
        incomplete = result.groupby(key)["Premium missing"].any()
        totals.loc[incomplete, "Earned Premium"] = float("nan")
        denominator = totals["Earned Premium"].replace(0, float("nan"))
        totals["Paid loss ratio (%)"] = totals["Paid claims"] / denominator * 100
        totals["Ultimate loss ratio (%)"] = totals["Illustrative ultimate claims"] / denominator * 100
        return totals
    st.subheader("Claims and fictional premium by month")
    st.line_chart(aggregate("Month")[["Earned Premium", "Paid claims", "Illustrative ultimate claims"]])
    st.subheader("Client comparison · fictional premiums")
    st.dataframe(aggregate("Client").round(2))
    st.subheader("Monthly calculation detail")
    st.dataframe(result.round(2), hide_index=True)
    with st.expander("Calculation and coverage assumptions", expanded=True):
        st.write("Ultimate claims = paid claims ÷ assumed paid-completion proportion. Unpaid reserve = ultimate − paid. Paid and ultimate loss ratios divide their respective claims totals by the earned premium supplied in Excel.")
        st.write("Every selected client/product combination is expected to have a premium row for every selected month. This conservative completeness check can flag inactive months; narrow your selection or supply explicit zero rows where correct. Client and product spelling must match the claims file.")
        st.write("The unpaid reserve includes reported outstanding claims and IBNR together; no separate IBNR is estimated. Under-review batches are not added again. Completion is a uniform assumption, not a fitted claims-development factor. Zero paid claims produces zero reserve under this simple model.")
        st.write("The supplied paid file is an August 2026 processing extract. It may omit payments outside that extract, so this is not full historical experience. Actual monthly exposure and annual benefit exhaustion are not calculated.")
    count = int(scoped["Checks"].ne("").sum())
    if count:
        st.warning(f"Scenario totals retain {count:,} records flagged by the validation checks.")


def main():
    st.set_page_config(page_title=TITLE, page_icon="📊", layout="wide")
    st.markdown("""<style>
    .stApp {background:#f5f7fb;}
    h1,h2,h3 {color:#122849;}
    [data-testid="stMetric"] {background:white;padding:20px;border-radius:12px;border:1px solid #e0e7f0;}
    </style>""", unsafe_allow_html=True)
    st.title(TITLE)
    st.caption("Version 0.5 · Claims overview · Reserves & loss ratio · Payment checks")
    st.info("Local prototype for fictional data. Online sign-in and user permissions are not configured yet.")
    upload = st.sidebar.file_uploader("Open paid claims workbook", type=["xlsx"])
    member_upload = st.sidebar.file_uploader("Open membership workbook (optional)", type=["xlsx"])
    review_upload = st.sidebar.file_uploader("Open claims under review (optional)", type=["xlsx"])
    premium_upload = st.sidebar.file_uploader("Open premium workbook (optional)", type=["xlsx"])
    if upload is None:
        st.subheader("Start with your claims workbook")
        st.write("Choose paid_claims.xlsx in the sidebar. The workbook must contain the 'paid claims' sheet.")
        st.write("You will be able to filter by client, product, claim type, currency and treatment month.")
        return
    try:
        # Session-local data only: no global cache or file writes.
        raw = pd.read_excel(io.BytesIO(upload.getvalue()), sheet_name="paid claims", engine="openpyxl")
        data, rules = prepare(raw)
    except Exception as exc:
        st.error(f"Could not read the workbook: {exc}")
        return
    members = None
    if member_upload is not None:
        try:
            members = pd.read_excel(io.BytesIO(member_upload.getvalue()), sheet_name="membership_data")
            members.columns = members.columns.astype(str).str.strip()
            member_ids = members["4037"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
            claim_ids = data["MEMBER_NUMBER"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
            data["Member not in register"] = ~claim_ids.isin(member_ids)
            rules.append("Member not in register")
            data["Checks"] = data[rules].apply(lambda row: "; ".join(row.index[row]), axis=1)
        except Exception as exc:
            st.error(f"Membership workbook could not be matched: {exc}")
            return
    st.sidebar.subheader("Filters")
    filtered = data.copy()
    selections = {}
    for label, col in [("Client", "CLIENTS"), ("Product", "SCHEME_NAME"), ("Claim type", "CLAIM_TYPE"), ("Treatment month", "Treatment month")]:
        selected = st.sidebar.selectbox(label, ["All"] + sorted(data[col].unique().tolist()))
        selections[col] = selected
        if selected != "All":
            filtered = filtered[filtered[col] == selected]
    currency = st.sidebar.selectbox("Currency", sorted(data["CLAIM_CURRECY"].unique().tolist()))
    filtered = filtered[filtered["CLAIM_CURRECY"] == currency]
    if filtered.empty:
        st.warning("No claims match these filters. Choose a different selection.")
        return
    flagged = filtered[filtered["Checks"] != ""]
    paid = filtered["AMOUNT_AGREED"].sum(min_count=1)
    cards = st.columns(4)
    cards[0].metric("Paid amount · " + currency, "Unavailable" if pd.isna(paid) else f"{paid:,.2f}")
    cards[1].metric("Claim records", f"{len(filtered):,}")
    cards[2].metric("Records needing review", f"{len(flagged):,}")
    lag = filtered.loc[filtered["Processing lag (days)"] >= 0, "Processing lag (days)"].median()
    cards[3].metric("Median processing lag", "Unavailable" if pd.isna(lag) else f"{lag:,.0f} days")
    st.caption("Totals retain source payment exceptions. Missing amounts are excluded from sums and flagged. Lag measures treatment to processing, not provider submission time.")
    overview, reserves, checks, membership_tab, pipeline, details = st.tabs(["Overview", "Reserves & loss ratio", "Payment checks", "Membership", "Under review", "About the data"])
    with reserves:
        render_reserves(data, [("Paid claims workbook", upload), ("Membership workbook", member_upload), ("Under-review workbook", review_upload), ("Separate premium workbook", premium_upload)], members)
    with overview:
        st.subheader("Paid claims by treatment month")
        monthly = filtered.groupby("Treatment month")["AMOUNT_AGREED"].sum(min_count=1).rename("Paid amount")
        st.bar_chart(monthly, color="#2864d7")
        left, right = st.columns(2)
        with left:
            st.subheader("Paid amount by claim type")
            types = filtered.groupby("CLAIM_TYPE")["AMOUNT_AGREED"].sum(min_count=1).sort_values(ascending=False).rename("Paid amount")
            st.bar_chart(types, color="#128b91")
        with right:
            st.subheader("Client summary")
            summary = filtered.groupby("CLIENTS").agg(Records=("CLAIM_NUMBER", "size"), Paid=("AMOUNT_AGREED", lambda s: s.sum(min_count=1)))
            st.dataframe(summary.sort_values("Paid", ascending=False))
        st.subheader("Rules-based observations")
        if len(flagged):
            st.write(f"{len(flagged):,} of {len(filtered):,} records in this selection triggered payment or data checks. Review them before relying on totals.")
        else:
            st.write("No records triggered the implemented checks in this selection. A register match alone does not establish eligibility at treatment date.")
        st.caption("Observations use explicit validation rules, not an AI model or fraud determination.")
    with checks:
        st.subheader("Review indicators")
        st.dataframe(pd.DataFrame({"Check": rules, "Flagged records": [int(filtered[r].sum()) for r in rules]}), hide_index=True)
        st.caption("One record can trigger several checks. Above-limit checks use the workbook's illustrative per-claim limits; they do not test annual benefit exhaustion.")
        if not flagged.empty:
            st.dataframe(flagged[["CLAIM_NUMBER", "CLIENTS", "SCHEME_NAME", "CLAIM_TYPE"] + AMOUNTS + ["Checks"]], hide_index=True)
        if members is None:
            st.info("Unmatched members cannot be checked until the membership workbook is supplied.")
        else:
            st.caption("Member matching checks whether the ID exists in the register. Policy-date eligibility and client/product consistency are not yet tested.")
    with membership_tab:
        if members is None:
            st.info("Open membership_data.xlsx in the sidebar to see membership and enable unmatched-member checks.")
        else:
            subset = members.copy()
            for claim_col, member_col in [("CLIENTS", "Company name"), ("SCHEME_NAME", "Product name")]:
                if selections[claim_col] != "All":
                    subset = subset[subset[member_col].astype(str).str.strip() == selections[claim_col]]
            st.subheader("Membership register")
            st.caption("Filtered by client and product only. This is the register snapshot, not treatment-month exposure. Claim-type, currency and treatment-month filters do not change this view.")
            st.metric("Distinct members in register", f"{subset['4037'].nunique():,}")
            st.dataframe(subset.groupby(["Company name", "Product name"])["4037"].nunique().rename("Members").reset_index(), hide_index=True)
            age = pd.to_numeric(subset["Age"], errors="coerce")
            bands = pd.cut(age, [-1, 17, 29, 39, 49, 59, 69, 200], labels=["0–17", "18–29", "30–39", "40–49", "50–59", "60–69", "70+"])
            st.subheader("Age distribution")
            st.bar_chart(bands.value_counts(sort=False).rename("Members"), color="#128b91")
            st.caption("Age uses the supplied snapshot values. Missing or out-of-range ages are excluded from the chart.")
    with pipeline:
        st.subheader("Claims under review · separate batch view")
        st.caption("This view has its own filters because batch data has no member, product or claim-type breakdown. Paid-claims filters do not apply here.")
        if review_upload is None:
            st.info("Open claims_under_review_corrected.xlsx in the sidebar.")
        else:
            try:
                frames = []
                for sheet, amount_col in [("Not Submitted", "BATCH AMOUNT"), ("Partially Submitted", "Unsubmitted Amount to Be Considered")]:
                    batch = pd.read_excel(io.BytesIO(review_upload.getvalue()), sheet_name=sheet).dropna(how="all")
                    batch.columns = batch.columns.astype(str).str.strip()
                    lookup = {c.lower(): c for c in batch.columns}
                    batch["Outstanding amount"] = clean_amounts(batch[lookup[amount_col.lower()]])
                    batch["Review status"] = sheet
                    for c in ["CLIENT NAME", "CURRENCY"]:
                        batch[c] = batch[c].fillna("Unknown").astype(str).str.strip()
                    frames.append(batch)
                batches = pd.concat(frames, ignore_index=True)
                review_client = st.selectbox("Review client", ["All"] + sorted(batches["CLIENT NAME"].unique().tolist()))
                review_currency = st.selectbox("Review currency", sorted(batches["CURRENCY"].unique().tolist()))
                batches = batches[batches["CURRENCY"] == review_currency]
                if review_client != "All":
                    batches = batches[batches["CLIENT NAME"] == review_client]
                outstanding = batches["Outstanding amount"].sum(min_count=1)
                st.metric("Outstanding pipeline · " + review_currency, "Unavailable" if pd.isna(outstanding) else f"{outstanding:,.2f}")
                st.dataframe(batches.groupby(["CLIENT NAME", "Review status"])["Outstanding amount"].sum(min_count=1).reset_index(), hide_index=True)
                if batches["Outstanding amount"].isna().any():
                    st.warning("Some outstanding amounts are missing or invalid and excluded from totals.")
                st.caption("Not Submitted uses BATCH AMOUNT. Partially Submitted uses only Unsubmitted Amount to Be Considered. No validity adjustment is applied; these amounts are not paid claims, IBNR or confirmed liabilities. Overlap with paid claims has not been reconciled.")
            except Exception as exc:
                st.error(f"Could not read claims under review: {exc}")
    with details:
        st.write("This first version is tailored to the supplied synthetic paid_claims.xlsx workbook. Its Notes sheet defines AMOUNT_AGREED as paid amount and all amounts as illustrative USD.")
        st.write("The supplied file includes deliberate payment errors. Checks are computed from amounts, not copied from the Test cases sheet.")
        st.write("Actual premiums and validated monthly exposure are not included. The Reserves & loss ratio tab reads fictional premiums from an Excel Premiums sheet or the membership annual premium column and uses editable paid-completion assumptions. Claims per covered life, annual benefit utilisation and a separately estimated IBNR are not calculated.")
        st.write("The earlier analysis notebook has not been ported in full. Its assumptions must be reconciled with this workbook before adding pricing and reserving calculations.")
        st.write("Names, member numbers and diagnosis text are not displayed. This is not an anonymisation guarantee; use fictional data until hosted access and disclosure controls are implemented.")


if __name__ == "__main__":
    main()
