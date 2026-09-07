"""Local Streamlit prototype. Run with --server.address 127.0.0.1.

Install: python -m pip install streamlit pandas openpyxl
This version has no hosted authentication. Use fictional data only.
"""
import io
import pandas as pd
import streamlit as st
import altair as alt

TITLE = "Claims Risk Management Intelligent Dashboard"
LOGO_DATA_URL = "data:image/png;base64,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"
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


def excel_paid_ratio(data, members, month, currency, clients, products):
    m = members.copy()
    m.columns = m.columns.astype(str).str.strip()
    required = {"4037", "Company name", "Product name", "annual premium", "Reporting Month"}
    if not required.issubset(m.columns):
        raise ValueError("Membership needs member ID, client, product, annual premium and Reporting Month.")
    ids = m["4037"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    if m["4037"].isna().any() or ids.eq("").any() or ids.duplicated().any():
        raise ValueError("Member IDs must be present and unique.")
    for col in ["Company name", "Product name"]:
        if m[col].isna().any():
            raise ValueError("Missing membership client or product.")
        m[col] = m[col].astype(str).str.strip()
    reporting = pd.to_datetime(m["Reporting Month"], errors="coerce").dt.strftime("%Y-%m")
    if reporting.isna().any() or not reporting.eq(month).all():
        raise ValueError("The membership snapshot must have Reporting Month equal to the selected processing month. Upload the matching snapshot.")
    m["annual premium"] = clean_amounts(m["annual premium"])
    if not m["annual premium"].map(lambda x: pd.notna(x) and 0 <= x < float("inf")).all():
        raise ValueError("Annual premiums must be valid nonnegative numbers.")
    m = m[m["Company name"].isin(clients) & m["Product name"].isin(products)]
    c = data[data["CLIENTS"].isin(clients) & data["SCHEME_NAME"].isin(products) & data["CLAIM_CURRECY"].str.upper().eq(currency) & data["CLAIM_PROCESSED_DATE"].dt.strftime("%Y-%m").eq(month)].copy()
    if c["AMOUNT_AGREED"].isna().any():
        raise ValueError("Invalid paid amounts in the selected processing month.")
    pre = m.groupby(["Company name", "Product name"])["annual premium"].sum().div(12).rename("Fictional monthly premium").reset_index().rename(columns={"Company name":"Client", "Product name":"Product"})
    paid = c.groupby(["CLIENTS", "SCHEME_NAME"])["AMOUNT_AGREED"].sum().rename("Paid claims").reset_index().rename(columns={"CLIENTS":"Client", "SCHEME_NAME":"Product"})
    result = pre.merge(paid, on=["Client", "Product"], how="outer", validate="one_to_one")
    result["Paid claims"] = result["Paid claims"].fillna(0)
    result["Paid claims ratio (%)"] = result["Paid claims"] / result["Fictional monthly premium"].replace(0,float("nan")) * 100
    return result, c


def render_excel_ratio(data, members):
    st.subheader("Paid claims ratio · Excel basis")
    st.info("Fictional premiums calibrated in Excel to a 75% August paid-claims ratio. The ratio below is calculated from the uploaded values, not forced to 75%.")
    st.caption("Processing-month basis: every listed member contributes annual premium ÷ 12 for the reporting month. Joining and expiry dates are not prorated. This matches the updated workbook's assumptions; it is not an incurred or ultimate loss ratio.")
    if members is None or "annual premium" not in members.columns:
        st.info("Select the updated membership workbook containing annual premium.")
        return

    months = sorted(data["CLAIM_PROCESSED_DATE"].dropna().dt.strftime("%Y-%m").unique().tolist())
    if not months:
        st.error("Valid processing dates are required.")
        return
    x,y=st.columns(2)
    month=x.selectbox("Processing month",months,index=len(months)-1,key="excel_month")
    currencies=sorted(data["CLAIM_CURRECY"].str.upper().unique().tolist())
    currency=y.selectbox("Annual premium currency · Excel assumption",currencies,index=currencies.index("USD") if "USD" in currencies else 0,key="excel_currency")
    st.caption("The updated workbook specifies USD in Premium assumptions. This selector assigns the membership rates to one currency; no currency conversion is performed.")
    clients=sorted(set(data["CLIENTS"]) | set(members["Company name"].dropna().astype(str).str.strip()))
    selected=st.multiselect("Clients · Excel basis",clients,default=clients,key="excel_clients")
    products=sorted(set(data.loc[data["CLIENTS"].isin(selected),"SCHEME_NAME"]) | set(members.loc[members["Company name"].isin(selected),"Product name"].dropna().astype(str).str.strip()))
    selected_p=st.multiselect("Products · Excel basis",products,default=products,key="excel_products")
    st.caption("These controls apply only to this view; sidebar treatment-month and claim-type filters do not apply.")
    if not selected or not selected_p:
        st.info("Select at least one client and product.")
        return
    try:
        result,scoped=excel_paid_ratio(data,members,month,currency,selected,selected_p)
    except ValueError as exc:
        st.error(str(exc));return
    missing=result["Fictional monthly premium"].isna().any()
    premium=result["Fictional monthly premium"].sum()
    paid=result["Paid claims"].sum()
    if missing:
        st.error("Claims have a client/product with no matching premium. Overall premium and ratio are withheld.")
    x,y,z=st.columns(3)
    x.metric("Paid claims · "+currency,f"{paid:,.2f}")
    y.metric("Fictional monthly premium · "+currency,"Incomplete" if missing else f"{premium:,.2f}")
    z.metric("Paid claims ratio · Excel basis",f"{paid/premium:.2%}" if not missing and premium>0 else "N/A")
    summary=result.groupby("Client")[["Paid claims","Fictional monthly premium"]].sum()
    incomplete=result.groupby("Client")["Fictional monthly premium"].apply(lambda v:v.isna().any())
    summary.loc[incomplete,"Fictional monthly premium"]=float("nan")
    summary["Paid claims ratio (%)"]=summary["Paid claims"]/summary["Fictional monthly premium"].replace(0,float("nan"))*100
    st.dataframe(summary.round(2))
    st.bar_chart(summary[["Paid claims","Fictional monthly premium"]])
    with st.expander("Product detail"):
        st.dataframe(result.round(2),hide_index=True)
    st.caption("All supplied paid amounts are retained, including deliberate payment errors and unmatched members assigned by client. No under-review amounts or reserve loading are added. Changing claims without updating the fictional premiums can change the ratio.")
    st.write("For the treatment-period reserve scenario, open the separate Reserves & loss ratio tab. Its date proration and incurred-period scope intentionally produce different results.")


COLORS = ["#3268f4", "#18b4ad", "#ffae42", "#9a70e8", "#f16ca6", "#6ba7f8", "#596980", "#ef785d"]


def chart_frame(chart):
    return chart.configure_view(strokeWidth=0).configure_axis(gridColor="#edf0f6", domain=False, labelColor="#758299", titleColor="#758299", labelFontSize=11).configure_legend(labelColor="#586780", title=None)


def donut(series, label, value="Records"):
    frame=series.rename(value).rename_axis(label).reset_index()
    if frame.empty or frame[value].lt(0).any() or frame[value].sum()<=0:
        st.info("No positive distribution available for this selection."); return
    chart=alt.Chart(frame).mark_arc(innerRadius=65,outerRadius=105,cornerRadius=2).encode(theta=alt.Theta(value+":Q"),color=alt.Color(label+":N",scale=alt.Scale(range=COLORS),legend=alt.Legend(orient="right")),tooltip=[label,alt.Tooltip(value+":Q",format=",.0f")]).properties(height=290)
    st.altair_chart(chart_frame(chart),use_container_width=True)


def bars(series, label, value="Paid amount", horizontal=False):
    frame=series.rename(value).rename_axis(label).reset_index()
    base=alt.Chart(frame).mark_bar(color=COLORS[0],cornerRadiusTopLeft=3,cornerRadiusTopRight=3)
    chart=base.encode(x=alt.X(value+":Q",title=None),y=alt.Y(label+":N",sort="-x",title=None),tooltip=[label,alt.Tooltip(value+":Q",format=",.2f")]) if horizontal else base.encode(x=alt.X(label+":N",title=None,axis=alt.Axis(labelAngle=0)),y=alt.Y(value+":Q",title=None),tooltip=[label,alt.Tooltip(value+":Q",format=",.2f")])
    st.altair_chart(chart_frame(chart.properties(height=290)),use_container_width=True)


def overview_panels(filtered, flagged, currency, claims_page=False):
    a,b=st.columns([1,1.3])
    with a,st.container(border=True):
        st.subheader("Benefit distribution")
        st.caption("Claim records by benefit type")
        donut(filtered.groupby("CLAIM_TYPE").size(),"Benefit")
    with b,st.container(border=True):
        st.subheader("Paid claims trend")
        st.caption(currency+" · grouped by treatment month")
        bars(filtered.groupby("Treatment month")["AMOUNT_AGREED"].sum(min_count=1),"Month")
    a,b=st.columns(2)
    with a,st.container(border=True):
        st.subheader("Client distribution")
        st.caption("Total paid amount · "+currency)
        bars(filtered.groupby("CLIENTS")["AMOUNT_AGREED"].sum(min_count=1).sort_values(ascending=False),"Client",horizontal=True)
    with b,st.container(border=True):
        st.subheader("Payment review")
        st.caption("Records that triggered at least one validation check")
        donut(pd.Series({"No flags":len(filtered)-len(flagged),"Needs review":len(flagged)}),"Review")
    if claims_page:
        with st.container(border=True):
            st.subheader("Provider analysis")
            if "NAME_OF_PAYEE" in filtered:
                providers=filtered.groupby("NAME_OF_PAYEE")["AMOUNT_AGREED"].sum(min_count=1).nlargest(10)
                bars(providers,"Provider",horizontal=True)
            st.caption("Highest paid totals reflect volume and case mix; they are not risk-adjusted performance rankings.")
    with st.container(border=True):
        st.subheader("Client summary")
        summary=filtered.groupby("CLIENTS").agg(Records=("CLAIM_NUMBER","size"),Paid=("AMOUNT_AGREED",lambda v:v.sum(min_count=1))).reset_index().rename(columns={"CLIENTS":"Client"})
        st.dataframe(summary.sort_values("Paid",ascending=False),hide_index=True,use_container_width=True)
    with st.expander("View claim detail"):
        cols=["CLAIM_NUMBER","CLIENTS","SCHEME_NAME","CLAIM_TYPE","Treatment month","AMOUNT_AGREED","Checks"]
        query=st.text_input("Search claim reference",key="claim_search")
        detail=filtered[cols]
        if query:detail=detail[detail.CLAIM_NUMBER.astype(str).str.contains(query,case=False,regex=False)]
        st.dataframe(detail,hide_index=True,use_container_width=True)


def main():
    st.set_page_config(page_title=TITLE, page_icon="📊", layout="wide")
    st.markdown("""<style>
    .stApp {background:#f5f7fb;color:#1c2941;}
    .block-container {padding-top:2rem;max-width:1500px;padding-bottom:3rem;}
    h1 {font-size:1.8rem!important;font-weight:650!important;letter-spacing:-.035em;color:#192652;}
    h2,h3 {color:#243365;font-size:1rem!important;font-weight:600!important;}
    [data-testid="stSidebar"] {background:#192453;min-width:240px;max-width:260px;}
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"], [data-testid="stSidebar"] label {color:#e5ebff;}
    [data-testid="stSidebar"] [data-testid="stCaptionContainer"] {color:#aebbe0;}
    [data-testid="stSidebar"] [data-testid="stRadio"] label {padding:9px 12px;border-radius:7px;width:100%;}
    [data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) {background:#304578;}
    [data-testid="stSidebar"] [data-testid="stFileUploader"] label {color:#e5ebff;}
    [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {background:#fff;color:#243365;}
    [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] * {color:#243365;}
    [data-testid="stMetric"] {background:white;padding:18px 20px;border:1px solid #e9edf5;border-radius:7px;border-bottom:3px solid #dce5ff;}
    [data-testid="stMetricLabel"] {color:#6f7b91;font-size:12px;}
    [data-testid="stMetricValue"] {color:#1c2941;font-size:1.65rem;font-weight:650;}
    [data-testid="stVerticalBlockBorderWrapper"]>div {background:white;border-color:#e9edf5!important;border-radius:8px;}
    [data-testid="stCaptionContainer"] {color:#7a879d;}
    [data-testid="stDataFrame"] {border:1px solid #edf0f5;border-radius:6px;}
    .brand {color:white;font-size:23px;font-weight:650;letter-spacing:2px;margin:10px 0 2px;}
    .brand-sub {color:#acbadd;font-size:10px;letter-spacing:2px;margin-bottom:28px;}
    .topline {font-size:10px;color:#75829b;letter-spacing:1.5px;border-bottom:1px solid #e4e9f2;padding-bottom:14px;margin-bottom:20px;}
    </style>""",unsafe_allow_html=True)
    st.sidebar.markdown(f'<div style="background:white;border-radius:8px;padding:12px;margin:8px 0 24px"><img src="{LOGO_DATA_URL}" alt="Health Systems Link" style="width:100%;height:auto;display:block"></div>',unsafe_allow_html=True)
    page=st.sidebar.radio("Workspace",["Overview","Claims analytics","Membership","Premiums & paid ratio","Reserves & loss ratio","Claims under review","Data quality","About the data"],key="navigation",label_visibility="collapsed")
    with st.sidebar.expander("Data setup",expanded=False):
        upload=st.file_uploader("Open paid claims workbook",type=["xlsx"])
        member_upload=st.file_uploader("Open membership workbook (optional)",type=["xlsx"])
        review_upload=st.file_uploader("Open claims under review (optional)",type=["xlsx"])
        premium_upload=st.file_uploader("Open premium workbook (optional)",type=["xlsx"])
    st.sidebar.caption("Version 0.8 · Fictional data prototype")
    st.title(TITLE)
    st.subheader(page)
    st.caption("Explore claims, membership and financial scenarios. All figures use the files you select.")
    if upload is None:
        st.subheader("Start with your claims workbook")
        st.write("Expand Data setup in the sidebar and choose paid_claims.xlsx. The workbook must contain the 'paid claims' sheet.")
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
    filtered=data.copy()
    selections={"CLIENTS":"All","SCHEME_NAME":"All","CLAIM_TYPE":"All","Treatment month":"All"}
    currency=sorted(data["CLAIM_CURRECY"].unique().tolist())[0]
    if page in ["Overview","Claims analytics","Membership","Data quality"]:
        with st.container(border=True):
            filters=st.columns(5)
            for box,(label,col) in zip(filters,[("Client","CLIENTS"),("Product","SCHEME_NAME"),("Claim type","CLAIM_TYPE"),("Treatment month","Treatment month")]):
                selected=box.selectbox(label,["All"]+sorted(data[col].unique().tolist()),key="filter_"+col,disabled=page=="Membership" and col in ["CLAIM_TYPE","Treatment month"])
                if page=="Membership" and col in ["CLAIM_TYPE","Treatment month"]:selected="All"
                selections[col]=selected
                if selected!="All":filtered=filtered[filtered[col]==selected]
            currency=filters[4].selectbox("Currency",sorted(data["CLAIM_CURRECY"].unique().tolist()),disabled=page=="Membership")
        filtered=filtered[filtered["CLAIM_CURRECY"]==currency]
        if filtered.empty and page!="Membership":
            st.info("No claims match these filters. Choose another selection.");return
    flagged = filtered[filtered["Checks"] != ""]
    paid = filtered["AMOUNT_AGREED"].sum(min_count=1)
    if page in ["Overview","Claims analytics","Data quality"]:
        cards = st.columns(4)
        cards[0].metric("Paid amount · " + currency, "Unavailable" if pd.isna(paid) else f"{paid:,.2f}")
        cards[1].metric("Claim records", f"{len(filtered):,}")
        cards[2].metric("Records needing review", f"{len(flagged):,}")
        lag = filtered.loc[filtered["Processing lag (days)"] >= 0, "Processing lag (days)"].median()
        cards[3].metric("Median processing lag", "Unavailable" if pd.isna(lag) else f"{lag:,.0f} days")
        st.caption("Totals retain source payment exceptions. Missing amounts are excluded from sums and flagged. Lag measures treatment to processing, not provider submission time.")
    if page == "Premiums & paid ratio":
        render_excel_ratio(data, members)
    if page == "Reserves & loss ratio":
        render_reserves(data, [("Paid claims workbook", upload), ("Membership workbook", member_upload), ("Under-review workbook", review_upload), ("Separate premium workbook", premium_upload)], members)
    if page in ["Overview","Claims analytics"]:
        overview_panels(filtered,flagged,currency,page=="Claims analytics")
    if page == "Data quality":
        st.subheader("Review indicators")
        st.dataframe(pd.DataFrame({"Check": rules, "Flagged records": [int(filtered[r].sum()) for r in rules]}), hide_index=True)
        st.caption("One record can trigger several checks. Above-limit checks use the workbook's illustrative per-claim limits; they do not test annual benefit exhaustion.")
        if not flagged.empty:
            st.dataframe(flagged[["CLAIM_NUMBER", "CLIENTS", "SCHEME_NAME", "CLAIM_TYPE"] + AMOUNTS + ["Checks"]], hide_index=True)
        if members is None:
            st.info("Unmatched members cannot be checked until the membership workbook is supplied.")
        else:
            st.caption("Member matching checks whether the ID exists in the register. Policy-date eligibility and client/product consistency are not yet tested.")
    if page == "Membership":
        if members is None:
            st.info("Open membership_data.xlsx in the sidebar to see membership and enable unmatched-member checks.")
        else:
            subset = members.copy()
            for claim_col, member_col in [("CLIENTS", "Company name"), ("SCHEME_NAME", "Product name")]:
                if selections[claim_col] != "All":
                    subset = subset[subset[member_col].astype(str).str.strip() == selections[claim_col]]
            st.caption("Membership snapshot · client and product filters only. Status describes relationship, not active/inactive membership.")
            age=pd.to_numeric(subset["Age"],errors="coerce")
            cards=st.columns(4)
            cards[0].metric("Total members",f"{subset['4037'].nunique():,}")
            cards[1].metric("Clients",f"{subset['Company name'].nunique():,}")
            cards[2].metric("Products",f"{subset['Product name'].nunique():,}")
            cards[3].metric("Average age",f"{age.mean():.1f}" if age.notna().any() else "N/A")
            a,b=st.columns(2)
            with a,st.container(border=True):
                st.subheader("Product enrollment")
                donut(subset.groupby("Product name")["4037"].nunique(),"Product","Members")
            with b,st.container(border=True):
                st.subheader("Relationship distribution")
                donut(subset.groupby("Status")["4037"].nunique(),"Relationship","Members")
            a,b=st.columns(2)
            with a,st.container(border=True):
                st.subheader("Age distribution")
                bands=pd.cut(age,[-1,17,29,39,49,59,69,200],labels=["0–17","18–29","30–39","40–49","50–59","60–69","70+"])
                bars(bands.value_counts(sort=False),"Age band","Members")
            with b,st.container(border=True):
                st.subheader("Gender distribution")
                donut(subset.groupby("Gender",dropna=False)["4037"].nunique(),"Gender","Members")
            with st.expander("View membership summary"):
                st.dataframe(subset.groupby(["Company name","Product name"])["4037"].nunique().rename("Members").reset_index(),hide_index=True,use_container_width=True)
    if page == "Claims under review":
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
    if page == "About the data":
        st.write("This dashboard is tailored to the supplied synthetic paid_claims.xlsx workbook. Its Notes sheet defines AMOUNT_AGREED as paid amount and all amounts as illustrative USD.")
        st.write("The supplied file includes deliberate payment errors. Checks are computed from amounts, not copied from the Test cases sheet.")
        st.write("Actual premiums and validated monthly exposure are not included. The Reserves & loss ratio tab reads fictional premiums from an Excel Premiums sheet or the membership annual premium column and uses editable paid-completion assumptions. Claims per covered life, annual benefit utilisation and a separately estimated IBNR are not calculated.")
        st.write("The earlier analysis notebook has not been ported in full. Its assumptions must be reconciled with this workbook before adding pricing and reserving calculations.")
        st.write("Names, member numbers and diagnosis text are not displayed. This is not an anonymisation guarantee; use fictional data until hosted access and disclosure controls are implemented.")


if __name__ == "__main__":
    main()
