"""Local Streamlit prototype. Run with --server.address 127.0.0.1.

Install: python -m pip install streamlit pandas openpyxl
This version has no hosted authentication. Use fictional data only.
"""
import io
import pandas as pd
import streamlit as st
import altair as alt

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
    st.sidebar.markdown('<div class="brand">HSL <span style="color:#769aff">◈</span></div><div class="brand-sub">HEALTHCARE ANALYTICS</div>',unsafe_allow_html=True)
    page=st.sidebar.radio("Workspace",["Overview","Claims analytics","Membership","Premiums & paid ratio","Reserves & loss ratio","Claims under review","Data quality","About the data"],key="navigation",label_visibility="collapsed")
    with st.sidebar.expander("Data setup",expanded=False):
        upload=st.file_uploader("Open paid claims workbook",type=["xlsx"])
        member_upload=st.file_uploader("Open membership workbook (optional)",type=["xlsx"])
        review_upload=st.file_uploader("Open claims under review (optional)",type=["xlsx"])
        premium_upload=st.file_uploader("Open premium workbook (optional)",type=["xlsx"])
    st.sidebar.caption("Version 0.7 · Fictional data prototype")
    st.markdown('<div class="topline">CLAIMS RISK MANAGEMENT INTELLIGENT DASHBOARD &nbsp; / &nbsp; ANALYTICS WORKSPACE</div>',unsafe_allow_html=True)
    st.title(page)
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
