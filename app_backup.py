import streamlit as st
import pandas as pd
import altair as alt
import html

# ============================================================
# PAGE CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="Customer Retention Intelligence",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# CUSTOM CSS
# ============================================================
st.markdown("""
<style>

/* ---------- MAIN PAGE ---------- */

.stApp {
    background-color: #F5F7FB;
}

.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
}

/* ---------- SIDEBAR ---------- */

section[data-testid="stSidebar"] {
    background-color: #182235;
}

section[data-testid="stSidebar"] * {
    color: #FFFFFF !important;
}

section[data-testid="stSidebar"] .stRadio label {
    color: #FFFFFF !important;
}

/* ---------- HEADINGS ---------- */

.main-title {
    font-size: 38px;
    font-weight: 750;
    color: #17365D;
    margin-bottom: 4px;
}

.subtitle {
    font-size: 17px;
    color: #5B6573;
    margin-bottom: 22px;
}

.section-title {
    font-size: 27px;
    font-weight: 700;
    color: #17365D;
    margin-top: 12px;
    margin-bottom: 18px;
}

.small-title {
    font-size: 21px;
    font-weight: 700;
    color: #17365D;
    margin-top: 18px;
    margin-bottom: 12px;
}

/* ---------- KPI CARDS ---------- */

.kpi-card {
    background: #FFFFFF;
    padding: 22px;
    border-radius: 15px;
    border: 1px solid #DCE3ED;
    box-shadow: 0 4px 14px rgba(24, 34, 53, 0.08);
    text-align: center;
}

.kpi-title {
    font-size: 14px;
    color: #687385;
    font-weight: 650;
    margin-bottom: 8px;
}

.kpi-value {
    font-size: 30px;
    font-weight: 800;
    color: #17365D;
}

.blue {
    border-top: 5px solid #2F75B5;
}

.green {
    border-top: 5px solid #28A745;
}

.orange {
    border-top: 5px solid #F39C12;
}

.red {
    border-top: 5px solid #D9534F;
}

.purple {
    border-top: 5px solid #7A5AF8;
}

/* ---------- INSIGHT BOXES ---------- */

.insight-box {
    background: #FFFFFF;
    color: #263445;
    padding: 18px 20px;
    border-radius: 13px;
    border-left: 5px solid #2F75B5;
    border-top: 1px solid #DCE3ED;
    border-right: 1px solid #DCE3ED;
    border-bottom: 1px solid #DCE3ED;
    box-shadow: 0 3px 10px rgba(24, 34, 53, 0.06);
    margin-top: 14px;
}

.insight-title {
    color: #17365D;
    font-weight: 750;
    margin-bottom: 7px;
}

/* ---------- FINDING BOXES ---------- */

.finding-blue {
    background: #E8F3FB;
    color: #17365D;
    border-left: 5px solid #2F75B5;
}

.finding-green {
    background: #EAF7EE;
    color: #185C2D;
    border-left: 5px solid #28A745;
}

.finding-orange {
    background: #FFF5E5;
    color: #744900;
    border-left: 5px solid #F39C12;
}

.finding-red {
    background: #FDECEC;
    color: #7C2525;
    border-left: 5px solid #D9534F;
}

/* ---------- TABLE ---------- */

.custom-table-wrap {
    background: #FFFFFF;
    border: 1px solid #DCE3ED;
    border-radius: 12px;
    overflow-x: auto;
    box-shadow: 0 3px 10px rgba(24, 34, 53, 0.05);
}

.custom-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 14px;
    color: #263445;
    background: #FFFFFF;
}

.custom-table th {
    background: #17365D;
    color: #FFFFFF;
    padding: 11px 12px;
    text-align: center;
    font-weight: 700;
    border-bottom: 2px solid #DCE3ED;
}

.custom-table td {
    padding: 10px 12px;
    text-align: center;
    color: #263445;
    border-bottom: 1px solid #E8ECF2;
    background: #FFFFFF;
}

.custom-table tr:nth-child(even) td {
    background: #F7F9FC;
}

/* ---------- NOTES ---------- */

.note-box {
    background: #FFF7E6;
    color: #6B4A00;
    border: 1px solid #F1D48A;
    border-radius: 12px;
    padding: 14px 16px;
    margin-top: 14px;
}

.info-box {
    background: #E9F4FF;
    color: #174A70;
    border: 1px solid #B9D9F3;
    border-radius: 12px;
    padding: 14px 16px;
    margin-top: 14px;
}

.success-box {
    background: #EAF7EE;
    color: #185C2D;
    border: 1px solid #B9E2C5;
    border-radius: 12px;
    padding: 14px 16px;
    margin-top: 14px;
}

.danger-box {
    background: #FDECEC;
    color: #7C2525;
    border: 1px solid #F1B8B8;
    border-radius: 12px;
    padding: 14px 16px;
    margin-top: 14px;
}

/* ---------- SELECT BOX ---------- */

[data-testid="stSelectbox"] label {
    color: #17365D !important;
    font-weight: 650 !important;
}

[data-testid="stSelectbox"] div[data-baseweb="select"] {
    background-color: #FFFFFF !important;
}

[data-testid="stSelectbox"] div[data-baseweb="select"] * {
    color: #17365D !important;
}

/* ---------- FOOTER ---------- */

.footer {
    text-align: center;
    color: #6E7886;
    font-size: 13px;
    margin-top: 38px;
    padding-top: 15px;
    border-top: 1px solid #DCE3ED;
}

</style>
""", unsafe_allow_html=True)

# ============================================================
# LOAD DATA
# ============================================================

@st.cache_data
def load_data():
    excel_file = "Bank_Customer_Churn_Analysis_Final.xlsx"
    prediction_file = "customer_churn_risk_predictions.csv"

    data = pd.read_excel(
        excel_file,
        sheet_name="European_Bank (1)"
    )

    predictions = pd.read_csv(prediction_file)

    return data, predictions


try:
    data, predictions = load_data()
except Exception as e:
    st.error("Could not load the project files.")
    st.code(str(e))
    st.stop()

# ============================================================
# DATA PREPARATION
# ============================================================

data["Status"] = data["IsActiveMember"].map({
    1: "Active",
    0: "Inactive"
})

data["Credit Card"] = data["HasCrCard"].map({
    1: "Has Card",
    0: "No Card"
})

data["Churn Status"] = data["Exited"].map({
    1: "Churned",
    0: "Retained"
})

data["Relationship Strength"] = (
    data["IsActiveMember"] + data["NumOfProducts"]
)

data["Priority Risk"] = (
    (data["IsActiveMember"] == 0) &
    (data["NumOfProducts"] == 1) &
    (data["Age"] >= 40)
).map({
    True: "Priority Risk",
    False: "Standard"
})

data["Age Group"] = pd.cut(
    data["Age"],
    bins=[0, 29, 39, 49, 59, 200],
    labels=["18–29", "30–39", "40–49", "50–59", "60+"]
)

# ============================================================
# KPI CALCULATIONS
# ============================================================

total_customers = len(data)

total_churned = int(
    data["Exited"].sum()
)

overall_churn = (
    total_churned /
    total_customers *
    100
)

test_set_customers = len(predictions)

high_risk = int(
    (predictions["Risk_Flag"] == "High Risk").sum()
)

high_risk_rate = (
    high_risk /
    test_set_customers *
    100
)

# ============================================================
# ALTair CHART HELPER
# ============================================================

def create_bar_chart(
    df,
    x_field,
    y_field,
    x_title,
    y_title,
    color="#2F75B5",
    sort=None
):

    tooltip_fields = []

    for column in df.columns:
        tooltip_fields.append(
            alt.Tooltip(
                column,
                type="nominal" if df[column].dtype == "object"
                else "quantitative"
            )
        )

    chart = (
        alt.Chart(df)
        .mark_bar(
            color=color,
            cornerRadiusTopLeft=7,
            cornerRadiusTopRight=7
        )
        .encode(
            x=alt.X(
                x_field,
                title=x_title,
                sort=sort,
                axis=alt.Axis(
                    labelColor="#17365D",
                    titleColor="#17365D",
                    labelAngle=-30
                )
            ),
            y=alt.Y(
                y_field,
                title=y_title,
                axis=alt.Axis(
                    labelColor="#17365D",
                    titleColor="#17365D",
                    gridColor="#E3E8EF"
                )
            ),
            tooltip=tooltip_fields
        )
        .properties(
            height=350,
            background="white"
        )
        .configure_view(
            stroke=None
        )
        .configure_axis(
            domainColor="#CBD3DF"
        )
    )

    return chart


# ============================================================
# HTML TABLE HELPER
# ============================================================

def render_table(df):

    safe_df = df.copy()

    return st.markdown(
        '<div class="custom-table-wrap">'
        + safe_df.to_html(
            index=False,
            classes="custom-table",
            border=0
        )
        + '</div>',
        unsafe_allow_html=True
    )


# ============================================================
# HEADER
# ============================================================

st.markdown(
    '<div class="main-title">'
    '📊 Customer Retention Intelligence Dashboard'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">'
    'Customer Engagement, Product Utilization & Churn Risk Analytics'
    '</div>',
    unsafe_allow_html=True
)

# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title("📌 Dashboard Menu")

page = st.sidebar.radio(
    "Select Analysis",
    [
        "Overview",
        "Customer Engagement",
        "Product Utilization",
        "Financial Commitment",
        "Salary & Balance",
        "Retention Strength",
        "Geography",
        "Priority Risk",
        "ML Risk Analysis",
        "Customer Lookup"
    ]
)

st.sidebar.divider()

st.sidebar.write("**Project:** Banking Customer Retention")
st.sidebar.write("**Dataset:** 10,000 customers")
st.sidebar.write("**Model:** Random Forest")

st.sidebar.divider()

st.sidebar.download_button(
    label="⬇ Download Predictions",
    data=predictions.to_csv(index=False),
    file_name="customer_churn_risk_predictions.csv",
    mime="text/csv"
)

# ============================================================
# OVERVIEW
# ============================================================

if page == "Overview":

    st.markdown(
        '<div class="section-title">'
        'Executive Overview'
        '</div>',
        unsafe_allow_html=True
    )

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown(
            f"""
            <div class="kpi-card blue">
                <div class="kpi-title">Total Customers</div>
                <div class="kpi-value">{total_customers:,}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            f"""
            <div class="kpi-card red">
                <div class="kpi-title">Churned Customers</div>
                <div class="kpi-value">{total_churned:,}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with col3:
        st.markdown(
            f"""
            <div class="kpi-card orange">
                <div class="kpi-title">Overall Churn Rate</div>
                <div class="kpi-value">{overall_churn:.2f}%</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with col4:
        st.markdown(
            f"""
            <div class="kpi-card green">
                <div class="kpi-title">Test-Set High-Risk Customers</div>
                <div class="kpi-value">{high_risk:,}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown(
        f"""
        <div class="info-box">
        <b>ML Risk Scope:</b>
        The {high_risk:,} high-risk customers shown here come from the
        {test_set_customers:,}-customer prediction/test dataset, not the
        full 10,000-customer dataset.
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-title">'
        'Key Project Findings'
        '</div>',
        unsafe_allow_html=True
    )

    c1, c2 = st.columns(2)

    with c1:

        st.markdown(
            """
            <div class="insight-box finding-blue">
            <div class="insight-title">🔵 Customer Engagement</div>
            Inactive customers show a higher observed churn rate
            than active customers.
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown(
            """
            <div class="insight-box finding-orange">
            <div class="insight-title">🟠 Product Utilization</div>
            Customers with one product show a higher observed churn
            rate than customers with two products.
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown(
            """
            <div class="insight-box finding-green">
            <div class="insight-title">🟢 Credit Card</div>
            Credit-card ownership shows limited differentiation
            in observed churn.
            </div>
            """,
            unsafe_allow_html=True
        )

    with c2:

        st.markdown(
            """
            <div class="insight-box finding-red">
            <div class="insight-title">🔴 Priority Risk</div>
            Inactive customers with one product and age 40+
            form the project-defined priority-risk group.
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown(
            """
            <div class="insight-box finding-blue">
            <div class="insight-title">🔵 Geography</div>
            Observed churn differs across France, Germany and Spain.
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown(
            """
            <div class="insight-box finding-orange">
            <div class="insight-title">🟠 ML Risk</div>
            The Random Forest model provides customer-level churn
            risk probabilities for prioritization.
            </div>
            """,
            unsafe_allow_html=True
        )

# ============================================================
# CUSTOMER ENGAGEMENT
# ============================================================

elif page == "Customer Engagement":

    st.markdown(
        '<div class="section-title">'
        'Customer Engagement & Churn'
        '</div>',
        unsafe_allow_html=True
    )

    engagement = (
        data.groupby("Status")["Exited"]
        .agg(["count", "sum"])
        .reset_index()
    )

    engagement["Churn Rate"] = (
        engagement["sum"] /
        engagement["count"] *
        100
    )

    engagement["Status"] = pd.Categorical(
        engagement["Status"],
        categories=["Active", "Inactive"],
        ordered=True
    )

    engagement = engagement.sort_values("Status")

    chart_df = engagement.copy()

    chart_df["Churn Rate"] = chart_df["Churn Rate"].round(2)

    col1, col2 = st.columns([1.5, 1])

    with col1:

        chart = create_bar_chart(
            chart_df,
            "Status",
            "Churn Rate",
            "Engagement Status",
            "Observed Churn Rate (%)",
            "#2F75B5",
            ["Active", "Inactive"]
        )

        st.altair_chart(
            chart,
            use_container_width=True
        )

    with col2:

        display = engagement.copy()

        display.columns = [
            "Engagement Status",
            "Customers",
            "Churned Customers",
            "Churn Rate (%)"
        ]

        display["Churn Rate (%)"] = (
            display["Churn Rate (%)"].round(2)
        )

        render_table(display)

    active_rate = float(
        engagement.loc[
            engagement["Status"] == "Active",
            "Churn Rate"
        ].iloc[0]
    )

    inactive_rate = float(
        engagement.loc[
            engagement["Status"] == "Inactive",
            "Churn Rate"
        ].iloc[0]
    )

    st.markdown(
        f"""
        <div class="insight-box">
        <div class="insight-title">💡 Engagement Insight</div>
        Active customers recorded an observed churn rate of
        <b>{active_rate:.2f}%</b>, compared with
        <b>{inactive_rate:.2f}%</b> for inactive customers.
        </div>
        """,
        unsafe_allow_html=True
    )

# ============================================================
# PRODUCT UTILIZATION
# ============================================================

elif page == "Product Utilization":

    st.markdown(
        '<div class="section-title">'
        'Product Utilization & Churn'
        '</div>',
        unsafe_allow_html=True
    )

    product = (
        data.groupby("NumOfProducts")["Exited"]
        .agg(["count", "sum"])
        .reset_index()
    )

    product["Churn Rate"] = (
        product["sum"] /
        product["count"] *
        100
    )

    product["Product Group"] = (
        product["NumOfProducts"].astype(str)
        + " Product(s)"
    )

    col1, col2 = st.columns([1.5, 1])

    with col1:

        chart_df = product.copy()
        chart_df["Churn Rate"] = chart_df["Churn Rate"].round(2)

        chart = create_bar_chart(
            chart_df,
            "Product Group",
            "Churn Rate",
            "Number of Products",
            "Observed Churn Rate (%)",
            "#7A5AF8"
        )

        st.altair_chart(
            chart,
            use_container_width=True
        )

    with col2:

        display = product[
            [
                "NumOfProducts",
                "count",
                "sum",
                "Churn Rate"
            ]
        ].copy()

        display.columns = [
            "Number of Products",
            "Customers",
            "Churned Customers",
            "Churn Rate (%)"
        ]

        display["Churn Rate (%)"] = (
            display["Churn Rate (%)"].round(2)
        )

        render_table(display)

    st.markdown(
        """
        <div class="note-box">
        <b>⚠ Interpretation Note:</b>
        The 3-product and 4-product groups are much smaller than
        the 1-product and 2-product groups. Their observed churn
        rates should therefore be interpreted cautiously.
        </div>
        """,
        unsafe_allow_html=True
    )

# ============================================================
# FINANCIAL COMMITMENT
# ============================================================

elif page == "Financial Commitment":

    st.markdown(
        '<div class="section-title">'
        'Financial Commitment vs Engagement'
        '</div>',
        unsafe_allow_html=True
    )

    balance_threshold = data["Balance"].quantile(0.75)

    data_fc = data.copy()

    data_fc["Balance Group"] = data_fc["Balance"].apply(
        lambda x:
        "High Balance"
        if x >= balance_threshold
        else "Other Balance"
    )

    financial = (
        data_fc.groupby(
            ["Balance Group", "Status"]
        )["Exited"]
        .agg(["count", "sum"])
        .reset_index()
    )

    financial["Churn Rate"] = (
        financial["sum"] /
        financial["count"] *
        100
    )

    financial["Group"] = (
        financial["Balance Group"]
        + " - "
        + financial["Status"]
    )

    st.markdown(
        f"""
        <div class="info-box">
        <b>High-Balance Definition:</b>
        Customers at or above the 75th percentile of balance,
        ₹{balance_threshold:,.2f}.
        </div>
        """,
        unsafe_allow_html=True
    )

    chart_df = financial.copy()
    chart_df["Churn Rate"] = chart_df["Churn Rate"].round(2)

    chart = create_bar_chart(
        chart_df,
        "Group",
        "Churn Rate",
        "Balance / Engagement Group",
        "Observed Churn Rate (%)",
        "#2F75B5"
    )

    st.altair_chart(
        chart,
        use_container_width=True
    )

    display = financial[
        [
            "Balance Group",
            "Status",
            "count",
            "sum",
            "Churn Rate"
        ]
    ].copy()

    display.columns = [
        "Balance Group",
        "Engagement Status",
        "Customers",
        "Churned Customers",
        "Churn Rate (%)"
    ]

    display["Churn Rate (%)"] = (
        display["Churn Rate (%)"].round(2)
    )

    render_table(display)

# ============================================================
# SALARY & BALANCE
# ============================================================

elif page == "Salary & Balance":

    st.markdown(
        '<div class="section-title">'
        'Salary–Balance Analysis'
        '</div>',
        unsafe_allow_html=True
    )

    salary_threshold = (
        data["EstimatedSalary"].quantile(0.75)
    )

    balance_threshold = (
        data["Balance"].quantile(0.75)
    )

    sb = data.copy()

    sb["Salary Group"] = sb["EstimatedSalary"].apply(
        lambda x:
        "High Salary"
        if x >= salary_threshold
        else "Other Salary"
    )

    sb["Balance Group"] = sb["Balance"].apply(
        lambda x:
        "Higher Balance"
        if x >= balance_threshold
        else "Low Balance"
    )

    salary_balance = (
        sb.groupby(
            ["Salary Group", "Balance Group"]
        )["Exited"]
        .agg(["count", "sum"])
        .reset_index()
    )

    salary_balance["Churn Rate"] = (
        salary_balance["sum"] /
        salary_balance["count"] *
        100
    )

    salary_balance["Group"] = (
        salary_balance["Salary Group"]
        + " - "
        + salary_balance["Balance Group"]
    )

    st.markdown(
        f"""
        <div class="info-box">
        <b>Thresholds:</b>
        High Salary ≥ ₹{salary_threshold:,.2f}
        &nbsp; | &nbsp;
        Higher Balance ≥ ₹{balance_threshold:,.2f}
        </div>
        """,
        unsafe_allow_html=True
    )

    chart_df = salary_balance.copy()
    chart_df["Churn Rate"] = (
        chart_df["Churn Rate"].round(2)
    )

    chart = create_bar_chart(
        chart_df,
        "Group",
        "Churn Rate",
        "Salary / Balance Group",
        "Observed Churn Rate (%)",
        "#F39C12"
    )

    st.altair_chart(
        chart,
        use_container_width=True
    )

    display = salary_balance[
        [
            "Salary Group",
            "Balance Group",
            "count",
            "sum",
            "Churn Rate"
        ]
    ].copy()

    display.columns = [
        "Salary Group",
        "Balance Group",
        "Customers",
        "Churned Customers",
        "Churn Rate (%)"
    ]

    display["Churn Rate (%)"] = (
        display["Churn Rate (%)"].round(2)
    )

    render_table(display)

# ============================================================
# RETENTION STRENGTH
# ============================================================

elif page == "Retention Strength":

    st.markdown(
        '<div class="section-title">'
        'Retention Strength Analysis'
        '</div>',
        unsafe_allow_html=True
    )

    strength = (
        data.groupby("Relationship Strength")["Exited"]
        .agg(["count", "sum"])
        .reset_index()
    )

    strength["Churn Rate"] = (
        strength["sum"] /
        strength["count"] *
        100
    )

    strength["Strength Score"] = (
        strength["Relationship Strength"]
    )

    chart_df = strength.copy()
    chart_df["Churn Rate"] = (
        chart_df["Churn Rate"].round(2)
    )

    chart = create_bar_chart(
        chart_df,
        "Strength Score",
        "Churn Rate",
        "Relationship Strength Score",
        "Observed Churn Rate (%)",
        "#7A5AF8"
    )

    st.altair_chart(
        chart,
        use_container_width=True
    )

    display = strength[
        [
            "Strength Score",
            "count",
            "sum",
            "Churn Rate"
        ]
    ].copy()

    display.columns = [
        "Strength Score",
        "Customers",
        "Churned Customers",
        "Churn Rate (%)"
    ]

    display["Churn Rate (%)"] = (
        display["Churn Rate (%)"].round(2)
    )

    render_table(display)

    st.markdown(
        """
        <div class="note-box">
        <b>Definition:</b>
        Relationship Strength is a project-defined proxy calculated
        as <b>IsActiveMember + NumOfProducts</b>.
        It is not an industry-standard metric.
        </div>
        """,
        unsafe_allow_html=True
    )

# ============================================================
# GEOGRAPHY
# ============================================================

elif page == "Geography":

    st.markdown(
        '<div class="section-title">'
        'Geographic Churn Analysis'
        '</div>',
        unsafe_allow_html=True
    )

    geography = (
        data.groupby("Geography")["Exited"]
        .agg(["count", "sum"])
        .reset_index()
    )

    geography["Churn Rate"] = (
        geography["sum"] /
        geography["count"] *
        100
    )

    chart_df = geography.copy()
    chart_df["Churn Rate"] = (
        chart_df["Churn Rate"].round(2)
    )

    chart = create_bar_chart(
        chart_df,
        "Geography",
        "Churn Rate",
        "Country",
        "Observed Churn Rate (%)",
        "#2F75B5"
    )

    st.altair_chart(
        chart,
        use_container_width=True
    )

    display = geography.copy()

    display.columns = [
        "Geography",
        "Customers",
        "Churned Customers",
        "Churn Rate (%)"
    ]

    display["Churn Rate (%)"] = (
        display["Churn Rate (%)"].round(2)
    )

    render_table(display)

    st.markdown(
        '<div class="small-title">'
        'Geography & Engagement'
        '</div>',
        unsafe_allow_html=True
    )

    geo_engagement = (
        data.groupby(
            ["Geography", "Status"]
        )["Exited"]
        .agg(["count", "sum"])
        .reset_index()
    )

    geo_engagement["Churn Rate"] = (
        geo_engagement["sum"] /
        geo_engagement["count"] *
        100
    )

    geo_engagement["Group"] = (
        geo_engagement["Geography"]
        + " - "
        + geo_engagement["Status"]
    )

    chart_df = geo_engagement.copy()
    chart_df["Churn Rate"] = (
        chart_df["Churn Rate"].round(2)
    )

    chart = create_bar_chart(
        chart_df,
        "Group",
        "Churn Rate",
        "Geography / Engagement",
        "Observed Churn Rate (%)",
        "#28A745"
    )

    st.altair_chart(
        chart,
        use_container_width=True
    )

# ============================================================
# PRIORITY RISK
# ============================================================

elif page == "Priority Risk":

    st.markdown(
        '<div class="section-title">'
        'Priority Risk Analysis'
        '</div>',
        unsafe_allow_html=True
    )

    priority = (
        data.groupby("Priority Risk")["Exited"]
        .agg(["count", "sum"])
        .reset_index()
    )

    priority["Churn Rate"] = (
        priority["sum"] /
        priority["count"] *
        100
    )

    chart_df = priority.copy()
    chart_df["Churn Rate"] = (
        chart_df["Churn Rate"].round(2)
    )

    chart = create_bar_chart(
        chart_df,
        "Priority Risk",
        "Churn Rate",
        "Risk Group",
        "Observed Churn Rate (%)",
        "#D9534F"
    )

    st.altair_chart(
        chart,
        use_container_width=True
    )

    display = priority.copy()

    display.columns = [
        "Priority Risk Group",
        "Customers",
        "Churned Customers",
        "Churn Rate (%)"
    ]

    display["Churn Rate (%)"] = (
        display["Churn Rate (%)"].round(2)
    )

    render_table(display)

    st.markdown(
        '<div class="small-title">'
        'Priority Risk by Geography'
        '</div>',
        unsafe_allow_html=True
    )

    priority_geo = (
        data[
            data["Priority Risk"] == "Priority Risk"
        ]
        .groupby("Geography")["Exited"]
        .agg(["count", "sum"])
        .reset_index()
    )

    priority_geo["Churn Rate"] = (
        priority_geo["sum"] /
        priority_geo["count"] *
        100
    )

    chart_df = priority_geo.copy()
    chart_df["Churn Rate"] = (
        chart_df["Churn Rate"].round(2)
    )

    chart = create_bar_chart(
        chart_df,
        "Geography",
        "Churn Rate",
        "Geography",
        "Observed Churn Rate (%)",
        "#D9534F"
    )

    st.altair_chart(
        chart,
        use_container_width=True
    )

    display = priority_geo.copy()

    display.columns = [
        "Geography",
        "Priority Customers",
        "Churned Customers",
        "Churn Rate (%)"
    ]

    display["Churn Rate (%)"] = (
        display["Churn Rate (%)"].round(2)
    )

    render_table(display)

    st.markdown(
        """
        <div class="note-box">
        <b>Definition:</b>
        Priority Risk is a project-defined rule:
        <b>Inactive + 1 Product + Age 40 or above</b>.
        </div>
        """,
        unsafe_allow_html=True
    )

# ============================================================
# MACHINE LEARNING RISK
# ============================================================

elif page == "ML Risk Analysis":

    st.markdown(
        '<div class="section-title">'
        'Machine-Learning Churn Risk Analysis'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        f"""
        <div class="info-box">
        <b>Prediction Dataset:</b>
        {test_set_customers:,} customers.
        The risk results below refer to this prediction/test dataset.
        </div>
        """,
        unsafe_allow_html=True
    )

    risk_summary = (
        predictions.groupby("Risk_Flag")["Actual_Exited"]
        .agg(["count", "sum"])
        .reset_index()
    )

    risk_summary["Observed Churn Rate"] = (
        risk_summary["sum"] /
        risk_summary["count"] *
        100
    )

    risk_summary["Observed Churn Rate"] = (
        risk_summary["Observed Churn Rate"].round(2)
    )

    col1, col2 = st.columns([1.4, 1])

    with col1:

        chart = create_bar_chart(
            risk_summary,
            "Risk_Flag",
            "Observed Churn Rate",
            "Risk Group",
            "Observed Churn Rate (%)",
            "#D9534F"
        )

        st.altair_chart(
            chart,
            use_container_width=True
        )

    with col2:

        display = risk_summary.copy()

        display.columns = [
            "Risk Group",
            "Customers",
            "Churned Customers",
            "Observed Churn Rate (%)"
        ]

        render_table(display)

    st.markdown(
        '<div class="small-title">'
        'Churn Probability Distribution'
        '</div>',
        unsafe_allow_html=True
    )

    probability_bins = pd.cut(
        predictions["Churn_Probability"],
        bins=[0, 0.2, 0.4, 0.6, 0.8, 1],
        labels=[
            "0–20%",
            "20–40%",
            "40–60%",
            "60–80%",
            "80–100%"
        ],
        include_lowest=True
    )

    probability_table = (
        probability_bins
        .value_counts()
        .sort_index()
        .rename_axis("Probability Range")
        .reset_index(name="Customers")
    )

    chart = create_bar_chart(
        probability_table,
        "Probability Range",
        "Customers",
        "Churn Probability",
        "Customers",
        "#2F75B5"
    )

    st.altair_chart(
        chart,
        use_container_width=True
    )

    st.markdown(
        """
        <div class="note-box">
        <b>Interpretation:</b>
        Model risk represents predicted churn probability used for
        customer prioritization. It should not be interpreted as a
        guarantee of future churn.
        </div>
        """,
        unsafe_allow_html=True
    )

# ============================================================
# CUSTOMER LOOKUP
# ============================================================

elif page == "Customer Lookup":

    st.markdown(
        '<div class="section-title">'
        'Customer-Level Churn Risk Lookup'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        f"""
        <div class="info-box">
        Search is available for the {test_set_customers:,}
        customers contained in the prediction/test dataset.
        </div>
        """,
        unsafe_allow_html=True
    )

    customer_ids = (
        predictions["CustomerId"]
        .astype(str)
        .tolist()
    )

    selected_id = st.selectbox(
        "Select Customer ID",
        customer_ids
    )

    customer_result = predictions[
        predictions["CustomerId"].astype(str)
        == selected_id
    ]

    if not customer_result.empty:

        row = customer_result.iloc[0]

        probability = float(
            row["Churn_Probability"]
        )

        risk = str(
            row["Risk_Flag"]
        )

        actual = int(
            row["Actual_Exited"]
        )

        col1, col2, col3 = st.columns(3)

        with col1:

            st.markdown(
                f"""
                <div class="kpi-card orange">
                    <div class="kpi-title">Churn Probability</div>
                    <div class="kpi-value">
                        {probability * 100:.2f}%
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

        with col2:

            risk_class = (
                "red"
                if risk == "High Risk"
                else "green"
            )

            st.markdown(
                f"""
                <div class="kpi-card {risk_class}">
                    <div class="kpi-title">Risk Group</div>
                    <div class="kpi-value">
                        {html.escape(risk)}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

        with col3:

            actual_text = (
                "Churned"
                if actual == 1
                else "Retained"
            )

            actual_class = (
                "red"
                if actual == 1
                else "green"
            )

            st.markdown(
                f"""
                <div class="kpi-card {actual_class}">
                    <div class="kpi-title">
                        Actual Churn Status
                    </div>
                    <div class="kpi-value">
                        {actual_text}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

        display = customer_result.copy()

        display["Churn_Probability"] = (
            display["Churn_Probability"]
            .round(4)
        )

        display.columns = [
            "Customer ID",
            "Actual Exited",
            "Churn Probability",
            "Risk Flag"
        ]

        render_table(display)

# ============================================================
# FOOTER
# ============================================================

st.markdown(
    '<div class="footer">'
    'Customer Retention Intelligence | MBA Finance Internship Project'
    '</div>',
    unsafe_allow_html=True
)