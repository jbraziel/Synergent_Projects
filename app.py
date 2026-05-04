import streamlit as st
from datetime import date
from pathlib import Path
import generate_proposal as gp
from database import initialize_database, save_proposal, search_proposals, load_proposal, delete_proposal, update_proposal_status, lock_proposal, unlock_proposal
import os
import json

st.set_page_config(page_title="Marketing Proposal Generator", layout="wide")

initialize_database()

# -----------------------------
# CSS
# -----------------------------
st.markdown(
    """
<style>
    html, body, [class*="css"] {
        font-family: 'Montserrat', sans-serif;
    }
    
    .stApp {
        background-color: #f7f9f5;
    }
    
    h1, h2, h3 {
        color: #2f3a2f;
    }
    
    section[data-testid="stSidebar"] {
        background-color: #eef6e8;
    }
    
    section[data-testid="stSidebar"] button {
       background-color: white !important;
       color: #2f3a2f !important;
       border: 1px solid #d5e6cc !important;
       border-radius: 10px !important;
       padding: 0.75rem 1rem !important;
       margin-bottom: 0.35rem !important;
       font-weight: 600 !important;
    }
    
    section[data-testid="stSidebar"] button:hover {
       background-color: #e8f4df !important;
       border-color: #76bd22 !important;
    }
    
    div[data-testid="stHorizontalBlock"] {
        gap: 0.25rem !important;
    }
    
    .proposal-table-header {
        background-color: #eef6e8;
        padding: 8px 10px;
        border-radius: 8px;
        font-weight: 700;
        margin-bottom: 6px;
    }
    
    .proposal-row {
        background-color: white;
        padding: 6px 10px;
        border-radius: 8px;
        border: 1px solid #e3e8df;
        margin-bottom: 6px;
    }
    
    .proposal-row:hover {
       background-color: #f0f8ea;
       border-color: #76bd22;
    }
    
    strong {
        font-size: 13px;
    }
    
    hr {
        border: none;
        border-top: 1px solid #e5e5e5;
    }
    
    div[data-testid="stSelectbox"] > div {
        min-width: 90px !important;
    }
    
    div[data-testid="stSelectbox"] div[data-baseweb="select"] {
        font-size: 12px !important;
    }
    
    /* Button layout */
    div[data-testid="stButton"] {
        margin: 0 !important;
        display: flex !important;
        align-items: center !important;
        height: 100% !important;
    }
    
    div[data-testid="stButton"] button,
    div[data-testid="stDownloadButton"] button {
        height: 38px !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 12px !important;
        border: none !important;
    }
    
    /* Normal buttons: OPEN and cycle */
    div[data-testid="stButton"] button[kind="secondary"] {
        background-color: #76bd22 !important;
        color: white !important;
    }
    
    div[data-testid="stButton"] button[kind="secondary"]:hover {
        background-color: #5f9f1b !important;
        color: white !important;
    }
    
    /* Delete button */
    div[data-testid="stButton"] button[kind="primary"] {
        background-color: #c0392b !important;
        color: white !important;
    }
    
    div[data-testid="stButton"] button[kind="primary"]:hover {
        background-color: #a93226 !important;
        color: white !important;
    }
    
    /* Download button */
    div[data-testid="stDownloadButton"] button {
        background-color: #1f77d0 !important;
        color: white !important;
    }
    
    div[data-testid="stDownloadButton"] button:hover {
        background-color: #155a9c !important;
        color: white !important;
    }
    
    /* Small icon-style buttons */
    button[kind="secondary"],
    button[kind="primary"],
    div[data-testid="stDownloadButton"] button {
        min-width: 28px !important;
        padding: 0.15rem 0.35rem !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
    }
</style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------
# Constants
# -----------------------------
TEMPLATE_MAP = {
    "Auto Loan Recapture Campaign": "ACH_Auto_Proposal_Template.docx",
    "Credit Card Campaign": "Credit_Card_Proposal_Template.docx",
    "Home Lending Campaign": "Home_Lending_Proposal_Template.docx",
    "Consumer Loan Campaign": "Consumer_Loan_Proposal_Template.docx",
    "Checking Campaign": "Checking_Proposal_Template.docx",
    "Certificate Campaign": "CD_Proposal_Template.docx",
    "Synergent Email Platform Proposal": "Email_Platform_Proposal_Template.docx",
}

DEFAULT_COMPONENTS = [
    "Creative concept, strategy and design",
    "Preliminary data analysis",
    "Custom programmed data extract for mailing",
    "Proofing and testing",
    "Tracking, monitoring and reporting",
    "Mailing preparation and presorting",
    "Content development",
    "Responsive email template development",
    "Digital graphic assets / social media graphics",
    "5.5” x 8.5” full color variable postcards",
    "Unique URL and QR Code redirect for 12 months",
    "Optional call file for personal outreach / follow-up",
]

DEFAULT_TARGET_OPTIONS = [
    (910, "members making an ACH auto loan payment to another financial institution"),
    (472, "members making a recurring ACH payment between $400-$800 to another financial institution"),
    (628, "members who paid off their auto loan in the last 12 months"),
    (245, "members due to pay off their auto loan in the next 12 months"),
    (2047, "checking members who have a loan but no auto loan with the credit union"),
    (350, "members with high checking activity and no current auto loan"),
    (525, "members with direct deposit and no recent lending relationship"),
    (700, "members with external loan payment indicators"),
    (300, "members with prior auto loan history but no current auto loan"),
    (425, "members with strong product engagement and lending opportunity"),
]

STRAIGHT_COST_ITEMS = {
    "Preliminary Data Analysis": 100,
    "Strategy & Concept": 300,
    "Copy Writing": 100,
    "Unique URL": 150,
    "Tracking, Monitoring & Reporting": 100,
    "Campaign Management": 100,
    "Consultative Implementation": 500,
    "QR Code": 100,
    "ACH Program Fee": 2500,
}

REQUIRED_SECTIONS = [
    "Proposal Details",
    "Campaign Targets",
    "Conversion Metrics",
    "Campaign Components",
    "Cost Estimator",
]


# -----------------------------
# Helper functions
# -----------------------------

def auto_save_proposal():
    # Only auto-save after a proposal has been created/saved once
    if not st.session_state.get("current_proposal_id"):
        return

    saved_data = collect_saved_data()

    proposal_id = save_proposal(
        st.session_state.get("current_proposal_id"),
        st.session_state.proposal_name,
        st.session_state.credit_union,
        st.session_state.proposal_type,
        st.session_state.proposal_status,
        saved_data,
        st.session_state.msr,
        st.session_state.current_user
    )

    st.session_state.current_proposal_id = proposal_id

def collect_saved_data():
    data = {}

    keys_to_save = [
        "current_proposal_id",
        "proposal_status",
        "proposal_type",
        "proposal_name",
        "credit_union",
        "proposal_date",
        "conversion_rate",
        "avg_auto_balance",
        "auto_interest_rate",
        "auto_term_years",
        "custom_targets",
        "custom_components",
        "custom_costs",
        "include_creative",
        "creative_hours",
        "include_data_mining",
        "data_mining_hours",
        "include_list_procurement",
        "list_procurement_raw",
        "include_print",
        "print_raw",
        "include_email_labor",
        "email_hours",
        "include_email_sends",
        "email_send_count",
        "msr",
        "file_path"
    ]

    for key in keys_to_save:
        value = st.session_state.get(key)

        if key == "proposal_date" and value is not None:
            value = value.isoformat()

        data[key] = value

    for i, _ in enumerate(DEFAULT_TARGET_OPTIONS, start=1):
        data[f"target_include_saved_{i}"] = st.session_state.get(f"target_include_saved_{i}", i <= 5)
        data[f"target_count_saved_{i}"] = st.session_state.get(f"target_count_saved_{i}", DEFAULT_TARGET_OPTIONS[i - 1][0])

    for i, _ in enumerate(DEFAULT_COMPONENTS, start=1):
        data[f"component_saved_{i}"] = st.session_state.get(f"component_saved_{i}", True)

    for name in STRAIGHT_COST_ITEMS:
        data[f"cost_{name}"] = st.session_state.get(f"cost_{name}", True)

    for sec in REQUIRED_SECTIONS:
        data[f"complete_{sec}"] = st.session_state.get(f"complete_{sec}", False)

    return data


def load_saved_data(data):
    for key, value in data.items():
        if key == "proposal_date" and value:
            value = date.fromisoformat(value)

        st.session_state[key] = value

def section_complete_checkbox(section_name):
    saved_key = f"complete_{section_name}"
    widget_key = f"widget_{saved_key}"

    if saved_key not in st.session_state:
        st.session_state[saved_key] = False

    if widget_key not in st.session_state:
        st.session_state[widget_key] = st.session_state[saved_key]

    value = st.checkbox(
        "Mark this section complete",
        key=widget_key
    )

    st.session_state[saved_key] = value

def persistent_checkbox(label, saved_key, default=True):
    widget_key = f"widget_{saved_key}"

    if saved_key not in st.session_state:
        st.session_state[saved_key] = default

    if widget_key not in st.session_state:
        st.session_state[widget_key] = st.session_state[saved_key]

    checked = st.checkbox(label, key=widget_key)
    st.session_state[saved_key] = checked

    return checked


def section_status(section_name):
    selected_targets = get_selected_targets()
    selected_components = get_selected_components()
    costs = calculate_costs()

    if section_name == "Proposal Details":
        complete = (
            bool(st.session_state.proposal_name.strip())
            and bool(st.session_state.credit_union.strip())
            and st.session_state.proposal_type in TEMPLATE_MAP
        )

    elif section_name == "Campaign Targets":
        complete = len(selected_targets) > 0 and sum(c for c, _ in selected_targets) > 0

    elif section_name == "Conversion Metrics":
        complete = (
            parse_percent(st.session_state.conversion_rate) > 0
            and st.session_state.avg_auto_balance > 0
            and parse_percent(st.session_state.auto_interest_rate) > 0
            and st.session_state.auto_term_years > 0
        )

    elif section_name == "Campaign Components":
        complete = len(selected_components) > 0

    elif section_name == "Cost Estimator":
        complete = costs["campaign_1_calc"] > 0

    elif section_name == "Generate Proposal":
        complete = False

    else:
        complete = False

    return "✅" if complete else "⚠️"

def parse_percent(percent_text):
    clean = str(percent_text).replace("%", "").strip()
    try:
        return float(clean) / 100
    except ValueError:
        return 0


def format_date_windows(d):
    return d.strftime("%B %#d, %Y")


def money(value):
    return f"${value:,.0f}"


def calculate_first_year_interest(principal, annual_rate, years):
    if principal <= 0 or annual_rate <= 0 or years <= 0:
        return 0

    monthly_rate = annual_rate / 12
    total_payments = years * 12

    monthly_payment = principal * (
        (monthly_rate * (1 + monthly_rate) ** total_payments)
        / ((1 + monthly_rate) ** total_payments - 1)
    )

    balance = principal
    total_interest = 0

    for _ in range(12):
        interest = balance * monthly_rate
        principal_paid = monthly_payment - interest
        balance -= principal_paid
        total_interest += interest

    return round(total_interest)

def persistent_checkbox(label, saved_key, default=True):
    widget_key = f"widget_{saved_key}"

    if saved_key not in st.session_state:
        st.session_state[saved_key] = default

    if widget_key not in st.session_state:
        st.session_state[widget_key] = st.session_state[saved_key]

    checked = st.checkbox(label, key=widget_key)

    st.session_state[saved_key] = checked

    return checked


def init_state():
    defaults = {
        "active_section": "Proposal Library",
        "proposal_type": "Auto Loan Recapture Campaign",
        "proposal_name": "ACH Auto Loan Recapture Campaign Proposal",
        "credit_union": "Sample Credit Union",
        "proposal_date": date.today(),
        "conversion_rate": "2%",
        "avg_auto_balance": 18500,
        "auto_interest_rate": "6.00%",
        "auto_term_years": 5,
        "custom_targets": [],
        "custom_components": [],
        "custom_costs": [],
        "include_creative": True,
        "creative_hours": 5.0,
        "include_data_mining": True,
        "data_mining_hours": 3.0,
        "include_list_procurement": False,
        "list_procurement_raw": 0.0,
        "include_print": True,
        "print_raw": 1500.0,
        "include_email_labor": True,
        "email_hours": 2.0,
        "include_email_sends": True,
        "email_send_count": 1,
        "complete_Proposal Details": False,
        "complete_Campaign Targets": False,
        "complete_Conversion Metrics": False,
        "complete_Campaign Components": False,
        "complete_Cost Estimator": False,
        "current_proposal_id": None,
        "proposal_status": "Draft",
        "msr": "",
        "current_user": ""
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

    for i, (_, _) in enumerate(DEFAULT_TARGET_OPTIONS, start=1):
        if f"target_include_{i}" not in st.session_state:
            st.session_state[f"target_include_{i}"] = i <= 5
        if f"target_count_{i}" not in st.session_state:
            st.session_state[f"target_count_{i}"] = DEFAULT_TARGET_OPTIONS[i - 1][0]

    for i, _ in enumerate(DEFAULT_COMPONENTS, start=1):
        if f"component_{i}" not in st.session_state:
            st.session_state[f"component_{i}"] = True

    for name in STRAIGHT_COST_ITEMS:
        if f"cost_{name}" not in st.session_state:
            st.session_state[f"cost_{name}"] = True


def get_selected_targets():
    selected_targets = []

    for i, (_, description) in enumerate(DEFAULT_TARGET_OPTIONS, start=1):
        include = st.session_state.get(f"target_include_saved_{i}", i <= 5)
        count = st.session_state.get(f"target_count_saved_{i}", DEFAULT_TARGET_OPTIONS[i - 1][0])

        if include:
            selected_targets.append((int(count), description))

    for item in st.session_state.get("custom_targets", []):
        count = item.get("count", 0)
        description = item.get("description", "")

        if count > 0 and description.strip():
            selected_targets.append((int(count), description.strip()))

    return selected_targets


def get_selected_components():
    selected_components = []

    for i, component in enumerate(DEFAULT_COMPONENTS, start=1):
        include = st.session_state.get(f"component_saved_{i}", True)

        if include:
            selected_components.append(component)

    for component in st.session_state.get("custom_components", []):
        if component.strip():
            selected_components.append(component.strip())

    return selected_components

def format_status(status):
    colors = {
        "Draft": "#6c757d",        # gray
        "CU Review": "#3555bd",    # blue
        "Signed": "#76bd22",       # green
        "Declined": "#dc3545"      # red
    }
    color = colors.get(status, "black")
    return f"<span style='color:{color}; font-weight:600'>{status}</span>"

def status_color(status):
    colors = {
        "Draft": "#6c757d",
        "CU Review": "#3555bd",
        "Signed": "#76bd22",
        "Declined": "#dc3545"
    }
    return colors.get(status, "black")


def calculate_costs():
    markup_rate = 1.35

    creative_cost = (
        st.session_state.creative_hours * 105
        if st.session_state.include_creative
        else 0
    )

    data_mining_cost = (
        st.session_state.data_mining_hours * 200
        if st.session_state.include_data_mining
        else 0
    )

    list_procurement_cost = (
        st.session_state.list_procurement_raw * markup_rate
        if st.session_state.include_list_procurement
        else 0
    )

    print_cost = (
        st.session_state.print_raw * markup_rate
        if st.session_state.include_print
        else 0
    )

    email_labor_cost = (
        st.session_state.email_hours * 105
        if st.session_state.include_email_labor
        else 0
    )

    email_send_cost = (
        st.session_state.email_send_count * 100
        if st.session_state.include_email_sends
        else 0
    )

    straight_cost_total = 0
    selected_straight_costs = []

    for name, cost in STRAIGHT_COST_ITEMS.items():
        if st.session_state.get(f"cost_{name}", True):
            straight_cost_total += cost
            selected_straight_costs.append((name, cost))

    custom_costs_total = 0

    for item in st.session_state.get("custom_costs", []):
        if item.get("name", "").strip() and item.get("amount", 0) > 0:
            custom_costs_total += item["amount"]

    estimated_cost_total = (
        creative_cost
        + data_mining_cost
        + list_procurement_cost
        + print_cost
        + email_labor_cost
        + email_send_cost
        + straight_cost_total
        + custom_costs_total
    )

    ach_program_fee = (
        2500
        if any(name == "ACH Program Fee" for name, _ in selected_straight_costs)
        else 0
    )

    repeating_cost_total = (
        print_cost
        + list_procurement_cost
        + email_send_cost
        + ach_program_fee
    )

    one_time_cost_total = estimated_cost_total - repeating_cost_total

    campaign_1_calc = estimated_cost_total
    campaign_2_calc = one_time_cost_total + (repeating_cost_total * 2)
    campaign_2_per_calc = campaign_2_calc / 2
    campaign_4_calc = (one_time_cost_total + (repeating_cost_total * 4)) * 0.90
    campaign_4_per_calc = campaign_4_calc / 4

    return {
        "creative_cost": creative_cost,
        "data_mining_cost": data_mining_cost,
        "list_procurement_cost": list_procurement_cost,
        "print_cost": print_cost,
        "email_labor_cost": email_labor_cost,
        "email_send_cost": email_send_cost,
        "straight_cost_total": straight_cost_total,
        "custom_costs_total": custom_costs_total,
        "estimated_cost_total": estimated_cost_total,
        "one_time_cost_total": one_time_cost_total,
        "repeating_cost_total": repeating_cost_total,
        "campaign_1_calc": campaign_1_calc,
        "campaign_2_calc": campaign_2_calc,
        "campaign_2_per_calc": campaign_2_per_calc,
        "campaign_4_calc": campaign_4_calc,
        "campaign_4_per_calc": campaign_4_per_calc,
    }


init_state()


# -----------------------------
# Sidebar
# -----------------------------
with st.sidebar:
    st.image("logo.png", width=150)

    # Initialize if not set
    if "current_user" not in st.session_state:
        st.session_state.current_user = ""
    
    user = st.selectbox(
        "Who is using the proposal tool?",
        ["", "Jen", "Erica", "Shannan", "Doug"],
        index=["", "Jen", "Erica", "Shannan", "Doug"].index(st.session_state.current_user)
        if st.session_state.current_user in ["", "Jen", "Erica", "Shannan", "Doug"]
        else 0
    )

    st.session_state.current_user = user
    
    # Stop app until user is selected
    if not st.session_state.current_user:
        st.warning("Please select your name to continue")
        st.stop()

    st.caption(f"Editing: {st.session_state.proposal_name}")

    st.markdown("---")

    if st.session_state.get("current_proposal_id"):
       st.markdown("### Current Proposal")

       st.caption(st.session_state.get("proposal_name", ""))

    if st.button("Close Proposal"):
        unlock_proposal(
            st.session_state.current_proposal_id,
            st.session_state.current_user
        )

        st.session_state.current_proposal_id = None
        st.session_state.active_section = "Proposal Library"
        st.rerun()

    st.markdown("---")

    if st.button("📁 Proposal Library", key="nav_library", use_container_width=True):
        st.session_state.active_section = "Proposal Library"
        st.rerun()

    if st.session_state.active_section != "Proposal Library":
        st.markdown("### Proposal Workflow")

        workflow_sections = [
            "Proposal Details",
            "Campaign Targets",
            "Conversion Metrics",
            "Campaign Components",
            "Cost Estimator",
            "Generate Proposal",

        ]

        for sec in workflow_sections:
            selected = st.session_state.active_section == sec

            if sec == "Generate Proposal":
                icon = "🚀"
            else:
                icon = "✅" if st.session_state.get(f"complete_{sec}", False) else "⬜"

            col_icon, col_button = st.columns([0.18, 0.82])

            with col_icon:
                st.markdown(
                    f"<div style='font-size:20px; padding-top:8px; text-align:center;'>{icon}</div>",
                    unsafe_allow_html=True
                )

            with col_button:
                label = f"▶ {sec}" if selected else sec

                if st.button(label, key=f"nav_{sec}", use_container_width=True):
                    st.session_state.active_section = sec
                    
section = st.session_state.active_section

# -----------------------------
# Header
# -----------------------------
col_logo, col_title = st.columns([0.6, 8])

with col_logo:
    st.image("swirl.png", width=75)

with col_title:
    st.markdown(
        """
        <h1 style="margin:0; padding-top:8px; line-height:1.1;">Marketing Proposal Generator</h1>
        <p style="color:#76bd22; margin:0; font-size:16px;">Campaign Planning & Proposal Tool</p>
        """,
        unsafe_allow_html=True,
    )

st.markdown(
    "<hr style='border: 2px solid #76bd22; margin-top: 8px; margin-bottom: 20px;'>",
    unsafe_allow_html=True,
)

# -----------------------------
# Progress Bar
# -----------------------------
REQUIRED_SECTIONS = [
    "Proposal Details",
    "Campaign Targets",
    "Conversion Metrics",
    "Campaign Components",
    "Cost Estimator",
]

completed_sections = [
    sec for sec in REQUIRED_SECTIONS
    if st.session_state.get(f"complete_{sec}", False)
]



# -----------------------------
# Shared calculations
# -----------------------------
selected_targets = get_selected_targets()
selected_components = get_selected_components()
total_targets = sum(count for count, _ in selected_targets)

conversion_rate_decimal = parse_percent(st.session_state.conversion_rate)
estimated_loans_refinanced = round(total_targets * conversion_rate_decimal)
amount_refinanced = estimated_loans_refinanced * st.session_state.avg_auto_balance
auto_interest_rate_decimal = parse_percent(st.session_state.auto_interest_rate)
estimated_first_year_interest = calculate_first_year_interest(
    amount_refinanced,
    auto_interest_rate_decimal,
    st.session_state.auto_term_years,
)

costs = calculate_costs()

# ============================================================
# Proposal Library
# ============================================================
if section == "Proposal Library":
    st.subheader("Proposal Library")

    col1, col2, col3 = st.columns([0.50, 0.25, 0.25])

    with col1:
        search_text = st.text_input("Search by proposal name, credit union, or proposal type")

    with col2:
        status_filter = st.selectbox(
            "Status",
            ["All", "Draft", "CU Review", "Signed", "Declined"]
        )

    with col3:
        msr_filter = st.selectbox(
            "MSR",
            ["All", "Shannan", "Erica"]
        )

    results = search_proposals(search_text, status_filter, msr_filter)

    st.markdown("### Start New Proposal")

    if st.button("Create New Proposal", key="create_new_proposal"):
        st.session_state.current_proposal_id = None
        st.session_state.proposal_status = "Draft"

        for sec in REQUIRED_SECTIONS:
            st.session_state[f"complete_{sec}"] = False
            st.session_state[f"widget_complete_{sec}"] = False

        st.session_state.active_section = "Proposal Details"
        st.rerun()

    st.markdown("---")

    if not results:
        st.info("No saved proposals found.")
    else:
        st.markdown("### Saved Proposals")

        h1, h2, h3, h4, h5, h6, h7, h8 = st.columns(
           [0.20, 0.15, 0.10, 0.10, 0.14, 0.05, 0.04, 0.03]
        )

        with h1: st.markdown("<span style='font-size:15px; font-weight:700;'>Proposal</span>", unsafe_allow_html=True)
        with h2: st.markdown("<span style='font-size:15px; font-weight:700;'>Credit Union</span>", unsafe_allow_html=True)
        with h3: st.markdown("<span style='font-size:15px; font-weight:700;'>Status</span>", unsafe_allow_html=True)
        with h4: st.markdown("<span style='font-size:15px; font-weight:700;'>MSR</span>", unsafe_allow_html=True)
        with h5: st.markdown("<span style='font-size:15px; font-weight:700;'>Last Updated</span>", unsafe_allow_html=True)
        with h6: st.markdown("<span style='font-size:15px; font-weight:700;'>Actions</span>", unsafe_allow_html=True)
        with h7: st.markdown("<span style='font-size:15px; font-weight:700;'></span>", unsafe_allow_html=True)
        with h8: st.markdown("<span style='font-size:15px; font-weight:700;'></span>", unsafe_allow_html=True)
        
        st.markdown("---")

        for proposal_id, proposal_name, credit_union, proposal_type, status, updated_at, msr, updated_by, locked_by, locked_at in results:
            col1, col2, col3, col4, col5, col6, col7, col8 = st.columns(
                [0.20, 0.15, 0.10, 0.10, 0.14, 0.05, 0.02, 0.02]
            )

            with col1:
                st.write(f"**{proposal_name}**")
                
                if locked_by and locked_by != st.session_state.current_user:
                    st.caption(f"🔒 Editing: {locked_by}")

            with col2:
                st.write(credit_union)

            with col3:
                 status_options = ["Draft", "CU Review", "Signed", "Declined"]
             
                 status_colors = {
                     "Draft": "#f0ad4e",
                     "CU Review": "#3555bd",
                     "Signed": "#76bd22",
                     "Declined": "#dc3545",
                 }
             
                 if status not in status_options:
                     status = "Draft"
             
                 next_status = status_options[
                     (status_options.index(status) + 1) % len(status_options)
                 ]
             
                 pill_col, button_col = st.columns([0.60, 0.40])
             
                 with pill_col:
                     st.markdown(
                         f"""
                         <div style="
                             background-color:{status_colors[status]};
                             color:white;
                             border-radius:20px;
                             padding:4px 10px;
                             font-size:15px;
                             font-weight:700;
                             text-align:center;
                             width:90px;
                             line-height:18px;
                             margin-top:4px;
                         ">
                             {status}
                         </div>
                         """,
                         unsafe_allow_html=True
                     )
             
                 with button_col:
                     if st.button(
                         "↻",
                         key=f"cycle_status_{proposal_id}",
                         help=f"Change status to {next_status}"
                     ):
                         update_proposal_status(proposal_id, next_status)
                         st.rerun()

            with col4:
                st.write(msr or "")

            with col5:
                st.caption(f"Last updated: {updated_at}")
                if updated_by:
                    st.caption(f"By: {updated_by}")

            with col6:
                if st.button("OPEN", key=f"open_{proposal_id}", type="secondary", help="Continue editing proposal"):
                    saved_data = load_proposal(proposal_id)

                    if saved_data:
                        lock_proposal(proposal_id, st.session_state.current_user)

                        load_saved_data(saved_data)
                        st.session_state.current_proposal_id = proposal_id
                        st.session_state.active_section = "Proposal Details"
                        st.rerun()

            with col7:
                if st.button(
                    "✕",
                    key=f"delete_{proposal_id}",
                    type="primary",
                    help="Delete this Proposal"
                ):
                    delete_proposal(proposal_id)
                    st.rerun()

            with col8:
                saved_data = load_proposal(proposal_id)

                # If saved_data is stored as JSON text, convert it back to a dictionary
                if isinstance(saved_data, str):
                    import json
                    saved_data = json.loads(saved_data)
            
                file_path = saved_data.get("file_path") if saved_data else None


                if file_path and os.path.exists(file_path):
                    with open(file_path, "rb") as file:
                       st.download_button(
                          label="↓",
                          data=file,
                          file_name=os.path.basename(file_path),
                          mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                          key=f"download_docx_{proposal_id}",
                          help="Download proposal"
                        )
                else:
                     st.caption("No file")

            st.markdown("<hr style='margin: 4px 0;'>", unsafe_allow_html=True)

# ============================================================
# Proposal Details
# ============================================================
elif section == "Proposal Details":
    st.subheader("Proposal Details")

    st.session_state.proposal_type = st.selectbox(
        "Select Proposal Template",
        list(TEMPLATE_MAP.keys()),
        index=list(TEMPLATE_MAP.keys()).index(st.session_state.proposal_type),
    )

    selected_template = TEMPLATE_MAP[st.session_state.proposal_type]

    if not Path(selected_template).exists():
        st.warning(f"Template file not found yet: {selected_template}")

    col1, col2 = st.columns(2)

    with col1:
        st.session_state.proposal_name = st.text_input(
            "Proposal Name",
            st.session_state.proposal_name,
        )

        st.session_state.credit_union = st.text_input(
            "Credit Union Name",
            st.session_state.credit_union,
        )

    with col2:
        st.session_state.proposal_date = st.date_input(
            "Proposal Date",
            st.session_state.proposal_date,
        )

        options = ["Shannan", "Erica"]

        st.session_state.msr = st.selectbox(
            "MSR",
            options,
            index=options.index(st.session_state.get("msr"))
            if st.session_state.get("msr") in options else 0
        )

    section_complete_checkbox("Proposal Details")
    auto_save_proposal()

# ============================================================
# Campaign Targets
# ============================================================
elif section == "Campaign Targets":
    st.subheader("Campaign Targets")
    st.caption("Select target segments, adjust counts, or add custom targets.")

    for i, (_, description) in enumerate(DEFAULT_TARGET_OPTIONS, start=1):
        col_check, col_count, col_desc = st.columns([0.08, 0.18, 0.74])

        with col_check:
            persistent_checkbox(
                "",
                saved_key=f"target_include_saved_{i}",
                default=i <= 5
            )

        with col_count:
             count_key = f"target_count_saved_{i}"

             if count_key not in st.session_state:
                st.session_state[count_key] = DEFAULT_TARGET_OPTIONS[i - 1][0]

             st.session_state[count_key] = st.number_input(
                 "Count",
                 min_value=0,
                 value=st.session_state[count_key],
                 step=1,
                 key=f"widget_{count_key}",
                 label_visibility="collapsed",
            )

        with col_desc:
            st.text(description)

    st.markdown("#### Custom Campaign Targets")

    if st.button("Add Target"):
        st.session_state.custom_targets.append(
            {"count": 0, "description": ""}
        )
        st.rerun()

    updated_custom_targets = []

    for i, item in enumerate(st.session_state.custom_targets):
        col_count, col_desc = st.columns([0.25, 0.75])

        with col_count:
            count = st.number_input(
                f"Custom Target {i + 1} Count",
                min_value=0,
                value=int(item.get("count", 0)),
                step=1,
                key=f"custom_target_count_input_{i}",
            )

        with col_desc:
            description = st.text_input(
                f"Custom Target {i + 1} Description",
                value=item.get("description", ""),
                key=f"custom_target_desc_input_{i}",
                placeholder="Example: members with external auto loan payment indicators",
            )

        updated_custom_targets.append(
            {"count": int(count), "description": description}
        )

    st.session_state.custom_targets = updated_custom_targets

    selected_targets = get_selected_targets()
    total_targets = sum(count for count, _ in selected_targets)
    st.success(f"Total selected targets: {total_targets:,}")

    section_complete_checkbox("Campaign Targets")
    auto_save_proposal()
   

# ============================================================
# Conversion Metrics
# ============================================================
elif section == "Conversion Metrics":
    st.subheader("Estimated Conversion Metrics")

    st.markdown("### Inputs")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.session_state.conversion_rate = st.text_input(
            "Target Conversion Rate",
            st.session_state.conversion_rate,
        )

    with col2:
        st.session_state.avg_auto_balance = st.number_input(
            "Average Balance Refinanced",
            min_value=0,
            value=st.session_state.avg_auto_balance,
            step=500,
        )

    with col3:
        st.session_state.auto_interest_rate = st.text_input(
            "Average Interest Rate",
            st.session_state.auto_interest_rate,
        )

    with col4:
        st.session_state.auto_term_years = st.number_input(
            "Average Term Years",
            min_value=1,
            value=st.session_state.auto_term_years,
            step=1,
        )

    selected_targets = get_selected_targets()
    total_targets = sum(count for count, _ in selected_targets)

    conversion_rate_decimal = parse_percent(st.session_state.conversion_rate)
    estimated_loans_refinanced = round(total_targets * conversion_rate_decimal)
    amount_refinanced = estimated_loans_refinanced * st.session_state.avg_auto_balance
    auto_interest_rate_decimal = parse_percent(st.session_state.auto_interest_rate)
    estimated_first_year_interest = calculate_first_year_interest(
        amount_refinanced,
        auto_interest_rate_decimal,
        st.session_state.auto_term_years,
    )

    st.markdown("### Calculated Results")

    col5, col6, col7 = st.columns(3)

    with col5:
        st.text_input(
            "Estimated Loans Refinanced",
            value=f"{estimated_loans_refinanced:,}",
            disabled=True,
        )

    with col6:
        st.text_input(
            "Estimated Amount Refinanced",
            value=f"${amount_refinanced:,.0f}",
            disabled=True,
        )

    with col7:
        st.text_input(
            "Estimated First-Year Interest",
            value=f"${estimated_first_year_interest:,.0f}",
            disabled=True,
        )

    st.caption(
        "Estimated loans refinanced = total targets × conversion rate. "
        "Estimated first-year interest is calculated using an amortized loan schedule."
    )

    section_complete_checkbox("Conversion Metrics")
    auto_save_proposal()

# ============================================================
# Campaign Components
# ============================================================
elif section == "Campaign Components":
    st.subheader("Campaign Components")
    st.caption("Select standard components or add custom components.")

    for i, component in enumerate(DEFAULT_COMPONENTS, start=1):
       persistent_checkbox(
           component,
           saved_key=f"component_saved_{i}",
           default=True
       )

    st.markdown("#### Custom Campaign Components")

    if st.button("Add Component"):
        st.session_state.custom_components.append("")
        st.rerun()

    updated_components = []

    for i, component in enumerate(st.session_state.custom_components):
        updated_value = st.text_input(
            f"Custom Component {i + 1}",
            value=component,
            key=f"custom_component_input_{i}",
            placeholder="Example: custom landing page development",
        )

        updated_components.append(updated_value)

    st.session_state.custom_components = updated_components

    st.success(f"Selected components: {len(get_selected_components())}")

    section_complete_checkbox("Campaign Components")
    auto_save_proposal()

# ============================================================
# Cost Estimator
# ============================================================
elif section == "Cost Estimator":
    st.subheader("Estimated Costs")
    st.caption("Select internal cost items. These inputs calculate proposal pricing but do not appear in the proposal.")

    st.markdown("### Hourly Cost Items")

    col1, col2 = st.columns(2)

    with col1:
        st.session_state.include_creative = st.checkbox(
            "Creative Concept & Design",
            st.session_state.include_creative,
        )
        st.session_state.creative_hours = st.number_input(
            "Creative Hours",
            min_value=0.0,
            value=st.session_state.creative_hours,
            step=0.5,
        )
        st.caption("Estimated at $105/hour")

    with col2:
        st.session_state.include_data_mining = st.checkbox(
            "Targeted Data Mining",
            st.session_state.include_data_mining,
        )
        st.session_state.data_mining_hours = st.number_input(
            "Programming Hours",
            min_value=0.0,
            value=st.session_state.data_mining_hours,
            step=0.5,
        )
        st.caption("Estimated at $200/hour")

    st.markdown("### Marked-Up Cost Items")
    st.caption("These costs are marked up by 35%.")

    col1, col2 = st.columns(2)

    with col1:
        st.session_state.include_list_procurement = st.checkbox(
            "List Procurement",
            st.session_state.include_list_procurement,
        )
        st.session_state.list_procurement_raw = st.number_input(
            "List Procurement Cost",
            min_value=0.0,
            value=st.session_state.list_procurement_raw,
            step=50.0,
        )

    with col2:
        st.session_state.include_print = st.checkbox(
            "Variable Print Production",
            st.session_state.include_print,
        )
        st.session_state.print_raw = st.number_input(
            "Variable Print Production Cost",
            min_value=0.0,
            value=st.session_state.print_raw,
            step=50.0,
        )

    st.markdown("### Email Costs")

    col1, col2 = st.columns(2)

    with col1:
        st.session_state.include_email_labor = st.checkbox(
            "Email Development Hours",
            st.session_state.include_email_labor,
        )
        st.session_state.email_hours = st.number_input(
            "Email Development Hours",
            min_value=0.0,
            value=st.session_state.email_hours,
            step=0.5,
        )
        st.caption("Estimated at $105/hour")

    with col2:
        st.session_state.include_email_sends = st.checkbox(
            "Email Sends",
            st.session_state.include_email_sends,
        )
        st.session_state.email_send_count = st.number_input(
            "Number of Email Sends",
            min_value=0,
            value=st.session_state.email_send_count,
            step=1,
        )
        st.caption("Charged at $100 per email send.")

    st.markdown("### Straight Cost Items")

    for name, cost in STRAIGHT_COST_ITEMS.items():
        st.checkbox(
            f"{name} — {money(cost)}",
            key=f"cost_{name}",
        )

    st.markdown("### Custom Costs")

    if st.button("Add Custom Cost"):
        st.session_state.custom_costs.append(
            {"name": "", "amount": 0.0}
        )
        st.rerun()

    updated_custom_costs = []

    for i, item in enumerate(st.session_state.custom_costs):
        col1, col2 = st.columns([0.7, 0.3])

        with col1:
            name = st.text_input(
                f"Custom Cost {i + 1} Name",
                value=item.get("name", ""),
                key=f"custom_cost_name_input_{i}",
            )

        with col2:
            amount = st.number_input(
                f"Custom Cost {i + 1} Amount",
                min_value=0.0,
                value=float(item.get("amount", 0.0)),
                step=50.0,
                key=f"custom_cost_amount_input_{i}",
            )

        updated_custom_costs.append(
            {"name": name, "amount": amount}
        )

    st.session_state.custom_costs = updated_custom_costs

    costs = calculate_costs()

    st.markdown("### Calculated Proposal Pricing")

    with st.expander("View Internal Cost Summary"):
        st.markdown("#### Hourly Costs")
        st.write(f"Creative Concept & Design: {money(costs['creative_cost'])}")
        st.write(f"Targeted Data Mining: {money(costs['data_mining_cost'])}")
        st.write(f"Email Development: {money(costs['email_labor_cost'])}")

        st.markdown("#### Marked-Up Costs")
        st.write(f"List Procurement: {money(costs['list_procurement_cost'])}")
        st.write(f"Variable Print Production: {money(costs['print_cost'])}")
        st.caption("List Procurement and Variable Print Production include a 35% markup.")

        st.markdown("#### Email Send Costs")
        st.write(f"Email Sends: {money(costs['email_send_cost'])}")

        st.markdown("#### Straight Costs")
        st.write(f"Selected straight cost items total: {money(costs['straight_cost_total'])}")

        st.markdown("#### Custom Costs")
        st.write(f"Custom costs total: {money(costs['custom_costs_total'])}")

        st.markdown("#### Pricing Logic")
        st.write(f"One-time cost total: {money(costs['one_time_cost_total'])}")
        st.write(f"Repeating cost total: {money(costs['repeating_cost_total'])}")
        st.write(f"Total estimated one-campaign cost: {money(costs['campaign_1_calc'])}")

        st.caption(
        "For 2+ campaign pricing, only Variable Print Production, Email Sends, "
        "List Procurement, and ACH Program Fee repeat. Four-campaign pricing applies a 10% discount."
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.session_state.campaign_1_cost_override = st.text_input(
            "One Campaign Cost",
            money(costs["campaign_1_calc"]),
        )

    with col2:
        st.session_state.campaign_2_cost_override = st.text_input(
            "Two Campaigns Total Cost",
            money(costs["campaign_2_calc"]),
        )
        st.session_state.campaign_2_per_cost_override = st.text_input(
            "Two Campaigns Per Campaign Cost",
            money(costs["campaign_2_per_calc"]),
        )

    with col3:
        st.session_state.campaign_4_cost_override = st.text_input(
            "Four Campaigns Total Cost",
            money(costs["campaign_4_calc"]),
        )
        st.session_state.campaign_4_per_cost_override = st.text_input(
            "Four Campaigns Per Campaign Cost",
            money(costs["campaign_4_per_calc"]),
        )

    section_complete_checkbox("Cost Estimator")  
    auto_save_proposal()


# ============================================================
# Generate Proposal
# ============================================================
elif section == "Generate Proposal":
    st.subheader("Generate Proposal")

    selected_template = TEMPLATE_MAP[st.session_state.proposal_type]
    selected_targets = get_selected_targets()
    selected_components = get_selected_components()
    total_targets = sum(count for count, _ in selected_targets)

    conversion_rate_decimal = parse_percent(st.session_state.conversion_rate)
    estimated_loans_refinanced = round(total_targets * conversion_rate_decimal)
    amount_refinanced = estimated_loans_refinanced * st.session_state.avg_auto_balance
    auto_interest_rate_decimal = parse_percent(st.session_state.auto_interest_rate)
    estimated_first_year_interest = calculate_first_year_interest(
        amount_refinanced,
        auto_interest_rate_decimal,
        st.session_state.auto_term_years,
    )

    costs = calculate_costs()

    campaign_1_cost = st.session_state.get(
        "campaign_1_cost_override",
        money(costs["campaign_1_calc"]),
    )
    campaign_2_cost = st.session_state.get(
        "campaign_2_cost_override",
        money(costs["campaign_2_calc"]),
    )
    campaign_2_per_cost = st.session_state.get(
        "campaign_2_per_cost_override",
        money(costs["campaign_2_per_calc"]),
    )
    campaign_4_cost = st.session_state.get(
        "campaign_4_cost_override",
        money(costs["campaign_4_calc"]),
    )
    campaign_4_per_cost = st.session_state.get(
        "campaign_4_per_cost_override",
        money(costs["campaign_4_per_calc"]),
    )

    # -----------------------------
    # Completion check
    # -----------------------------
    incomplete_sections = [
        sec for sec in REQUIRED_SECTIONS
        if not st.session_state.get(f"complete_{sec}", False)
    ]

    generate_clicked = False

    if incomplete_sections:
        st.warning(
            "Complete all sections before generating the proposal:\n\n- "
            + "\n- ".join(incomplete_sections)
        )

        st.markdown("### Go to Missing Section")

        for missing_section in incomplete_sections:
            if st.button(
                f"Go to {missing_section}",
                key=f"go_to_{missing_section}"
            ):
                st.session_state.active_section = missing_section
                st.rerun()

        st.button("Generate Proposal", disabled=True)

    else:
        st.success("All required sections are complete. You can generate the proposal.")

        generate_clicked = st.button(
            "Generate Proposal",
            key="generate_proposal_enabled",
            disabled=False
        )

    # -----------------------------
    # Summary Display
    # -----------------------------
    st.markdown("### Review Summary")
    st.write(f"Proposal Type: {st.session_state.proposal_type}")
    st.write(f"Credit Union: {st.session_state.credit_union}")
    st.write(f"Total Targets: {total_targets:,}")
    st.write(f"One Campaign Cost: {money(costs['campaign_1_calc'])}")

    # -----------------------------
    # Save Proposal
    # -----------------------------
    st.markdown("### Save Proposal")

    st.session_state.proposal_status = st.selectbox(
        "Proposal Status",
        ["Draft", "CU Review", "Signed", "Declined"],
        index=["Draft", "CU Review", "Signed", "Declined"].index(
            st.session_state.get("proposal_status", "Draft")
        )
    )

    col_save, col_save_complete = st.columns(2)

    with col_save:
        if st.button("Save Proposal"):
            saved_data = collect_saved_data()

            proposal_id = save_proposal(
               st.session_state.get("current_proposal_id"),
               st.session_state.proposal_name,
               st.session_state.credit_union,
               st.session_state.proposal_type,
               st.session_state.proposal_status,
               saved_data,
               st.session_state.msr,
               st.session_state.current_user
            )

            st.session_state.current_proposal_id = proposal_id
            st.success("Proposal saved.")

    with col_save_complete:
        if st.button("Save as Complete"):
            st.session_state.proposal_status = "Complete"
            saved_data = collect_saved_data()

            proposal_id = save_proposal(
               st.session_state.get("current_proposal_id"),
               st.session_state.proposal_name,
               st.session_state.credit_union,
               st.session_state.proposal_type,
               st.session_state.proposal_status,
               saved_data,
               st.session_state.msr,
               st.session_state.current_user
            )              

            st.session_state.current_proposal_id = proposal_id
            st.success("Proposal saved as complete.")

    # -----------------------------
    # Run Generation
    # -----------------------------
    if generate_clicked:
        if not selected_targets:
            st.error("Please select at least one campaign target.")
        elif not selected_components:
            st.error("Please select at least one campaign component.")
        elif not Path(selected_template).exists():
            st.error(f"Template file is missing: {selected_template}")
        else:
            gp.TEMPLATE_PATH = selected_template

            gp.proposal_data["{{proposal_name}}"] = st.session_state.proposal_name
            gp.proposal_data["{{proposal_date}}"] = format_date_windows(
                st.session_state.proposal_date
            )
            gp.proposal_data["{{creditunion_name}}"] = st.session_state.credit_union

            gp.proposal_data["{{target_conversion_rate}}"] = st.session_state.conversion_rate
            gp.proposal_data["{{total_targets}}"] = f"{total_targets:,}"

            gp.proposal_data["{{conversions}}"] = f"{estimated_loans_refinanced:,}"
            gp.proposal_data["{{avg_auto_balance}}"] = f"${st.session_state.avg_auto_balance:,.0f}"
            gp.proposal_data["{{amount_refinanced}}"] = f"${amount_refinanced:,.0f}"
            gp.proposal_data["{{first_year_interest}}"] = f"${estimated_first_year_interest:,.0f}"
            gp.proposal_data["{{auto_interest_rate}}"] = st.session_state.auto_interest_rate
            gp.proposal_data["{{auto_term_years}}"] = str(st.session_state.auto_term_years)

            gp.proposal_data["{{total_targets_2}}"] = f"{total_targets * 2:,}"
            gp.proposal_data["{{total_targets_4}}"] = f"{total_targets * 4:,}"

            gp.proposal_data["{{campaign_1_cost}}"] = campaign_1_cost
            gp.proposal_data["{{campaign_2_cost}}"] = campaign_2_cost
            gp.proposal_data["{{campaign_2_per_cost}}"] = campaign_2_per_cost
            gp.proposal_data["{{campaign_4_cost}}"] = campaign_4_cost
            gp.proposal_data["{{campaign_4_per_cost}}"] = campaign_4_per_cost

            gp.target_segments.clear()
            gp.target_segments.extend(selected_targets)

            gp.campaign_components.clear()
            gp.campaign_components.extend(selected_components)

            import os
            from datetime import datetime
            
            date_str = datetime.today().strftime("%Y%m%d")
            
            credit_union_clean = st.session_state.credit_union
            proposal_name_clean = st.session_state.proposal_name
            
            file_name = f"{date_str} {credit_union_clean} {proposal_name_clean}.docx"
            
            os.makedirs("generated_proposals", exist_ok=True)
            
            file_path = os.path.join("generated_proposals", file_name)
            
            gp.main(output_path=file_path)


            saved_data = collect_saved_data()
            saved_data["file_path"] = file_path
            st.session_state.file_path = file_path

            
            proposal_id = save_proposal(
                st.session_state.get("current_proposal_id"),
                st.session_state.proposal_name,
                st.session_state.credit_union,
                st.session_state.proposal_type,
                st.session_state.proposal_status,
                saved_data,
                st.session_state.msr,
                st.session_state.current_user
            )

            st.session_state.current_proposal_id = proposal_id

            st.success("Proposal generated successfully!")

            with open(file_path, "rb") as file:

                from datetime import datetime

                date_str = datetime.today().strftime("%Y%m%d")
                
                credit_union_clean = st.session_state.credit_union.replace(" ", " ")
                proposal_name_clean = st.session_state.proposal_name.replace(" ", " ")
                
                file_name = f"{date_str} {credit_union_clean} {proposal_name_clean}.docx"

                st.download_button(
                    label="Download Generated Proposal",
                    data=file,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
