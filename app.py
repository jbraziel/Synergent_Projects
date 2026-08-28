import streamlit as st
from datetime import date, datetime
from zoneinfo import ZoneInfo
from pathlib import Path
import generate_proposal as gp
from database import (
    initialize_database, save_proposal, search_proposals, load_proposal, delete_proposal,
    update_proposal_status, lock_proposal, unlock_proposal, get_pricing_snapshot,
    get_pricing_settings, update_pricing_setting, add_fixed_cost, get_pricing_history,
    save_pricing_snapshot, get_backend_status, get_default_pricing_snapshot,
    get_next_generation_version,
)
from config import get_admin_users
from file_storage import (
    store_file, store_bytes, read_bytes, stored_file_exists, display_name,
    copy_stored_file, make_object_path, get_storage_status, is_cloud_mode as cloud_file_mode,
)
import os
import json
import re
import shutil
import csv
from copy import deepcopy


st.set_page_config(page_title="Marketing Proposal Generator", page_icon="swirl.png", layout="wide")

try:
    initialize_database()
except RuntimeError as exc:
    st.error(str(exc))
    st.stop()



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

    /* Make the active section completion control easy to spot. */
    section[data-testid="stSidebar"] div[data-testid="stCheckbox"] {
       background-color: white !important;
       border: 1px solid #cfe3c3 !important;
       border-left: 4px solid #76bd22 !important;
       border-radius: 9px !important;
       padding: 0.55rem 0.65rem !important;
       margin: 0.15rem 0 0.55rem 0 !important;
    }

    section[data-testid="stSidebar"] div[data-testid="stCheckbox"] label {
       font-weight: 650 !important;
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

    /* Keep unavailable file placeholders aligned without looking active. */
    div[data-testid="stButton"] button:disabled {
        background-color: #e8ece6 !important;
        color: #9aa397 !important;
        opacity: 1 !important;
        cursor: default !important;
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

    /* Proposal Library action buttons: true full-width squares, matching the file buttons.
       Streamlit may put the key class on the stButton wrapper itself OR on an ancestor,
       so target both structures. */
    div[class*="st-key-open_"][data-testid="stButton"],
    div[class*="st-key-duplicate_"][data-testid="stButton"],
    div[class*="st-key-delete_"][data-testid="stButton"],
    div[class*="st-key-open_"] div[data-testid="stButton"],
    div[class*="st-key-duplicate_"] div[data-testid="stButton"],
    div[class*="st-key-delete_"] div[data-testid="stButton"] {
        width: 100% !important;
        min-width: 100% !important;
        display: block !important;
    }

    div[class*="st-key-open_"][data-testid="stButton"] button,
    div[class*="st-key-duplicate_"][data-testid="stButton"] button,
    div[class*="st-key-delete_"][data-testid="stButton"] button,
    div[class*="st-key-open_"] div[data-testid="stButton"] button,
    div[class*="st-key-duplicate_"] div[data-testid="stButton"] button,
    div[class*="st-key-delete_"] div[data-testid="stButton"] button {
        width: 100% !important;
        min-width: 100% !important;
        max-width: none !important;
        height: 40px !important;
        padding: 0 !important;
        margin: 0 !important;
        white-space: nowrap !important;
        font-size: 16px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
    }

    /* Top sidebar navigation: compact, consistent, app-like. */
    .sidebar-nav-label {
        color: #73806f;
        font-size: 10px;
        font-weight: 750;
        letter-spacing: 0.10em;
        margin: -0.05rem 0 0.30rem 0.10rem;
    }

    section[data-testid="stSidebar"] div[class*="st-key-nav_library"] button,
    section[data-testid="stSidebar"] div[class*="st-key-nav_new_proposal"] button,
    section[data-testid="stSidebar"] div[class*="st-key-nav_admin"] button {
        height: 36px !important;
        min-height: 36px !important;
        padding: 0 0.30rem !important;
        margin: 0 !important;
        border-radius: 9px !important;
        border: 1px solid #cbdcc2 !important;
        background: #ffffff !important;
        color: #3f4c3b !important;
        font-size: 11px !important;
        font-weight: 700 !important;
        line-height: 1 !important;
        box-shadow: 0 1px 2px rgba(47, 58, 47, 0.06) !important;
        white-space: nowrap !important;
    }

    section[data-testid="stSidebar"] div[class*="st-key-nav_library"] button:hover,
    section[data-testid="stSidebar"] div[class*="st-key-nav_admin"] button:hover {
        background: #f3f9ef !important;
        border-color: #76bd22 !important;
        color: #4d8518 !important;
        box-shadow: 0 2px 5px rgba(72, 105, 50, 0.10) !important;
    }

    /* Current navigation destination. */
    section[data-testid="stSidebar"] div[class*="st-key-nav_library"] button[kind="primary"],
    section[data-testid="stSidebar"] div[class*="st-key-nav_admin"] button[kind="primary"] {
        background: #76bd22 !important;
        border-color: #76bd22 !important;
        color: #ffffff !important;
        box-shadow: 0 2px 5px rgba(82, 132, 26, 0.18) !important;
    }

    section[data-testid="stSidebar"] div[class*="st-key-nav_library"] button[kind="primary"]:hover,
    section[data-testid="stSidebar"] div[class*="st-key-nav_admin"] button[kind="primary"]:hover {
        background: #65a91d !important;
        border-color: #65a91d !important;
        color: #ffffff !important;
    }

    /* New is an action, not a destination: keep it distinct but restrained. */
    section[data-testid="stSidebar"] div[class*="st-key-nav_new_proposal"] button {
        background: #edf7e6 !important;
        border-color: #a9cf8f !important;
        color: #4f851c !important;
    }

    section[data-testid="stSidebar"] div[class*="st-key-nav_new_proposal"] button:hover {
        background: #76bd22 !important;
        border-color: #76bd22 !important;
        color: #ffffff !important;
        box-shadow: 0 2px 5px rgba(82, 132, 26, 0.18) !important;
    }
</style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------
# Constants
# -----------------------------
TEMPLATE_MAP = {
    "Auto Loan Recapture Campaign": "ACH_Auto_Proposal_Template.pptx",
    "Synergent Email Platform Proposal": "EMP_Proposal_Template.pptx",
    "General Lending Campaign": "Lending_Proposal_Template.pptx",
    "Credit Card Campaign": "Credit_Card_Proposal_Template.pptx",
}

# People who can use the proposal tool. Legacy aliases keep older saved
# proposals compatible with the full-name display used in the UI.
USER_OPTIONS = [
    "Jen Braziel",
    "Shannan Heacock",
    "Erica Vachon",
    "Melanie Moore",
]

USER_NAME_ALIASES = {
    "Jen": "Jen Braziel",
    "Shannan": "Shannan Heacock",
    "Erica": "Erica Vachon",
    "Jen Braziel": "Jen Braziel",
    "Shannan Heacock": "Shannan Heacock",
    "Erica Vachon": "Erica Vachon",
    "Melanie Moore": "Melanie Moore",
}

MSR_OPTIONS = ["Shannan Heacock", "Erica Vachon"]
MSR_LEGACY_ALIASES = {
    "Shannan": "Shannan Heacock",
    "Erica": "Erica Vachon",
    "Shannan Heacock": "Shannan Heacock",
    "Erica Vachon": "Erica Vachon",
}

ADMIN_USERS = get_admin_users()

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

# Fixed costs are now database-backed and administered from the Admin area.
# Their labels/amounts are read from each proposal's frozen pricing schedule.

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

def build_credit_card_campaign_objectives():

    objectives = []

    if "Balance Transfer" in st.session_state.credit_card_goals:
        objectives.append(
            "BALANCE TRANSFER\n"
            "Increase credit card balances and strengthen member borrowing relationships by encouraging members to transfer higher-rate external credit card balances to the credit union’s credit card offering."
        )

    if "New Card Acquisition" in st.session_state.credit_card_goals:
        objectives.append(
            "NEW CARD ACQUISITION\n"
            "Increase credit card adoption among targeted member segments through targeted outreach and promotion of the credit union’s credit card program."
        )

    if "Card Utilization & Activation" in st.session_state.credit_card_goals:
        objectives.append(
            "CARD UTILIZATION & ACTIVATION\n"
            "Increase active credit card usage and strengthen top-of-wallet positioning by encouraging existing cardholders to engage more consistently with their credit union credit card."
        )

    return "\n\n".join(objectives)

def get_selected_credit_card_targets():
    selected_targets = []

    credit_card_target_groups = {
        "cc_bt": [
            "Members with recurring ACH payments to external credit card providers"
        ],
        "cc_new": [
            "Members with checking and no credit card with the credit union"
        ],
        "cc_util": [
            "Existing credit card holders with low or inactive card usage"
        ],
    }

    for key_prefix, default_targets in credit_card_target_groups.items():

        for i, description in enumerate(default_targets):
            include = st.session_state.get(
                f"{key_prefix}_target_include_saved_{i}",
                False
            )

            count = st.session_state.get(
                f"{key_prefix}_target_count_saved_{i}",
                0
            )

            if include and count > 0:
                selected_targets.append(
                    (int(count), description)
                )

        custom_key = f"{key_prefix}_custom_targets"

        for target in st.session_state.get(custom_key, []):

            if isinstance(target, str):
                target = {
                    "count": 0,
                    "description": target
                }

            count = target.get("count", 0)
            description = target.get("description", "")

            if count > 0 and description.strip():
                selected_targets.append(
                    (int(count), description.strip())
                )

    return selected_targets

def credit_card_target_section(section_label, key_prefix, default_targets):
    st.markdown(f"#### {section_label}")

    for i, target in enumerate(default_targets):
        include_saved_key = f"{key_prefix}_target_include_saved_{i}"
        count_saved_key = f"{key_prefix}_target_count_saved_{i}"

        if include_saved_key not in st.session_state:
            st.session_state[include_saved_key] = False

        if count_saved_key not in st.session_state:
            st.session_state[count_saved_key] = 0

        col_check, col_count, col_desc = st.columns([0.08, 0.18, 0.74])

        with col_check:
            st.session_state[include_saved_key] = st.checkbox(
                "",
                value=st.session_state[include_saved_key],
                key=f"widget_{include_saved_key}"
            )

        with col_count:
            st.session_state[count_saved_key] = st.number_input(
                "Count",
                min_value=0,
                value=int(st.session_state[count_saved_key]),
                step=1,
                key=f"widget_{count_saved_key}",
                label_visibility="collapsed"
            )

        with col_desc:
            st.text(target)

    custom_key = f"{key_prefix}_custom_targets"

    if custom_key not in st.session_state:
        st.session_state[custom_key] = []

    if st.button(
        f"Add Custom {section_label} Target",
        key=f"add_{key_prefix}_target"
    ):
        st.session_state[custom_key].append(
            {"count": 0, "description": ""}
        )
        st.rerun()

    st.session_state[custom_key] = [
        {"count": 0, "description": target}
        if isinstance(target, str)
        else target
        for target in st.session_state[custom_key]
    ]

    updated_targets = []

    for i, target in enumerate(st.session_state[custom_key]):
        col_count, col_desc = st.columns([0.25, 0.75])

        with col_count:
            count = st.number_input(
                f"Custom {section_label} Target {i + 1} Count",
                min_value=0,
                value=int(target.get("count", 0)),
                step=1,
                key=f"{key_prefix}_custom_target_count_{i}",
            )

        with col_desc:
            description = st.text_input(
                f"Custom {section_label} Target {i + 1}",
                value=target.get("description", ""),
                key=f"{key_prefix}_custom_target_desc_{i}",
            )

        updated_targets.append(
            {"count": int(count), "description": description}
        )

    st.session_state[custom_key] = updated_targets

def convert_pptx_to_pdf(pptx_path):
    import os
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()

    powerpoint = None
    presentation = None

    try:
        pdf_path = pptx_path.replace(".pptx", ".pdf")

        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1

        presentation = powerpoint.Presentations.Open(
            os.path.abspath(pptx_path),
            WithWindow=False
        )

        presentation.SaveAs(
            os.path.abspath(pdf_path),
            32
        )

        return pdf_path

    finally:
        if presentation is not None:
            presentation.Close()

        if powerpoint is not None:
            powerpoint.Quit()

        pythoncom.CoUninitialize()

ADD_NEW_CU_OPTION = "➕ Add New Credit Union"
def load_credit_union_list():
    try:
        with open("CU List.csv", "r", encoding="utf-8-sig") as file:
            names = [line.strip() for line in file if line.strip() and line.strip() != ADD_NEW_CU_OPTION]
        return sorted(set(names))
    except FileNotFoundError:
        return []

def create_pricing_export_csv(pricing_folder):
    """Create an auditable pricing detail export for the current proposal.

    The export intentionally stores both calculated pricing and any manual proposal-price
    overrides so the internal pricing record always reconciles to the PowerPoint.
    """
    proposal_id = st.session_state.get("current_proposal_id") or "NEW"
    version = extract_generation_version(st.session_state.get("file_path"))
    version_suffix = f"_v{version}" if version else ""
    file_name = f"P{proposal_id}_Pricing{version_suffix}.csv"
    export_path = os.path.join(pricing_folder, file_name)

    rows = [
        ["Proposal ID", st.session_state.get("current_proposal_id") or ""],
        ["Proposal Name", st.session_state.proposal_name],
        ["Credit Union", st.session_state.credit_union],
        ["Proposal Type", st.session_state.proposal_type],
        ["Proposal Date", str(st.session_state.proposal_date)],
        ["MSR", st.session_state.msr],
        ["Prepared By", st.session_state.get("current_user", "")],
        ["Copied From Proposal ID", st.session_state.get("copied_from_proposal_id") or ""],
        [],
        ["PROPOSAL NOTES", ""],
        ["Notes", st.session_state.get("proposal_notes", "").strip() or "(none)"],
        [],
    ]

    if st.session_state.proposal_type == "Synergent Email Platform Proposal":
        emp = emp_pricing_details(st.session_state.total_subscribers)
        tier_cost = emp.get("tier_cost")
        rows.extend([
            ["EMP PRICING", ""],
            ["Total Subscribers", st.session_state.total_subscribers],
            ["Tier", emp.get("tier_name")],
            ["Base Monthly Cost", tier_cost if tier_cost is not None else "Custom"],
            ["Essentials Monthly Cost", emp.get("essentials_monthly", "Custom")],
            ["Premium Monthly Cost", emp.get("premium_monthly", "Custom")],
            ["Elite Monthly Cost", emp.get("elite_monthly", "Custom")],
            ["Essentials Implementation", emp.get("essentials_implementation", pricing_value("emp_essentials_implementation", 5500))],
            ["Premium Implementation", emp.get("premium_implementation", pricing_value("emp_premium_implementation", 8000))],
            ["Elite Implementation", emp.get("elite_implementation", pricing_value("emp_elite_implementation", 10500))],
        ])
    else:
        if st.session_state.proposal_type == "Credit Card Campaign":
            selected_targets = get_selected_credit_card_targets()
        else:
            selected_targets = get_selected_targets()

        selected_components = get_selected_components()
        total_targets = sum(count for count, _ in selected_targets)
        costs = calculate_costs()

        rows.extend([
            ["CAMPAIGN INPUTS", ""],
            ["Total Targets", total_targets],
            ["Target Conversion Rate", st.session_state.conversion_rate],
            ["Loan / Product Type", st.session_state.get("loan_type", "")],
            ["Campaign Weeks", st.session_state.get("campaign_weeks", "")],
        ])

        if st.session_state.proposal_type == "Credit Card Campaign":
            rows.extend([
                ["Credit Card Goals", "; ".join(st.session_state.get("credit_card_goals", []))],
                ["Average Credit Card Limit", st.session_state.get("avg_credit_card_limit", 0)],
                ["Average Credit Card Rate", st.session_state.get("avg_credit_card_rate", "")],
                ["Average Interchange Per Card", st.session_state.get("avg_interchange_per_card", 0)],
            ])
        else:
            conversion_rate_decimal = parse_percent(st.session_state.conversion_rate)
            estimated_loans_refinanced = round(total_targets * conversion_rate_decimal)
            amount_refinanced = estimated_loans_refinanced * st.session_state.avg_loan_balance
            loan_interest_rate_decimal = parse_percent(st.session_state.loan_interest_rate)
            estimated_first_year_interest = calculate_first_year_interest(
                amount_refinanced,
                loan_interest_rate_decimal,
                st.session_state.loan_term_years,
            )
            rows.extend([
                ["Average Loan Balance", st.session_state.avg_loan_balance],
                ["Average Interest Rate", st.session_state.loan_interest_rate],
                ["Average Term Years", st.session_state.loan_term_years],
                ["Estimated Conversions", estimated_loans_refinanced],
                ["Estimated Amount Financed", amount_refinanced],
                ["Estimated First-Year Interest", estimated_first_year_interest],
            ])

        rows.append([])
        rows.append(["TARGET SEGMENTS", "Count"])
        for count, description in selected_targets:
            rows.append([description, count])

        rows.append([])
        rows.append(["CAMPAIGN COMPONENTS", "Included"])
        for component in selected_components:
            rows.append([component, "Yes"])

        rows.extend([
            [],
            ["PRICING INPUT DETAIL", "Included", "Quantity / Raw Cost", "Rate / Markup", "Calculated Cost"],
            ["Creative Concept & Design", "Yes" if st.session_state.include_creative else "No", st.session_state.creative_hours, f"${pricing_value('creative_hourly_rate', 115):,.2f}/hour", costs["creative_cost"]],
            ["Targeted Data Mining", "Yes" if st.session_state.include_data_mining else "No", st.session_state.data_mining_hours, f"${pricing_value('programming_hourly_rate', 200):,.2f}/hour", costs["data_mining_cost"]],
            ["List Procurement", "Yes" if st.session_state.include_list_procurement else "No", st.session_state.list_procurement_raw, f"{pricing_value('list_markup_pct', 35):,.1f}% markup", costs["list_procurement_cost"]],
            ["Variable Print Production", "Yes" if st.session_state.include_print else "No", st.session_state.print_raw, f"{pricing_value('print_markup_pct', 35):,.1f}% markup", costs["print_cost"]],
            ["Email Development", "Yes" if st.session_state.include_email_labor else "No", st.session_state.email_hours, f"${pricing_value('email_development_hourly_rate', 115):,.2f}/hour", costs["email_labor_cost"]],
            ["Email Sends", "Yes" if st.session_state.include_email_sends else "No", st.session_state.email_send_count, f"${pricing_value('email_send_fee', 100):,.2f}/send", costs["email_send_cost"]],
            [],
            ["FIXED COSTS", "Amount", "Included", "Behavior"],
        ])

        for item in get_fixed_cost_records():
            name = item["label"]
            amount = item["value"]
            repeating_label = "Repeats per campaign" if item.get("is_repeating") else "One-time"
            rows.append([name, amount, "Yes" if st.session_state.get(f"cost_{name}", True) else "No", repeating_label])

        rows.extend([
            ["Fixed Cost Total", costs["straight_cost_total"], ""],
            [],
            ["CUSTOM COSTS", "Amount"],
        ])
        valid_custom_costs = [
            item for item in st.session_state.get("custom_costs", [])
            if item.get("name", "").strip() and float(item.get("amount", 0) or 0) > 0
        ]
        if valid_custom_costs:
            for item in valid_custom_costs:
                rows.append([item.get("name", "").strip(), item.get("amount", 0)])
        else:
            rows.append(["(none)", 0])

        rows.extend([
            ["Custom Costs Total", costs["custom_costs_total"]],
            [],
            ["PRICING SUMMARY", "Calculated", "Final Proposal Price"],
            ["One-Time Cost Total", costs["one_time_cost_total"], ""],
            ["Repeating Cost Total", costs["repeating_cost_total"], ""],
            ["1 Campaign", costs["campaign_1_calc"], st.session_state.get("campaign_1_cost_override", money(costs["campaign_1_calc"]))],
            ["2 Campaigns Total", costs["campaign_2_calc"], st.session_state.get("campaign_2_cost_override", money(costs["campaign_2_calc"]))],
            ["2 Campaigns Per Campaign", costs["campaign_2_per_calc"], st.session_state.get("campaign_2_per_cost_override", money(costs["campaign_2_per_calc"]))],
            ["4 Campaigns Total", costs["campaign_4_calc"], st.session_state.get("campaign_4_cost_override", money(costs["campaign_4_calc"]))],
            ["4 Campaigns Per Campaign", costs["campaign_4_per_calc"], st.session_state.get("campaign_4_per_cost_override", money(costs["campaign_4_per_calc"]))],
        ])

    with open(export_path, "w", newline="", encoding="utf-8-sig") as file:
        csv.writer(file).writerows(rows)

    return export_path

def clean_folder_name(name):
    name = re.sub(r'[<>:"/\\|?*]', "", name)
    return name.strip()


PROPOSAL_TYPE_FILE_LABELS = {
    "Auto Loan Recapture Campaign": "AutoLoan",
    "General Lending Campaign": "GeneralLending",
    "Credit Card Campaign": "CreditCard",
    "Synergent Email Platform Proposal": "EMP",
}


def filename_token(value, max_length=24):
    """Create a short, filesystem-safe identifier for human-readable filenames."""
    value = str(value or "").strip()
    value = re.sub(r"(?i)\bFederal Credit Union\b", "", value)
    value = re.sub(r"(?i)\bCredit Union\b", "", value)
    value = re.sub(r"(?i)\bFederal CU\b", "", value)
    value = value.replace("&", "")
    token = re.sub(r"[^A-Za-z0-9]+", "", value)
    return (token or "CU")[:max_length]


def proposal_file_stem(proposal_id=None):
    proposal_id = proposal_id or st.session_state.get("current_proposal_id") or "NEW"
    cu = filename_token(st.session_state.get("credit_union", "CU"))
    proposal_type = st.session_state.get("proposal_type", "Proposal")
    type_token = PROPOSAL_TYPE_FILE_LABELS.get(proposal_type, filename_token(proposal_type, 20))
    return f"P{proposal_id}_{cu}_{type_token}"


def extract_generation_version(file_ref):
    if not file_ref:
        return None
    match = re.search(r"(?:_|\b)v(\d+)(?:\.[^.]+)$", display_name(file_ref), re.IGNORECASE)
    return int(match.group(1)) if match else None


def lifecycle_filename(status, extension="pptx", source_ref=None):
    version = extract_generation_version(source_ref or st.session_state.get("file_path"))
    version_suffix = f"_v{version}" if version else ""
    return f"{proposal_file_stem()}_{status.upper()}{version_suffix}.{extension}"


def signed_pdf_filename(source_ref=None, existing_signed_ref=None):
    """Return a short signed-PDF name, preserving replacement revisions instead of overwriting."""
    version = extract_generation_version(source_ref or st.session_state.get("sent_file_path") or st.session_state.get("file_path"))
    version_suffix = f"_v{version}" if version else ""
    base = f"{proposal_file_stem()}_SIGNED{version_suffix}"

    if not existing_signed_ref:
        return f"{base}.pdf"

    existing_name = display_name(existing_signed_ref)
    revision_match = re.search(r"_r(\d+)\.pdf$", existing_name, re.IGNORECASE)
    next_revision = int(revision_match.group(1)) + 1 if revision_match else 2
    return f"{base}_r{next_revision}.pdf"


def get_credit_union_output_folder(credit_union):
    cu_folder = clean_folder_name(credit_union)
    
    SHARED_ROOT = os.environ.get("PROPOSAL_OUTPUT_ROOT", "generated_proposals")
    base_folder = os.path.join(SHARED_ROOT, cu_folder)

    drafts_folder = os.path.join(base_folder, "Drafts")
    sent_folder = os.path.join(base_folder, "Sent")
    signed_folder = os.path.join(base_folder, "Signed")
    pricing_folder = os.path.join(base_folder, "Pricing Exports")

    os.makedirs(drafts_folder, exist_ok=True)
    os.makedirs(sent_folder, exist_ok=True)
    os.makedirs(signed_folder, exist_ok=True)
    os.makedirs(pricing_folder, exist_ok=True)

    return base_folder, drafts_folder, sent_folder, signed_folder, pricing_folder

def get_active_pricing_schedule():
    """Return the proposal's frozen pricing schedule, creating one for new proposals."""
    snapshot = st.session_state.get("pricing_settings_snapshot")
    if not snapshot:
        snapshot = get_pricing_snapshot()
        st.session_state.pricing_settings_snapshot = snapshot
    return snapshot


def pricing_value(key, default=0.0):
    item = get_active_pricing_schedule().get(key, {})
    try:
        return float(item.get("value", default))
    except (TypeError, ValueError):
        return float(default)


def get_fixed_cost_records():
    records = []
    for key, item in get_active_pricing_schedule().items():
        if item.get("category") == "Fixed Costs" and item.get("active", True):
            row = dict(item)
            row["key"] = key
            row["value"] = float(row.get("value", 0) or 0)
            row["is_repeating"] = bool(row.get("is_repeating", False))
            records.append(row)
    return sorted(records, key=lambda x: (x.get("sort_order", 0), x.get("label", "")))


def get_fixed_cost_items():
    return {item["label"]: item["value"] for item in get_fixed_cost_records()}


def emp_pricing_details(total_subscribers):
    tier_cost, tier_name = calculate_emp_tier_cost(total_subscribers)
    if tier_cost is None:
        return {"tier_cost": None, "tier_name": tier_name}
    return {
        "tier_cost": tier_cost,
        "tier_name": tier_name,
        "essentials_monthly": tier_cost + pricing_value("emp_essentials_addon", 100),
        "premium_monthly": tier_cost + pricing_value("emp_premium_addon", 200),
        "elite_monthly": tier_cost + pricing_value("emp_elite_addon", 200),
        "essentials_implementation": pricing_value("emp_essentials_implementation", 5500),
        "premium_implementation": pricing_value("emp_premium_implementation", 8000),
        "elite_implementation": pricing_value("emp_elite_implementation", 10500),
    }


def build_pricing_audit_snapshot():
    """Capture the exact rate schedule, inputs and calculated/final prices used at generation."""
    snapshot = {
        "pricing_schedule": deepcopy(get_active_pricing_schedule()),
        "proposal_type": st.session_state.proposal_type,
        "proposal_name": st.session_state.proposal_name,
        "credit_union": st.session_state.credit_union,
        "proposal_date": str(st.session_state.proposal_date),
    }
    if st.session_state.proposal_type == "Synergent Email Platform Proposal":
        snapshot["emp"] = {
            "total_subscribers": st.session_state.total_subscribers,
            **emp_pricing_details(st.session_state.total_subscribers),
        }
    else:
        costs = calculate_costs()
        snapshot["inputs"] = {
            "creative_hours": st.session_state.creative_hours,
            "data_mining_hours": st.session_state.data_mining_hours,
            "list_procurement_raw": st.session_state.list_procurement_raw,
            "print_raw": st.session_state.print_raw,
            "email_hours": st.session_state.email_hours,
            "email_send_count": st.session_state.email_send_count,
            "custom_costs": deepcopy(st.session_state.get("custom_costs", [])),
        }
        snapshot["calculated_costs"] = costs
        snapshot["final_prices"] = {
            "campaign_1": st.session_state.get("campaign_1_cost_override", money(costs["campaign_1_calc"])),
            "campaign_2_total": st.session_state.get("campaign_2_cost_override", money(costs["campaign_2_calc"])),
            "campaign_2_per": st.session_state.get("campaign_2_per_cost_override", money(costs["campaign_2_per_calc"])),
            "campaign_4_total": st.session_state.get("campaign_4_cost_override", money(costs["campaign_4_calc"])),
            "campaign_4_per": st.session_state.get("campaign_4_per_cost_override", money(costs["campaign_4_per_calc"])),
        }
    return snapshot


def get_required_sections():
    if st.session_state.proposal_type == "Synergent Email Platform Proposal":
        return [
            "Proposal Details",
            "EMP Details",
        ]

    return [
        "Proposal Details",
        "Campaign Targets",
        "Conversion Metrics",
        "Campaign Components",
        "Cost Estimator",
    ]

def calculate_emp_tier_cost(total_subscribers):
    if 2500 <= total_subscribers <= 4999:
        return pricing_value("emp_tier_1_base", 59.54), "Tier 1"
    elif 5000 <= total_subscribers <= 9999:
        return pricing_value("emp_tier_2_base", 108.14), "Tier 2"
    elif 10000 <= total_subscribers <= 14999:
        return pricing_value("emp_tier_3_base", 156.74), "Tier 3"
    elif 15000 <= total_subscribers <= 24999:
        return pricing_value("emp_tier_4_base", 241.79), "Tier 4"
    elif 25000 <= total_subscribers <= 49999:
        return pricing_value("emp_tier_5_base", 363.29), "Tier 5"
    elif 50000 <= total_subscribers <= 74999:
        return pricing_value("emp_tier_6_base", 545.54), "Tier 6"
    else:
        return None, "Custom"

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
        "file_path",
        "sent_file_path",
        "sent_at",
        "signed_file_path",
        "signed_at",
        "signed_uploaded_by",
        "signed_original_name",
        "pricing_export_path",
        "proposal_notes",
        "loan_type",
        "campaign_weeks",
        "loan_interest_rate",
        "avg_loan_balance",
        "loan_term_years",
        "sent_pdf_path",
        "credit_card_goals",
        "cc_bt_custom_targets",
        "cc_new_custom_targets",
        "cc_util_custom_targets",
        "avg_credit_card_limit",
        "avg_credit_card_rate",
        "avg_interchange_per_card",
        "total_subscribers",
        "campaign_1_cost_override",
        "campaign_2_cost_override",
        "campaign_2_per_cost_override",
        "campaign_4_cost_override",
        "campaign_4_per_cost_override",
        "copied_from_proposal_id",
        "pricing_settings_snapshot",
    ]

    for key in keys_to_save:
        value = st.session_state.get(key)

        if key == "proposal_date" and value is not None:
            value = value.isoformat()

        data[key] = value

    for i, _ in enumerate(DEFAULT_TARGET_OPTIONS, start=1):
        data[f"target_include_saved_{i}"] = st.session_state.get(f"target_include_saved_{i}", i <= 5)
        data[f"target_count_saved_{i}"] = st.session_state.get(f"target_count_saved_{i}", DEFAULT_TARGET_OPTIONS[i - 1][0])

    credit_card_target_prefixes = {
       "cc_bt": 1,
       "cc_new": 1,
       "cc_util": 1,
    }

    for prefix, target_count in credit_card_target_prefixes.items():
        for i in range(target_count):
            data[f"{prefix}_target_include_saved_{i}"] = st.session_state.get(
                f"{prefix}_target_include_saved_{i}",
                False
            )
            data[f"{prefix}_target_count_saved_{i}"] = st.session_state.get(
                f"{prefix}_target_count_saved_{i}",
                0
            )

    for i, _ in enumerate(DEFAULT_COMPONENTS, start=1):
        data[f"component_saved_{i}"] = st.session_state.get(f"component_saved_{i}", True)

    for item in get_fixed_cost_records():
        name = item["label"]
        data[f"cost_{name}"] = st.session_state.get(f"cost_{name}", True)

    for sec in [
    "Proposal Details",
    "Campaign Targets",
    "Conversion Metrics",
    "Campaign Components",
    "Cost Estimator",
    "EMP Details",
    ]:
        data[f"complete_{sec}"] = st.session_state.get(f"complete_{sec}", False)

    return data


def clear_completion_widget_state():
    """Clear widget mirrors so a newly loaded proposal displays its own completion state."""
    for key in list(st.session_state.keys()):
        if key.startswith("sidebar_complete_") or key.startswith("widget_complete_"):
            del st.session_state[key]


def canonical_user_name(name):
    return USER_NAME_ALIASES.get(name, name or "")


def canonical_msr_name(name):
    return MSR_LEGACY_ALIASES.get(name, name or "")


def load_saved_data(data):
    clear_completion_widget_state()
    # Proposals saved before the Admin-pricing upgrade did not contain a rate schedule.
    # Freeze them to the legacy rates rather than silently applying a future Admin change.
    if not data.get("pricing_settings_snapshot"):
        st.session_state.pricing_settings_snapshot = get_default_pricing_snapshot()
    for key, value in data.items():
        if key == "proposal_date" and value:
            value = date.fromisoformat(value)
        elif key in {"current_user", "updated_by", "locked_by"}:
            value = canonical_user_name(value)
        elif key == "msr":
            value = canonical_msr_name(value)
        st.session_state[key] = value


def get_workflow_sections():
    return get_required_sections() + ["Generate Proposal"]


def go_to_next_section(section_name):
    sections = get_workflow_sections()
    if section_name in sections:
        current_index = sections.index(section_name)
        if current_index + 1 < len(sections):
            st.session_state.active_section = sections[current_index + 1]


def section_complete_checkbox(section_name):
    """Section completion is managed in the persistent sidebar progress navigator."""
    # Kept as a no-op so existing section calls remain simple and the page itself stays uncluttered.
    return


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
            and st.session_state.avg_loan_balance > 0
            and parse_percent(st.session_state.loan_interest_rate) > 0
            and st.session_state.loan_term_years > 0
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


def format_display_datetime(value):
    """Display stored timestamps as compact Eastern Time without changing DB storage."""
    if not value:
        return ""

    try:
        if isinstance(value, datetime):
            dt = value
        else:
            text = str(value).strip()
            # Supabase commonly returns ISO timestamps ending in Z or +00:00.
            dt = datetime.fromisoformat(text.replace("Z", "+00:00"))

        eastern = ZoneInfo("America/New_York")
        if dt.tzinfo is None:
            # Legacy/local SQLite timestamps were written as local wall-clock time.
            dt = dt.replace(tzinfo=eastern)
        else:
            dt = dt.astimezone(eastern)

        time_text = dt.strftime("%I:%M %p").lstrip("0")
        return f"{dt.month}/{dt.day}/{dt.year} {time_text}"
    except (TypeError, ValueError):
        # Never let an unexpected historical timestamp break the Proposal Library.
        return str(value)


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
        "current_user": "",
        "total_subscribers": 15000,
        "proposal_notes": "",
        "loan_type": "Auto Loan",
        "campaign_weeks": 8,
        "loan_interest_rate": "6.00%",
        "avg_loan_balance": 18500,
        "loan_term_years": 5,
        "sent_pdf_path": "",
        "credit_card_goals": [],
        "cc_bt_custom_targets": [],
        "cc_new_custom_targets": [],
        "cc_util_custom_targets": [],
        "avg_credit_card_limit": 5000,
        "avg_credit_card_rate": "14.99%",
        "avg_interchange_per_card": 75,
        "copied_from_proposal_id": None,
        "file_path": "",
        "sent_file_path": "",
        "sent_at": "",
        "signed_file_path": "",
        "signed_at": "",
        "signed_uploaded_by": "",
        "signed_original_name": "",
        "pricing_export_path": "",
        "pricing_settings_snapshot": None,
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

    if not st.session_state.get("pricing_settings_snapshot"):
        st.session_state.pricing_settings_snapshot = get_pricing_snapshot()

    for i, (_, _) in enumerate(DEFAULT_TARGET_OPTIONS, start=1):
        if f"target_include_{i}" not in st.session_state:
            st.session_state[f"target_include_{i}"] = i <= 5
        if f"target_count_{i}" not in st.session_state:
            st.session_state[f"target_count_{i}"] = DEFAULT_TARGET_OPTIONS[i - 1][0]

    for i, _ in enumerate(DEFAULT_COMPONENTS, start=1):
        if f"component_{i}" not in st.session_state:
            st.session_state[f"component_{i}"] = True

    for item in get_fixed_cost_records():
        name = item["label"]
        if f"cost_{name}" not in st.session_state:
            st.session_state[f"cost_{name}"] = True


def reset_proposal_state():
    """Start with a genuinely clean proposal while preserving the selected user."""
    current_user = st.session_state.get("current_user", "")
    for key in list(st.session_state.keys()):
        if key != "current_user":
            del st.session_state[key]
    init_state()
    st.session_state.current_user = current_user
    st.session_state.active_section = "Proposal Details"


def duplicate_proposal(source_proposal_id):
    """Create a new editable draft using another proposal as the starting point."""
    source_data = load_proposal(source_proposal_id)
    if not source_data:
        return None

    copied = deepcopy(source_data)
    original_name = copied.get("proposal_name", "Proposal")

    # Reset lifecycle/file fields. The new proposal shares inputs, not document history.
    copied.update({
        "current_proposal_id": None,
        "proposal_status": "Draft",
        "proposal_name": f"{original_name} - Copy",
        "proposal_date": date.today().isoformat(),
        "file_path": "",
        "sent_file_path": "",
        "sent_at": "",
        "sent_pdf_path": "",
        "signed_file_path": "",
        "signed_at": "",
        "signed_uploaded_by": "",
        "signed_original_name": "",
        "pricing_export_path": "",
        "copied_from_proposal_id": source_proposal_id,
        # A duplicate is a new proposal, so it starts on the current admin rate schedule.
        "pricing_settings_snapshot": get_pricing_snapshot(),
    })

    for sec in ["Proposal Details", "Campaign Targets", "Conversion Metrics", "Campaign Components", "Cost Estimator", "EMP Details"]:
        copied[f"complete_{sec}"] = False

    reset_proposal_state()
    load_saved_data(copied)
    st.session_state.current_proposal_id = None
    st.session_state.proposal_status = "Draft"
    st.session_state.copied_from_proposal_id = source_proposal_id

    new_id = save_proposal(
        None,
        st.session_state.proposal_name,
        st.session_state.credit_union,
        st.session_state.proposal_type,
        "Draft",
        collect_saved_data(),
        st.session_state.msr,
        st.session_state.current_user,
        copied_from_proposal_id=source_proposal_id,
    )
    st.session_state.current_proposal_id = new_id
    # Save once more so the JSON snapshot also contains its new ID.
    save_proposal(
        new_id,
        st.session_state.proposal_name,
        st.session_state.credit_union,
        st.session_state.proposal_type,
        "Draft",
        collect_saved_data(),
        st.session_state.msr,
        st.session_state.current_user,
        copied_from_proposal_id=source_proposal_id,
    )
    lock_proposal(new_id, st.session_state.current_user)
    st.session_state.active_section = "Proposal Details"
    return new_id


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
    creative_rate = pricing_value("creative_hourly_rate", 115)
    programming_rate = pricing_value("programming_hourly_rate", 200)
    email_rate = pricing_value("email_development_hourly_rate", 115)
    list_markup_rate = 1 + (pricing_value("list_markup_pct", 35) / 100)
    print_markup_rate = 1 + (pricing_value("print_markup_pct", 35) / 100)
    email_send_rate = pricing_value("email_send_fee", 100)
    four_campaign_discount = max(0.0, min(100.0, pricing_value("four_campaign_discount_pct", 10))) / 100

    creative_cost = (
        st.session_state.creative_hours * creative_rate
        if st.session_state.include_creative else 0
    )
    data_mining_cost = (
        st.session_state.data_mining_hours * programming_rate
        if st.session_state.include_data_mining else 0
    )
    list_procurement_cost = (
        st.session_state.list_procurement_raw * list_markup_rate
        if st.session_state.include_list_procurement else 0
    )
    print_cost = (
        st.session_state.print_raw * print_markup_rate
        if st.session_state.include_print else 0
    )
    email_labor_cost = (
        st.session_state.email_hours * email_rate
        if st.session_state.include_email_labor else 0
    )
    email_send_cost = (
        st.session_state.email_send_count * email_send_rate
        if st.session_state.include_email_sends else 0
    )

    straight_cost_total = 0
    repeating_fixed_cost_total = 0
    selected_straight_costs = []
    for item in get_fixed_cost_records():
        name = item["label"]
        cost = float(item["value"])
        if st.session_state.get(f"cost_{name}", True):
            straight_cost_total += cost
            selected_straight_costs.append((name, cost))
            if item.get("is_repeating", False):
                repeating_fixed_cost_total += cost

    custom_costs_total = 0
    for item in st.session_state.get("custom_costs", []):
        if item.get("name", "").strip() and item.get("amount", 0) > 0:
            custom_costs_total += item["amount"]

    estimated_cost_total = (
        creative_cost + data_mining_cost + list_procurement_cost + print_cost +
        email_labor_cost + email_send_cost + straight_cost_total + custom_costs_total
    )

    repeating_cost_total = (
        print_cost + list_procurement_cost + email_send_cost + repeating_fixed_cost_total
    )
    one_time_cost_total = estimated_cost_total - repeating_cost_total

    campaign_1_calc = estimated_cost_total
    campaign_2_calc = one_time_cost_total + (repeating_cost_total * 2)
    campaign_2_per_calc = campaign_2_calc / 2
    campaign_4_pre_discount = one_time_cost_total + (repeating_cost_total * 4)
    campaign_4_calc = campaign_4_pre_discount * (1 - four_campaign_discount)
    campaign_4_per_calc = campaign_4_calc / 4

    return {
        "creative_cost": creative_cost,
        "data_mining_cost": data_mining_cost,
        "list_procurement_cost": list_procurement_cost,
        "print_cost": print_cost,
        "email_labor_cost": email_labor_cost,
        "email_send_cost": email_send_cost,
        "straight_cost_total": straight_cost_total,
        "repeating_fixed_cost_total": repeating_fixed_cost_total,
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
    
    # Normalize legacy first-name values from existing browser sessions.
    st.session_state.current_user = canonical_user_name(st.session_state.current_user)
    user_options = [""] + USER_OPTIONS

    user = st.selectbox(
        "Who is using the proposal tool?",
        user_options,
        index=user_options.index(st.session_state.current_user)
        if st.session_state.current_user in user_options
        else 0
    )

    st.session_state.current_user = user
    
    # Stop app until user is selected
    if not st.session_state.current_user:
        st.warning("Please select your name to continue")
        st.stop()

    st.markdown("---")

    is_admin_user = st.session_state.current_user in ADMIN_USERS
    st.markdown('<div class="sidebar-nav-label">NAVIGATION</div>', unsafe_allow_html=True)
    nav_cols = st.columns(3 if is_admin_user else 2, gap="small")
    with nav_cols[0]:
        if st.button(
            "Library",
            key="nav_library",
            use_container_width=True,
            type="primary" if st.session_state.active_section == "Proposal Library" else "secondary",
            help="Open the Proposal Library",
        ):
            st.session_state.active_section = "Proposal Library"
            st.rerun()
    with nav_cols[1]:
        if st.button(
            "+ New",
            key="nav_new_proposal",
            use_container_width=True,
            help="Start a new proposal",
        ):
            if st.session_state.get("current_proposal_id"):
                unlock_proposal(st.session_state.current_proposal_id, st.session_state.current_user)
            reset_proposal_state()
            st.rerun()
    if is_admin_user:
        with nav_cols[2]:
            if st.button(
                "Admin",
                key="nav_admin",
                use_container_width=True,
                type="primary" if st.session_state.active_section == "Admin" else "secondary",
                help="Open pricing and configuration administration",
            ):
                if st.session_state.get("current_proposal_id"):
                    auto_save_proposal()
                    unlock_proposal(st.session_state.current_proposal_id, st.session_state.current_user)
                    st.session_state.current_proposal_id = None
                st.session_state.active_section = "Admin"
                st.rerun()

    if st.session_state.active_section not in ("Proposal Library", "Admin"):
        proposal_id = st.session_state.get("current_proposal_id")
        active = st.session_state.active_section
        required_sections = get_required_sections()

        # Put the current section completion control first so it is easy to find.
        if active in required_sections:
            st.markdown("### Current Section")
            st.caption(active)
            saved_key = f"complete_{active}"
            widget_key = f"sidebar_complete_{active}"
            if widget_key not in st.session_state:
                st.session_state[widget_key] = st.session_state.get(saved_key, False)

            checked = st.checkbox(
                f"Mark {active} complete",
                key=widget_key,
                help="This controls proposal progress and whether the proposal is ready to generate."
            )
            if checked != st.session_state.get(saved_key, False):
                st.session_state[saved_key] = checked
                if checked and not st.session_state.get("current_proposal_id"):
                    new_id = save_proposal(
                        None,
                        st.session_state.proposal_name,
                        st.session_state.credit_union,
                        st.session_state.proposal_type,
                        st.session_state.proposal_status,
                        collect_saved_data(),
                        st.session_state.msr,
                        st.session_state.current_user,
                        copied_from_proposal_id=st.session_state.get("copied_from_proposal_id"),
                    )
                    st.session_state.current_proposal_id = new_id
                    lock_proposal(new_id, st.session_state.current_user)
                elif st.session_state.get("current_proposal_id"):
                    auto_save_proposal()

            if checked:
                if st.button("Continue to Next Section →", key=f"continue_{active}", use_container_width=True):
                    go_to_next_section(active)
                    st.rerun()

            st.markdown("---")

        st.markdown("### Current Proposal")
        if proposal_id:
            st.caption(f"#{proposal_id} · {st.session_state.get('credit_union', '')}")
        st.markdown(f"**{st.session_state.get('proposal_name', '')}**")
        st.caption(f"{st.session_state.get('proposal_type', '')} · {st.session_state.get('proposal_status', 'Draft')}")

        if st.session_state.get("copied_from_proposal_id"):
            st.caption(f"Created from Proposal #{st.session_state.copied_from_proposal_id}")

        st.markdown("---")
        st.markdown("### Progress")

        completed_count = sum(
            1 for sec in required_sections
            if st.session_state.get(f"complete_{sec}", False)
        )
        total_required = len(required_sections)
        progress = completed_count / total_required if total_required else 1.0
        st.progress(progress)
        st.caption(f"{completed_count} of {total_required} sections complete · {progress:.0%}")

        if completed_count == total_required:
            st.success("✅ Ready to Generate")

        st.markdown("### Proposal Workflow")
        for sec in get_workflow_sections():
            selected = st.session_state.active_section == sec
            if sec == "Generate Proposal":
                icon = "🚀" if completed_count == total_required else "🔒"
            else:
                icon = "✅" if st.session_state.get(f"complete_{sec}", False) else "⬜"

            label = f"{icon}  {'▶ ' if selected else ''}{sec}"
            if st.button(label, key=f"nav_{sec}", use_container_width=True):
                st.session_state.active_section = sec
                st.rerun()

        st.markdown("---")
        if proposal_id and st.button("Close Proposal", key="close_current_proposal", use_container_width=True):
            unlock_proposal(proposal_id, st.session_state.current_user)
            st.session_state.current_proposal_id = None
            st.session_state.active_section = "Proposal Library"
            st.rerun()

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
# Progress is displayed persistently in the sidebar.
# -----------------------------

# -----------------------------
# Shared calculations
# -----------------------------
selected_targets = get_selected_targets()
selected_components = get_selected_components()
total_targets = sum(count for count, _ in selected_targets)

conversion_rate_decimal = parse_percent(st.session_state.conversion_rate)
estimated_loans_refinanced = round(total_targets * conversion_rate_decimal)
amount_refinanced = estimated_loans_refinanced * st.session_state.avg_loan_balance
loan_interest_rate_decimal = parse_percent(st.session_state.loan_interest_rate)
estimated_first_year_interest = calculate_first_year_interest(
    amount_refinanced,
    loan_interest_rate_decimal,
    st.session_state.loan_term_years,
)

costs = calculate_costs()

# ============================================================
# Delete confirmation dialog
# ============================================================
@st.dialog("Confirm Delete")
def confirm_delete_proposal(proposal_id, proposal_name, credit_union):
    st.markdown(f"**{proposal_name}**")
    st.caption(f"Proposal #{proposal_id} · {credit_union}")
    st.warning("This will permanently remove this proposal from the Proposal Library. This action cannot be undone.")

    cancel_col, confirm_col = st.columns(2)

    with cancel_col:
        if st.button(
            "Cancel",
            key=f"cancel_delete_{proposal_id}",
            use_container_width=True,
        ):
            st.rerun()

    with confirm_col:
        if st.button(
            "Delete Proposal",
            key=f"confirm_delete_{proposal_id}",
            type="primary",
            use_container_width=True,
        ):
            delete_proposal(proposal_id)
            st.rerun()


# ============================================================
# Admin
# ============================================================
if section == "Admin":
    st.subheader("Admin · Pricing & Configuration")

    if st.session_state.current_user not in ADMIN_USERS:
        st.error("Your selected user is not configured for Admin access.")
        st.stop()

    st.info(
        "Pricing changes apply to NEW proposals only. Every proposal freezes the rate schedule "
        "it started with, so changing a rate here will not recalculate an existing proposal."
    )
    st.caption(
        "During this development phase, Admin visibility is based on the user selected in the sidebar. "
        "When the app moves internally, this should be tied to authenticated Synergent user accounts."
    )

    pricing_tab, history_tab, storage_tab = st.tabs(["Pricing Settings", "Pricing History", "Storage Status"])

    with pricing_tab:
        current_settings = get_pricing_settings(include_inactive=True)
        categories = []
        for item in current_settings:
            if item["category"] not in categories:
                categories.append(item["category"])

        with st.form("admin_pricing_form"):
            edited_values = {}
            edited_repeating = {}
            edited_active = {}

            for category in categories:
                st.markdown(f"### {category}")
                category_items = [x for x in current_settings if x["category"] == category]

                for item in category_items:
                    key = item["key"]
                    label = item["label"]
                    value_type = item.get("value_type", "currency")
                    unit = item.get("unit", "")

                    if category == "Fixed Costs":
                        c1, c2, c3 = st.columns([0.58, 0.22, 0.20])
                    else:
                        c1, c2 = st.columns([0.72, 0.28])

                    with c1:
                        label_suffix = f" · {unit}" if unit else ""
                        st.markdown(f"**{label}**{label_suffix}")
                        if item.get("description"):
                            st.caption(item["description"])

                    with c2:
                        step = 0.5 if value_type == "percent" else 1.0
                        edited_values[key] = st.number_input(
                            f"Value · {label}",
                            min_value=0.0,
                            value=float(item["value"]),
                            step=step,
                            key=f"admin_value_{key}",
                            label_visibility="collapsed",
                        )

                    if category == "Fixed Costs":
                        with c3:
                            edited_repeating[key] = st.checkbox(
                                "Repeat / campaign",
                                value=bool(item.get("is_repeating", False)),
                                key=f"admin_repeat_{key}",
                            )
                            edited_active[key] = st.checkbox(
                                "Active",
                                value=bool(item.get("active", True)),
                                key=f"admin_active_{key}",
                            )
                    else:
                        edited_repeating[key] = bool(item.get("is_repeating", False))
                        edited_active[key] = bool(item.get("active", True))

                st.markdown("---")

            save_pricing_changes = st.form_submit_button("Save Pricing Changes", type="primary")

        if save_pricing_changes:
            changes = 0
            for item in current_settings:
                key = item["key"]
                if update_pricing_setting(
                    key,
                    edited_values[key],
                    st.session_state.current_user,
                    is_repeating=edited_repeating[key],
                    active=edited_active[key],
                ):
                    changes += 1
            if changes:
                st.success(f"Saved {changes} pricing change{'s' if changes != 1 else ''}. New proposals will use the updated schedule.")
            else:
                st.info("No pricing values changed.")
            st.rerun()

        st.markdown("### Add a Fixed Cost")
        st.caption("New active fixed costs automatically appear in the Cost Estimator for new proposals.")
        with st.form("add_fixed_cost_form", clear_on_submit=True):
            add1, add2, add3 = st.columns([0.55, 0.25, 0.20])
            with add1:
                new_fixed_label = st.text_input("Cost name", placeholder="Example: Custom Landing Page Setup")
            with add2:
                new_fixed_value = st.number_input("Amount", min_value=0.0, value=0.0, step=25.0)
            with add3:
                new_fixed_repeating = st.checkbox("Repeats per campaign")
            add_fixed_clicked = st.form_submit_button("Add Fixed Cost")

        if add_fixed_clicked:
            if not new_fixed_label.strip():
                st.error("Enter a name for the fixed cost.")
            elif new_fixed_value <= 0:
                st.error("Enter an amount greater than $0.")
            else:
                add_fixed_cost(new_fixed_label, new_fixed_value, st.session_state.current_user, new_fixed_repeating)
                st.success(f"Added {new_fixed_label.strip()} to the current pricing schedule.")
                st.rerun()

    with history_tab:
        st.markdown("### Recent Pricing Changes")
        history = get_pricing_history(200)
        settings_by_key = {x["key"]: x for x in get_pricing_settings(include_inactive=True)}
        if not history:
            st.info("No pricing changes have been recorded yet.")
        else:
            for row in history:
                setting = settings_by_key.get(row.get("setting_key"), {})
                label = setting.get("label", row.get("setting_key", "Pricing Setting"))
                old_value = float(row.get("old_value", 0) or 0)
                new_value = float(row.get("new_value", 0) or 0)
                unit = setting.get("unit", "")
                if setting.get("value_type") == "percent":
                    change_text = f"{old_value:,.1f}% → {new_value:,.1f}%"
                else:
                    change_text = f"${old_value:,.2f} → ${new_value:,.2f}"
                metadata_changes = []
                if row.get("old_is_repeating") is not None and row.get("new_is_repeating") is not None:
                    if bool(row.get("old_is_repeating")) != bool(row.get("new_is_repeating")):
                        metadata_changes.append(
                            "repeat per campaign ON" if bool(row.get("new_is_repeating")) else "repeat per campaign OFF"
                        )
                if row.get("old_active") is not None and row.get("new_active") is not None:
                    if bool(row.get("old_active")) != bool(row.get("new_active")):
                        metadata_changes.append("activated" if bool(row.get("new_active")) else "deactivated")

                st.markdown(f"**{label}** · {change_text}")
                detail = f"Changed {format_display_datetime(row.get('changed_at', ''))} by {row.get('changed_by', '')}{' · ' + unit if unit else ''}"
                if metadata_changes:
                    detail += " · " + ", ".join(metadata_changes)
                st.caption(detail)
                st.markdown("---")

    with storage_tab:
        db_status = get_backend_status()
        file_status = get_storage_status()
        c1, c2 = st.columns(2)
        with c1:
            st.metric("Proposal Database", db_status["data_mode"])
            st.caption(db_status["database_location"])
        with c2:
            st.metric("Proposal Files", file_status["file_mode"])
            if file_status.get("bucket"):
                st.caption(f"Private bucket: {file_status['bucket']}")

        if db_status["data_mode"] == "Local SQLite":
            st.warning(
                "Cloud persistence is not connected yet. The app is still using its local SQLite database "
                "and local generated files. Follow CLOUD_SETUP.md when you're ready to connect Supabase."
            )
        else:
            st.success("Cloud persistence is enabled. Proposal records and pricing settings are using Supabase.")

# ============================================================
# Proposal Library
# ============================================================
elif section == "Proposal Library":
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
            ["All"] + MSR_OPTIONS
        )

    results = search_proposals(search_text, status_filter, msr_filter)

    st.markdown("### Start New Proposal")

    if st.button("➕ Create New Proposal", key="create_new_proposal", type="secondary"):
        reset_proposal_state()
        st.rerun()

    st.markdown("---")

    if not results:
        st.info("No saved proposals found.")
    else:
        st.markdown("### Saved Proposals")

        h1, h2, h3, h4, h5, h6, h7 = st.columns(
           [0.23, 0.15, 0.11, 0.09, 0.14, 0.14, 0.14]
        )

        with h1: st.markdown("<span style='font-size:15px; font-weight:700;'>Proposal</span>", unsafe_allow_html=True)
        with h2: st.markdown("<span style='font-size:15px; font-weight:700;'>Credit Union</span>", unsafe_allow_html=True)
        with h3: st.markdown("<span style='font-size:15px; font-weight:700;'>Status</span>", unsafe_allow_html=True)
        with h4: st.markdown("<span style='font-size:15px; font-weight:700;'>MSR</span>", unsafe_allow_html=True)
        with h5: st.markdown("<span style='font-size:15px; font-weight:700;'>Last Updated</span>", unsafe_allow_html=True)
        with h6: st.markdown("<span style='font-size:15px; font-weight:700;'>Actions</span>", unsafe_allow_html=True)
        with h7: st.markdown("<span style='font-size:15px; font-weight:700;'>Files</span>", unsafe_allow_html=True)
        
        st.markdown("---")

        for proposal_id, proposal_name, credit_union, proposal_type, status, updated_at, msr, updated_by, locked_by, locked_at in results:
            col1, col2, col3, col4, col5, col6, col7 = st.columns(
                [0.23, 0.15, 0.11, 0.09, 0.14, 0.14, 0.14]
            )

            display_msr = canonical_msr_name(msr)
            display_updated_by = canonical_user_name(updated_by)
            display_locked_by = canonical_user_name(locked_by)

            with col1:
                st.write(f"**{proposal_name}**")
                st.caption(f"Proposal #{proposal_id} · {proposal_type}")
                
                if display_locked_by and display_locked_by != st.session_state.current_user:
                    st.caption(f"🔒 Editing: {display_locked_by}")

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
                 
                 current_status = status

                 if current_status not in status_options:
                     current_status = "Draft"
             
                 next_status = status_options[
                     (status_options.index(current_status) + 1) % len(status_options)
                 ]
             
                 pill_col, button_col = st.columns([0.70, 0.30])
             
                 with pill_col:
                     st.markdown(
                         f"""
                         <div style="
                             background-color:{status_colors[current_status]};
                             color:white;
                             border-radius:20px;
                             padding:4px 10px;
                             font-size:14px;
                             font-weight:700;
                             text-align:center;
                             width:82px;
                             line-height:18px;
                             margin-top:4px;
                         ">
                             {current_status}
                         </div>
                         """,
                         unsafe_allow_html=True
                     )
             
                 with button_col:
                     if st.button(
                         "↻",
                         key=f"cycle_status_{proposal_id}_{current_status}",
                         help=f"Change status to {next_status}",
                         use_container_width=True,
                     ):
                         update_proposal_status(proposal_id, next_status)
                         st.rerun()

            with col4:
                st.write(display_msr or "")

            with col5:
                st.caption(f"Last updated: {format_display_datetime(updated_at)}")
                if display_updated_by:
                    st.caption(f"By: {display_updated_by}")

            with col6:
                # Match the Files controls: four equal slots with three compact icon buttons.
                edit_col, copy_col, delete_col, _action_spacer = st.columns(4)

                with edit_col:
                    if st.button(
                        "✎",
                        key=f"open_{proposal_id}",
                        type="secondary",
                        help="Edit proposal",
                        use_container_width=True,
                    ):
                        saved_data = load_proposal(proposal_id)
                        if saved_data:
                            reset_proposal_state()
                            lock_proposal(proposal_id, st.session_state.current_user)
                            load_saved_data(saved_data)
                            st.session_state.current_proposal_id = proposal_id
                            st.session_state.active_section = "Proposal Details"
                            st.rerun()

                with copy_col:
                    if st.button(
                        "⧉",
                        key=f"duplicate_{proposal_id}",
                        help="Copy proposal",
                        use_container_width=True,
                    ):
                        new_id = duplicate_proposal(proposal_id)
                        if new_id:
                            st.rerun()
                        else:
                            st.error("Unable to copy this proposal.")

                with delete_col:
                    if st.button(
                        "✕",
                        key=f"delete_{proposal_id}",
                        type="primary",
                        help="Delete proposal",
                        use_container_width=True,
                    ):
                        confirm_delete_proposal(
                            proposal_id,
                            proposal_name,
                            credit_union,
                        )

            with col7:
                saved_data = load_proposal(proposal_id)
            
                if isinstance(saved_data, str):
                    saved_data = json.loads(saved_data)
            
                file_path = saved_data.get("file_path")
                sent_file_path = saved_data.get("sent_file_path")
                signed_file_path = saved_data.get("signed_file_path")
                pricing_export_path = saved_data.get("pricing_export_path")
            
                dl1, dl2, dl3, dl4 = st.columns(4)
            
                with dl1:
                    if file_path and stored_file_exists(file_path):
                        st.download_button(
                            "📄",
                            data=read_bytes(file_path),
                            file_name=display_name(file_path),
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            key=f"draft_{proposal_id}",
                            help="Download Draft",
                            use_container_width=True,
                        )
                    else:
                        st.button("📄", key=f"draft_missing_{proposal_id}", disabled=True, help="Draft not generated yet", use_container_width=True)

                with dl2:
                    if sent_file_path and stored_file_exists(sent_file_path):
                        st.download_button(
                            "📤",
                            data=read_bytes(sent_file_path),
                            file_name=display_name(sent_file_path),
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            key=f"sent_{proposal_id}",
                            help="Download Sent Proposal",
                            use_container_width=True,
                        )
                    else:
                        st.button("📤", key=f"sent_missing_{proposal_id}", disabled=True, help="No sent version yet", use_container_width=True)
            
                with dl3:
                    if signed_file_path and stored_file_exists(signed_file_path):
                        st.download_button(
                            "✅",
                            data=read_bytes(signed_file_path),
                            file_name=display_name(signed_file_path),
                            mime=(
                                "application/pdf"
                                if display_name(signed_file_path).lower().endswith(".pdf")
                                else "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            ),
                            key=f"signed_{proposal_id}",
                            help="Download Signed Proposal",
                            use_container_width=True,
                        )
                    else:
                        st.button("✅", key=f"signed_missing_{proposal_id}", disabled=True, help="No signed version yet", use_container_width=True)
            
                with dl4:
                    if pricing_export_path and stored_file_exists(pricing_export_path):
                        st.download_button(
                            "💲",
                            data=read_bytes(pricing_export_path),
                            file_name=display_name(pricing_export_path),
                            mime="text/csv",
                            key=f"pricing_{proposal_id}",
                            help="Download Pricing Export",
                            use_container_width=True,
                        )
                    else:
                        st.button("💲", key=f"pricing_missing_{proposal_id}", disabled=True, help="Pricing export not generated yet", use_container_width=True)

            st.markdown("<hr style='margin: 4px 0;'>", unsafe_allow_html=True)

# ============================================================
# Proposal Details
# ============================================================
elif section == "Proposal Details":
    st.subheader("Proposal Details")
    if st.session_state.get("copied_from_proposal_id"):
        st.info(f"This draft was created from Proposal #{st.session_state.copied_from_proposal_id}. Changes here will not affect the original proposal.")
    section_complete_checkbox("Proposal Details")

    previous_type = st.session_state.get("proposal_type")

    selected_type = st.selectbox(
        "Select Proposal Template",
        list(TEMPLATE_MAP.keys()),
        index=list(TEMPLATE_MAP.keys()).index(
            st.session_state.proposal_type
        ),
    )
    
    # Auto-update default proposal name ONLY when template changes
    if selected_type != previous_type:
    
        defaults = {
            "Auto Loan Recapture Campaign":
                "ACH Auto Loan Recapture Proposal",
    
            "Synergent Email Platform Proposal":
                "Synergent Email Platform Proposal",
        }
    
        st.session_state.proposal_name = defaults.get(
            selected_type,
            ""
        )

    st.session_state.proposal_type = selected_type

    selected_template = TEMPLATE_MAP[st.session_state.proposal_type]

    # Lending-only proposal fields
    if st.session_state.proposal_type == "General Lending Campaign":
        st.markdown("### Lending Campaign Details")
    
        loan_type_options = [
            "Auto Loan",
            "Personal Loan",
            "Vacation Loan",
            "Home Equity Loan",
            "Mortgage",
            "Other"
        ]
    
        selected_loan_type = st.selectbox(
            "Loan Type",
            loan_type_options,
            index=loan_type_options.index(st.session_state.loan_type)
            if st.session_state.loan_type in loan_type_options else loan_type_options.index("Other")
        )
    
        if selected_loan_type == "Other":
            st.session_state.loan_type = st.text_input(
                "Custom Loan Type",
                st.session_state.get("loan_type", "")
            )
        else:
            st.session_state.loan_type = selected_loan_type
    
        # Auto-update proposal name for lending proposal
        st.session_state.proposal_name = f"{st.session_state.loan_type} Campaign Proposal"
    
        campaign_week_options = ["4", "8", "12", "16", "Custom"]
    
        selected_campaign_weeks = st.selectbox(
            "Campaign Duration",
            campaign_week_options,
            index=campaign_week_options.index(str(st.session_state.campaign_weeks))
            if str(st.session_state.campaign_weeks) in campaign_week_options else campaign_week_options.index("Custom")
        )
    
        if selected_campaign_weeks == "Custom":
            st.session_state.campaign_weeks = st.number_input(
                "Custom Campaign Weeks",
                min_value=1,
                max_value=52,
                value=int(st.session_state.get("campaign_weeks", 8)),
                step=1
            )
        else:
            st.session_state.campaign_weeks = int(selected_campaign_weeks)
    
    if st.session_state.proposal_type == "Credit Card Campaign":

        st.markdown("### Credit Card Campaign Goals")
    
        goals = []
    
        if st.checkbox("Balance Transfer"):
            goals.append("Balance Transfer")
    
        if st.checkbox("New Card Acquisition"):
            goals.append("New Card Acquisition")
    
        if st.checkbox("Card Utilization & Activation"):
            goals.append("Card Utilization & Activation")
    
        st.session_state.credit_card_goals = goals
    
        st.session_state.proposal_name = (
            " + ".join(goals)
            + " Credit Card Campaign Proposal"
            if goals
            else "Credit Card Campaign Proposal"
        )

    if not Path(selected_template).exists():
        st.warning(f"Template file not found yet: {selected_template}")

    col1, col2 = st.columns(2)

    with col1:
        st.session_state.proposal_name = st.text_input(
            "Proposal Name",
            st.session_state.proposal_name,
        )

        credit_union_list = load_credit_union_list()

        ADD_NEW_CU_OPTION = "➕ Add New Credit Union"
        
        credit_union_options = credit_union_list + [ADD_NEW_CU_OPTION]
        
        current_credit_union = st.session_state.get(
            "credit_union",
            ""
        )
        
        if current_credit_union in credit_union_list:
            credit_union_index = credit_union_options.index(
                current_credit_union
            )
        else:
            credit_union_index = 0
        
        
        selected_credit_union = st.selectbox(
            "Credit Union Name",
            credit_union_options,
            index=credit_union_index
        )
        
        if selected_credit_union == ADD_NEW_CU_OPTION:
        
            custom_name = st.text_input(
                "Custom Credit Union Name",
                value="",
                key="custom_credit_union_name"
            )
        
            st.session_state.credit_union = custom_name
        
            save_new_cu = st.checkbox(
                "Save to Credit Union List",
                value=False
            )
        
            if (
                save_new_cu
                and custom_name.strip()
                and custom_name.lower()
                not in [x.lower() for x in credit_union_list]
            ):
        
                with open(
                    "CU List.csv",
                    "a",
                    encoding="utf-8-sig"
                ) as f:
                    f.write(f"\n{custom_name.strip()}")
        
                st.success(
                    f"Added '{custom_name.strip()}'"
                )
    
        else:
            st.session_state.credit_union = (
                selected_credit_union
            )

    with col2:
        st.session_state.proposal_date = st.date_input(
            "Proposal Date",
            st.session_state.proposal_date,
        )

        options = MSR_OPTIONS

        st.session_state.msr = st.selectbox(
            "MSR",
            options,
            index=options.index(st.session_state.get("msr"))
            if st.session_state.get("msr") in options else 0
        )
    
    if "proposal_notes" not in st.session_state:
        st.session_state.proposal_notes = ""

    st.session_state.proposal_notes = st.text_area(
        "Proposal Notes",
        value=st.session_state.proposal_notes,
        height=150,
        placeholder="Internal notes about proposal decisions..."
    )

    
    auto_save_proposal()

# ============================================================
# Campaign Targets
# ============================================================
elif section == "Campaign Targets":
    st.subheader("Campaign Targets")
    st.caption("Select target segments, adjust counts, or add custom targets.")
    
    if st.session_state.proposal_type == "Credit Card Campaign":

        st.markdown("### Credit Card Campaign Targets")
    
        if "Balance Transfer" in st.session_state.credit_card_goals:
            credit_card_target_section(
                "Balance Transfer",
                "cc_bt",
                ["Members with recurring ACH payments to external credit card providers"]
            )
    
        if "New Card Acquisition" in st.session_state.credit_card_goals:
            credit_card_target_section(
                "New Card Acquisition",
                "cc_new",
                ["Members with checking and no credit card"]
            )
    
        if "Card Utilization & Activation" in st.session_state.credit_card_goals:
            credit_card_target_section(
                "Card Utilization & Activation",
                "cc_util",
                ["Inactive credit card holders"]
            )

    else:

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
    
    
    if st.session_state.proposal_type == "Credit Card Campaign":

        st.markdown("### Credit Card Inputs")

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.session_state.conversion_rate = st.text_input(
                "Target Conversion Rate",
                st.session_state.conversion_rate,
            )

        with col2:
            st.session_state.avg_credit_card_limit = st.number_input(
                "Average Credit Card Limit",
                min_value=0,
                value=int(st.session_state.get("avg_credit_card_limit", 5000)),
                step=500,
            )

        with col3:
            st.session_state.avg_credit_card_rate = st.text_input(
                "Average Credit Card Rate",
                st.session_state.get("avg_credit_card_rate", "14.99%"),
            )

        with col4:
            st.session_state.avg_interchange_per_card = st.number_input(
                "Estimated First-Year Interchange Per Card",
                min_value=0,
                value=int(st.session_state.get("avg_interchange_per_card", 75)),
                step=5,
            )
        
        selected_targets = get_selected_credit_card_targets()
        total_targets = sum(count for count, _ in selected_targets)

        conversion_rate_decimal = parse_percent(st.session_state.conversion_rate)
        estimated_credit_cards_opened = round(total_targets * conversion_rate_decimal)

        estimated_total_credit_card_balance_transferred = (
            estimated_credit_cards_opened * st.session_state.avg_credit_card_limit
        )

        avg_credit_card_rate_decimal = parse_percent(st.session_state.avg_credit_card_rate)

        estimated_first_year_interest = (
            estimated_total_credit_card_balance_transferred
            * avg_credit_card_rate_decimal
        )

        estimated_first_year_interchange = (
            estimated_credit_cards_opened
            * st.session_state.avg_interchange_per_card
        )

        st.markdown("### Calculated Results")

        col5, col6, col7 = st.columns(3)

        with col5:
            st.text_input(
                "Estimated Credit Cards Opened",
                value=f"{estimated_credit_cards_opened:,}",
                disabled=True,
            )

        with col6:
            st.text_input(
                "Estimated Total Credit Card Balance Transferred",
                value=f"${estimated_total_credit_card_balance_transferred:,.0f}",
                disabled=True,
            )

        with col7:
            st.text_input(
                "Estimated First-Year Interest",
                value=f"${estimated_first_year_interest:,.0f}",
                disabled=True,
            )

        col8, col9 = st.columns(2)

        with col8:
            st.text_input(
                "Estimated First-Year Interchange",
                value=f"${estimated_first_year_interchange:,.0f}",
                disabled=True,
            )

        with col9:
            st.text_input(
                "Estimated Total First-Year Revenue",
                value=f"${estimated_first_year_interest + estimated_first_year_interchange:,.0f}",
                disabled=True,
            )

        st.caption(
            "Estimated cards opened = total targets × conversion rate. "
            "Estimated first-year interchange assumes $75 per opened credit card."
        )

    else:

        st.markdown("### Inputs")
    
        col1, col2, col3, col4 = st.columns(4)
    
        with col1:
            st.session_state.conversion_rate = st.text_input(
                "Target Conversion Rate",
                st.session_state.conversion_rate,
            )
        
        with col2:
            st.session_state.avg_loan_balance = st.number_input(
                "Average Loan Amount / Balance",
                min_value=0,
                value=int(st.session_state.get("avg_loan_balance", 25000)),
                step=500,
            )
        
        with col3:
            st.session_state.loan_interest_rate = st.text_input(
                "Average Loan Interest Rate",
                st.session_state.get("loan_interest_rate", "6.99%"),
            )
        
        with col4:
            st.session_state.loan_term_years = st.number_input(
                "Average Loan Term Years",
                min_value=1,
                value=int(st.session_state.get("loan_term_years", 5)),
                step=1,
            )
    
        selected_targets = get_selected_targets()
        total_targets = sum(count for count, _ in selected_targets)
    
        conversion_rate_decimal = parse_percent(st.session_state.conversion_rate)
        estimated_loans_refinanced = round(total_targets * conversion_rate_decimal)
        amount_refinanced = estimated_loans_refinanced * st.session_state.avg_loan_balance
        loan_interest_rate_decimal = parse_percent(st.session_state.loan_interest_rate)
        estimated_first_year_interest = calculate_first_year_interest(
            amount_refinanced,
            loan_interest_rate_decimal,
            st.session_state.loan_term_years,
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
    section_complete_checkbox("Campaign Components")


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

    
    auto_save_proposal()

# ============================================================
# Cost Estimator
# ============================================================
elif section == "Cost Estimator":
    st.subheader("Estimated Costs")
    st.caption("Select internal cost items. These inputs calculate proposal pricing but do not appear in the proposal.")
    section_complete_checkbox("Cost Estimator") 

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
        st.caption(f"Estimated at ${pricing_value('creative_hourly_rate', 115):,.2f}/hour")

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
        st.caption(f"Estimated at ${pricing_value('programming_hourly_rate', 200):,.2f}/hour")

    st.markdown("### Marked-Up Cost Items")
    st.caption(f"List markup: {pricing_value('list_markup_pct', 35):,.1f}% · Print markup: {pricing_value('print_markup_pct', 35):,.1f}%")

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
        st.caption(f"Estimated at ${pricing_value('email_development_hourly_rate', 115):,.2f}/hour")

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
        st.caption(f"Charged at ${pricing_value('email_send_fee', 100):,.2f} per email send.")

    st.markdown("### Fixed Cost Items")

    for item in get_fixed_cost_records():
        name = item["label"]
        cost = item["value"]
        repeat_note = " · repeats per campaign" if item.get("is_repeating") else ""
        st.checkbox(
            f"{name} — {money(cost)}{repeat_note}",
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
        st.caption(f"List Procurement includes a {pricing_value('list_markup_pct', 35):,.1f}% markup; Variable Print includes a {pricing_value('print_markup_pct', 35):,.1f}% markup.")

        st.markdown("#### Email Send Costs")
        st.write(f"Email Sends: {money(costs['email_send_cost'])}")

        st.markdown("#### Fixed Costs")
        st.write(f"Selected fixed cost items total: {money(costs['straight_cost_total'])}")

        st.markdown("#### Custom Costs")
        st.write(f"Custom costs total: {money(costs['custom_costs_total'])}")

        st.markdown("#### Pricing Logic")
        st.write(f"One-time cost total: {money(costs['one_time_cost_total'])}")
        st.write(f"Repeating cost total: {money(costs['repeating_cost_total'])}")
        st.write(f"Total estimated one-campaign cost: {money(costs['campaign_1_calc'])}")

        st.caption(
            "For 2+ campaign pricing, Variable Print Production, List Procurement, Email Sends, "
            f"and fixed costs marked as repeating recur per campaign. Four-campaign pricing applies a {pricing_value('four_campaign_discount_pct', 10):,.1f}% discount."
        )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.session_state.campaign_1_cost_override = st.text_input(
            "One Campaign Cost",
            value=st.session_state.get("campaign_1_cost_override", money(costs["campaign_1_calc"])),
        )

    with col2:
        st.session_state.campaign_2_cost_override = st.text_input(
            "Two Campaigns Total Cost",
            value=st.session_state.get("campaign_2_cost_override", money(costs["campaign_2_calc"])),
        )
        st.session_state.campaign_2_per_cost_override = st.text_input(
            "Two Campaigns Per Campaign Cost",
            value=st.session_state.get("campaign_2_per_cost_override", money(costs["campaign_2_per_calc"])),
        )

    with col3:
        st.session_state.campaign_4_cost_override = st.text_input(
            "Four Campaigns Total Cost",
            value=st.session_state.get("campaign_4_cost_override", money(costs["campaign_4_calc"])),
        )
        st.session_state.campaign_4_per_cost_override = st.text_input(
            "Four Campaigns Per Campaign Cost",
            value=st.session_state.get("campaign_4_per_cost_override", money(costs["campaign_4_per_calc"])),
        )

     
    auto_save_proposal()

# ============================================================
# EMP Details
# ============================================================

elif section == "EMP Details":
    st.subheader("EMP Details")
    section_complete_checkbox("EMP Details")

    st.session_state.total_subscribers = st.number_input(
        "Total Subscribers",
        min_value=0,
        value=st.session_state.get("total_subscribers", 15000),
        step=100,
    )

    tier_cost, tier_name = calculate_emp_tier_cost(
        st.session_state.total_subscribers
    )

    st.markdown("### Calculated Monthly Pricing")

    if tier_cost is None:
        st.warning("Subscriber count requires custom pricing.")
    else:
        emp = emp_pricing_details(st.session_state.total_subscribers)
        st.write(f"Tier: {tier_name}")
        st.write(f"Base: ${tier_cost:,.2f}")
        st.write(f"Essentials: ${emp['essentials_monthly']:,.2f}")
        st.write(f"Premium: ${emp['premium_monthly']:,.2f}")
        st.write(f"Elite: ${emp['elite_monthly']:,.2f}")
        st.caption(
            f"Implementation: Essentials {money(emp['essentials_implementation'])} · "
            f"Premium {money(emp['premium_implementation'])} · Elite {money(emp['elite_implementation'])}"
        )

    
    auto_save_proposal()

# ============================================================
# Generate Proposal
# ============================================================
elif section == "Generate Proposal":
    st.subheader("Generate Proposal")

    selected_template = TEMPLATE_MAP[st.session_state.proposal_type]
    is_emp_proposal = st.session_state.proposal_type == "Synergent Email Platform Proposal"

    # EMP proposals have their own subscriber/tier pricing model.  Do not let the
    # default campaign-target and campaign-cost session state leak into the EMP
    # review or generation workflow.
    if is_emp_proposal:
        selected_components = []
        selected_targets = []
        total_targets = 0
        conversion_rate_decimal = 0
        estimated_loans_refinanced = 0
        amount_refinanced = 0
        estimated_first_year_interest = 0
        target_roi = 0
        costs = {
            "campaign_1_calc": 0,
            "campaign_2_calc": 0,
            "campaign_2_per_calc": 0,
            "campaign_4_calc": 0,
            "campaign_4_per_calc": 0,
        }
        campaign_1_cost = ""
        campaign_2_cost = ""
        campaign_2_per_cost = ""
        campaign_4_cost = ""
        campaign_4_per_cost = ""
    else:
        selected_components = get_selected_components()
        if st.session_state.proposal_type == "Credit Card Campaign":
            selected_targets = get_selected_credit_card_targets()
        else:
            selected_targets = get_selected_targets()

        total_targets = sum(count for count, _ in selected_targets)

        conversion_rate_decimal = parse_percent(st.session_state.conversion_rate)
        estimated_loans_refinanced = round(total_targets * conversion_rate_decimal)
        amount_refinanced = estimated_loans_refinanced * st.session_state.avg_loan_balance
        loan_interest_rate_decimal = parse_percent(st.session_state.loan_interest_rate)
        estimated_first_year_interest = calculate_first_year_interest(
            amount_refinanced,
            loan_interest_rate_decimal,
            st.session_state.loan_term_years,
        )

        costs = calculate_costs()

        one_time_campaign_cost = costs["campaign_1_calc"]

        if one_time_campaign_cost > 0:
            target_roi = estimated_first_year_interest / one_time_campaign_cost
        else:
            target_roi = 0

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
    required_sections = get_required_sections()

    incomplete_sections = [
        sec for sec in required_sections
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

    if is_emp_proposal:
        tier_cost, tier_name = calculate_emp_tier_cost(st.session_state.total_subscribers)
        st.write(f"Total Subscribers: {st.session_state.total_subscribers:,}")
        st.write(f"Pricing Tier: {tier_name}")
        if tier_cost is None:
            st.write("Monthly Pricing: Custom pricing required")
        else:
            emp = emp_pricing_details(st.session_state.total_subscribers)
            st.write(f"Base Monthly Cost: {money(tier_cost)}")
            st.write(f"Essentials Monthly Cost: {money(emp['essentials_monthly'])}")
            st.write(f"Premium Monthly Cost: {money(emp['premium_monthly'])}")
            st.write(f"Elite Monthly Cost: {money(emp['elite_monthly'])}")
            st.caption(
                f"Implementation: Essentials {money(emp['essentials_implementation'])} · "
                f"Premium {money(emp['premium_implementation'])} · Elite {money(emp['elite_implementation'])}"
            )
    else:
        st.write(f"Total Targets: {total_targets:,}")
        st.write(f"Calculated One-Campaign Cost: {money(costs['campaign_1_calc'])}")
        st.write(f"Final One-Campaign Proposal Price: {campaign_1_cost}")

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
    # -----------------------------
    # Run Generation
    # -----------------------------
    with st.expander("1. Generate Draft", expanded=False):
       if st.session_state.get("file_path"):
        st.info(
            f"Current generated proposal:\n"
            f"{display_name(st.session_state.file_path)}"
        )
       if generate_clicked:
           emp_tier_cost, _ = calculate_emp_tier_cost(st.session_state.total_subscribers) if is_emp_proposal else (0, "")
           if not is_emp_proposal and not selected_targets:
               st.error("Please select at least one campaign target.")
           elif not is_emp_proposal and not selected_components:
               st.error("Please select at least one campaign component.")
           elif is_emp_proposal and emp_tier_cost is None:
               st.error("This subscriber count requires custom pricing.")
           elif not Path(selected_template).exists():
               st.error(f"Template file is missing: {selected_template}")
           else:
               gp.TEMPLATE_PATH = selected_template

               # generate_proposal.py keeps proposal_data at module scope. Remove EMP-only
               # keys before every run so generating an EMP cannot cause a later campaign
               # proposal to be treated as an EMP proposal.
               for emp_key in (
                   "{{total_subscribers}}",
                   "{{tier_cost}}",
                   "{{essentials_cost}}",
                   "{{premium_cost}}",
                   "{{elite_cost}}",
                   "{{emp_tier_number}}",
               ):
                   gp.proposal_data.pop(emp_key, None)
   
               gp.proposal_data["{{proposal_name}}"] = st.session_state.proposal_name
               gp.proposal_data["{{proposal_date}}"] = format_date_windows(
                   st.session_state.proposal_date
               )
               gp.proposal_data["{{creditunion_name}}"] = st.session_state.credit_union
               
               # Credit Card Objectives
               if st.session_state.proposal_type == "Credit Card Campaign":

                  gp.proposal_data["{{credit_card_campaign_objectives}}"] = (
                     build_credit_card_campaign_objectives()
                  )

               # -----------------------------
               # Credit Card Proposal Variables
               # -----------------------------
               if st.session_state.proposal_type == "Credit Card Campaign":

                  selected_targets = get_selected_credit_card_targets()
              
                  total_targets = sum(
                      count for count, _ in selected_targets
                  )
              
                  conversion_rate_decimal = parse_percent(
                      st.session_state.conversion_rate
                  )
              
                  estimated_credit_cards_opened = round(
                      total_targets * conversion_rate_decimal
                  )
              
                  estimated_total_card_balance = (
                      estimated_credit_cards_opened
                      * st.session_state.avg_credit_card_limit
                  )
              
                  avg_rate_decimal = parse_percent(
                      st.session_state.avg_credit_card_rate
                  )
              
                  estimated_first_year_interest = (
                      estimated_total_card_balance
                      * avg_rate_decimal
                  )
              
                  estimated_first_year_interchange = (
                      estimated_credit_cards_opened
                      * st.session_state.avg_interchange_per_card
                  )
              
                  gp.proposal_data[
                      "{{avg_credit_card_balance}}"
                  ] = f"${st.session_state.avg_credit_card_limit:,.0f}"
              
                  gp.proposal_data[
                      "{{avg_credit_card_rate}}"
                  ] = st.session_state.avg_credit_card_rate
              
                  gp.proposal_data[
                      "{{avg_interchange_per_card}}"
                  ] = f"${st.session_state.avg_interchange_per_card:,.0f}"
              
                  gp.proposal_data[
                      "{{estimated_credit_cards_opened}}"
                  ] = f"{estimated_credit_cards_opened:,}"
              
                  gp.proposal_data[
                      "{{estimated_total_card_balance}}"
                  ] = f"${estimated_total_card_balance:,.0f}"
              
                  gp.proposal_data[
                      "{{estimated_first_year_interest}}"
                  ] = f"${estimated_first_year_interest:,.0f}"
              
                  gp.proposal_data[
                      "{{estimated_first_year_interchange}}"
                  ] = f"${estimated_first_year_interchange:,.0f}"
              
                  target_lines = []
                  
                  for count, desc in selected_targets:
                      target_lines.append(
                          f"•  {count:,}  {desc}"
                      )
                  target_lines.append("")
                  target_lines.append(f"Total Targets (de-duped by SSN): {total_targets:,}")
                  gp.proposal_data["{{credit_card_target_segments}}"] = "\n".join(target_lines)
                  
   
               gp.proposal_data["{{target_conversion_rate}}"] = st.session_state.conversion_rate
               gp.proposal_data["{{total_targets}}"] = f"{total_targets:,}"
   
               gp.proposal_data["{{conversions}}"] = f"{estimated_loans_refinanced:,}"
               gp.proposal_data["{{amount_refinanced}}"] = f"${amount_refinanced:,.0f}"
               gp.proposal_data["{{first_year_interest}}"] = f"${estimated_first_year_interest:,.0f}"
               gp.proposal_data["{{loan_type}}"] = st.session_state.loan_type
               gp.proposal_data["{{campaign_weeks}}"] = str(st.session_state.campaign_weeks)
               gp.proposal_data["{{loan_interest_rate}}"] = st.session_state.loan_interest_rate
               gp.proposal_data["{{avg_loan_balance}}"] = f"${st.session_state.avg_loan_balance:,.0f}"
               gp.proposal_data["{{loan_term_years}}"] = str(st.session_state.loan_term_years)
   
               gp.proposal_data["{{total_targets_2}}"] = f"{total_targets * 2:,}"
               gp.proposal_data["{{total_targets_4}}"] = f"{total_targets * 4:,}"
   
               gp.proposal_data["{{campaign_1_cost}}"] = campaign_1_cost
               gp.proposal_data["{{campaign_2_cost}}"] = campaign_2_cost
               gp.proposal_data["{{campaign_2_per_cost}}"] = campaign_2_per_cost
               gp.proposal_data["{{campaign_4_cost}}"] = campaign_4_cost
               gp.proposal_data["{{campaign_4_per_cost}}"] = campaign_4_per_cost
   
               gp.proposal_data["{{target_ROI}}"] = f"${target_roi:.2f}"
   
               gp.target_segments.clear()
               gp.target_segments.extend(selected_targets)

               gp.credit_card_target_segments.clear()
               gp.credit_card_target_segments.extend(selected_targets)
   
               gp.campaign_components.clear()
               gp.campaign_components.extend(selected_components)
   
               # A short filename only needs the proposal ID, CU, type and generation version.
               # Save first if needed so every generated file always has a real proposal ID.
               if not st.session_state.get("current_proposal_id"):
                   proposal_id = save_proposal(
                       None,
                       st.session_state.proposal_name,
                       st.session_state.credit_union,
                       st.session_state.proposal_type,
                       st.session_state.proposal_status,
                       collect_saved_data(),
                       st.session_state.msr,
                       st.session_state.current_user,
                   )
                   st.session_state.current_proposal_id = proposal_id
               else:
                   proposal_id = st.session_state.current_proposal_id

               generation_version = get_next_generation_version(proposal_id)
               file_name = f"{proposal_file_stem(proposal_id)}_v{generation_version}.pptx"
               
               base_folder, drafts_folder, sent_folder, signed_folder, pricing_folder = get_credit_union_output_folder(
                   st.session_state.credit_union
               )
   
               file_path = os.path.join(drafts_folder, file_name)
   
               one_time_campaign_cost = costs["campaign_1_calc"]
   
               if one_time_campaign_cost > 0:
                   target_roi = estimated_first_year_interest / one_time_campaign_cost
               else:
                   target_roi = 0
               
               gp.proposal_data["{{target_ROI}}"] = f"${target_roi:.2f}"
               
               gp.proposal_data["{{total_targets_line}}"] = (
                   f"Total Targets (de-duped by SSN): {total_targets:,}"
               )
   
               if st.session_state.proposal_type == "Synergent Email Platform Proposal":
                   tier_cost, tier_name = calculate_emp_tier_cost(
                       st.session_state.total_subscribers
                   )
   
                   if tier_cost is None:
                       st.error("This subscriber count requires custom pricing.")
                       st.stop()
   
                   emp = emp_pricing_details(st.session_state.total_subscribers)
                   gp.proposal_data["{{total_subscribers}}"] = f"{st.session_state.total_subscribers:,}"
                   gp.proposal_data["{{tier_cost}}"] = f"${tier_cost:,.2f}"
                   gp.proposal_data["{{essentials_cost}}"] = f"${emp['essentials_monthly']:,.2f}"
                   gp.proposal_data["{{premium_cost}}"] = f"${emp['premium_monthly']:,.2f}"
                   gp.proposal_data["{{elite_cost}}"] = f"${emp['elite_monthly']:,.2f}"
                   gp.proposal_data["{{emp_tier_number}}"] = tier_name.replace("Tier ", "")
               
               gp.main(output_path=file_path)

               # Persist the generated draft. In cloud mode the local file is only staging.
               stored_draft_ref = store_file(
                   file_path,
                   make_object_path(st.session_state.credit_union, "Drafts", file_name),
               )

               # Save the generated draft first so the pricing export can include a proposal ID.
               st.session_state.file_path = stored_draft_ref
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

               # Pricing detail is a generation artifact, not a "mark sent" artifact.
               local_pricing_export_path = create_pricing_export_csv(pricing_folder)
               pricing_export_path = store_file(
                   local_pricing_export_path,
                   make_object_path(
                       st.session_state.credit_union,
                       "Pricing Exports",
                       os.path.basename(local_pricing_export_path),
                   ),
               )
               st.session_state.pricing_export_path = pricing_export_path

               # Every generation gets an immutable pricing audit snapshot.
               save_pricing_snapshot(
                   proposal_id,
                   build_pricing_audit_snapshot(),
                   st.session_state.current_user,
               )

               save_proposal(
                   proposal_id,
                   st.session_state.proposal_name,
                   st.session_state.credit_union,
                   st.session_state.proposal_type,
                   st.session_state.proposal_status,
                   collect_saved_data(),
                   st.session_state.msr,
                   st.session_state.current_user
               )

               st.success("Proposal and pricing detail generated successfully!")
               if cloud_file_mode():
                   st.caption("Saved to persistent cloud storage.")
               else:
                   st.caption(f"Saved to: {drafts_folder}")

       file_path = st.session_state.get("file_path")
       if file_path and stored_file_exists(file_path):
           st.download_button(
               label="Download Generated Proposal",
               data=read_bytes(file_path),
               file_name=display_name(file_path),
               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
               key="download_generated_proposal"
           )

       pricing_export_path = st.session_state.get("pricing_export_path")
       if pricing_export_path and stored_file_exists(pricing_export_path):
           st.download_button(
               label="Download Pricing Detail",
               data=read_bytes(pricing_export_path),
               file_name=display_name(pricing_export_path),
               mime="text/csv",
               key="download_generated_pricing"
           )

    # -----------------------------
    # Sent Proposal Snapshot
    # -----------------------------
    with st.expander("2. Send Snapshot", expanded=False):
        if st.session_state.get("file_path"):
           st.markdown("### Sent Proposal Snapshot")
           
           file_path = st.session_state.get("file_path")
           
           if file_path and stored_file_exists(file_path):
               st.info(f"Current generated proposal: {display_name(file_path)}")
           else:
               st.warning("Generate the proposal before marking it as sent.")
           
           if st.button("Mark Current Proposal as Sent to Credit Union"):
               file_path = st.session_state.get("file_path")
           
               if not file_path or not stored_file_exists(file_path):
                   st.error("Generate the proposal before marking it as sent.")
               else:
                   from datetime import datetime
           
                   base_folder, drafts_folder, sent_folder, signed_folder, pricing_folder = get_credit_union_output_folder(
                       st.session_state.credit_union
                   )
           
                   timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                   sent_file_name = lifecycle_filename("SENT", source_ref=file_path)

                   if cloud_file_mode():
                       sent_file_path = copy_stored_file(
                           file_path,
                           destination_object_path=make_object_path(
                               st.session_state.credit_union, "Sent", sent_file_name
                           ),
                       )
                   else:
                       sent_local_path = os.path.join(sent_folder, sent_file_name)
                       sent_file_path = copy_stored_file(
                           file_path, destination_local_path=sent_local_path
                       )

                   st.session_state.sent_file_path = sent_file_path
                   st.session_state.sent_at = timestamp
                   pricing_export_path = st.session_state.get("pricing_export_path")
                   if not pricing_export_path or not stored_file_exists(pricing_export_path):
                       local_pricing_export_path = create_pricing_export_csv(pricing_folder)
                       pricing_export_path = store_file(
                           local_pricing_export_path,
                           make_object_path(
                               st.session_state.credit_union,
                               "Pricing Exports",
                               os.path.basename(local_pricing_export_path),
                           ),
                       )
                       st.session_state.pricing_export_path = pricing_export_path
           
                   saved_data = collect_saved_data()
                   saved_data["file_path"] = file_path
                   saved_data["sent_file_path"] = sent_file_path
                   saved_data["sent_at"] = timestamp
                   saved_data["pricing_export_path"] = pricing_export_path
           
                   st.session_state.proposal_status = "CU Review"
       
                   update_proposal_status(
                       st.session_state.current_proposal_id,
                       "CU Review"
                   )
           
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
           
                   st.success(f"Sent snapshot saved: {sent_file_name}") 
                   if stored_file_exists(sent_file_path):
                       st.download_button(
                          label="Download Sent Proposal",
                          data=read_bytes(sent_file_path),
                          file_name=display_name(sent_file_path),
                          mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                          key="download_sent"
                       )
                   if pricing_export_path and stored_file_exists(pricing_export_path):
                       st.download_button(
                          label="Download Pricing Export",
                          data=read_bytes(pricing_export_path),
                          file_name=display_name(pricing_export_path),
                          mime="text/csv",
                          key="download_pricing"
                       )

    # -----------------------------
    # PDF Export
    # -----------------------------
    with st.expander("3. DocuSign Prep", expanded=False):
        if st.session_state.get("sent_file_path"):

            st.markdown("### PDF Export")
            
            if st.button("Create PDF from Sent Proposal"):
                sent_file_path = st.session_state.get("sent_file_path")

                if cloud_file_mode():
                    st.warning(
                        "PDF conversion uses desktop Microsoft PowerPoint and is not available on the "
                        "Streamlit cloud host. It will be available again when this app moves to the internal Windows server."
                    )
                elif not sent_file_path or not stored_file_exists(sent_file_path):
                    st.error("Mark the proposal as sent first.")
                else:
                    pdf_path = convert_pptx_to_pdf(sent_file_path)
                    st.session_state.sent_pdf_path = pdf_path
                    st.success("PDF created successfully.")
                    st.rerun()
            
            
            sent_pdf_path = st.session_state.get("sent_pdf_path")
            
            if sent_pdf_path and os.path.exists(sent_pdf_path):
            
                with open(sent_pdf_path, "rb") as file:
            
                    st.download_button(
                        label="Download Sent Proposal PDF",
                        data=file,
                        file_name=os.path.basename(sent_pdf_path),
                        mime="application/pdf",
                        key="download_sent_pdf"
                    )  
                     
    
        else:
            st.info("Mark proposal as sent to unlock DocuSign preparation.")

    # -----------------------------
    # Final Signed Proposal
    # -----------------------------
    with st.expander("4. Final Signed Proposal", expanded=False):
        sent_file_path = st.session_state.get("sent_file_path")
        signed_file_path = st.session_state.get("signed_file_path")

        if not sent_file_path or not stored_file_exists(sent_file_path):
            st.info("Mark the proposal as sent before uploading the completed DocuSign PDF.")
        else:
            st.markdown("### Upload Completed DocuSign PDF")
            st.caption(
                "After the proposal has been completed in DocuSign, upload the final signed PDF here. "
                "The Sent version remains on file separately."
            )

            if signed_file_path and stored_file_exists(signed_file_path):
                st.success(f"Signed PDF on file: {display_name(signed_file_path)}")
                signed_meta = []
                if st.session_state.get("signed_at"):
                    signed_meta.append(f"uploaded {format_display_datetime(st.session_state.signed_at)}")
                if st.session_state.get("signed_uploaded_by"):
                    signed_meta.append(f"by {st.session_state.signed_uploaded_by}")
                if signed_meta:
                    st.caption(" ".join(signed_meta))

                st.download_button(
                    label="Download Signed PDF",
                    data=read_bytes(signed_file_path),
                    file_name=display_name(signed_file_path),
                    mime="application/pdf",
                    key="download_existing_signed_pdf",
                )

                replace_signed = st.checkbox(
                    "Replace the signed PDF already on file",
                    value=False,
                    key="replace_existing_signed_pdf",
                    help="The replacement is saved as a new revision rather than overwriting the existing filename.",
                )
            else:
                replace_signed = True

            if replace_signed:
                uploaded_signed_pdf = st.file_uploader(
                    "Signed PDF",
                    type=["pdf"],
                    key="signed_pdf_uploader",
                    help="Upload the completed PDF returned from DocuSign.",
                )

                confirm_signed = st.checkbox(
                    "I confirm this is the completed signed proposal and should mark this proposal as Signed.",
                    value=False,
                    key="confirm_signed_pdf_upload",
                )

                upload_disabled = uploaded_signed_pdf is None or not confirm_signed
                if st.button(
                    "Upload & Mark Signed",
                    key="upload_and_mark_signed",
                    type="primary",
                    disabled=upload_disabled,
                ):
                    pdf_bytes = uploaded_signed_pdf.getvalue()

                    if not pdf_bytes.startswith(b"%PDF"):
                        st.error("The uploaded file does not appear to be a valid PDF.")
                    else:
                        previous_signed_ref = st.session_state.get("signed_file_path")
                        signed_name = signed_pdf_filename(
                            source_ref=sent_file_path,
                            existing_signed_ref=previous_signed_ref,
                        )
                        signed_object_path = make_object_path(
                            st.session_state.credit_union,
                            "Signed",
                            signed_name,
                        )

                        stored_signed_ref = store_bytes(
                            pdf_bytes,
                            signed_object_path,
                            content_type="application/pdf",
                        )

                        signed_timestamp = datetime.now(ZoneInfo("UTC")).isoformat()
                        st.session_state.proposal_status = "Signed"
                        st.session_state.signed_file_path = stored_signed_ref
                        st.session_state.signed_at = signed_timestamp
                        st.session_state.signed_uploaded_by = st.session_state.current_user
                        st.session_state.signed_original_name = uploaded_signed_pdf.name

                        proposal_id = save_proposal(
                            st.session_state.get("current_proposal_id"),
                            st.session_state.proposal_name,
                            st.session_state.credit_union,
                            st.session_state.proposal_type,
                            "Signed",
                            collect_saved_data(),
                            st.session_state.msr,
                            st.session_state.current_user,
                        )
                        st.session_state.current_proposal_id = proposal_id
                        update_proposal_status(proposal_id, "Signed")

                        st.success(f"Signed proposal saved: {signed_name}")
                        st.rerun()
