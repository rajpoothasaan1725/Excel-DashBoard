import base64
from datetime import datetime
import os

import pandas as pd
import streamlit as st

# =========================================================
# 1. PAGE CONFIGURATION
# =========================================================

st.set_page_config(page_title="Transworld Home Dashboard", layout="wide")


# =========================================================
# 2. SESSION STATE
# =========================================================

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "username" not in st.session_state:
    st.session_state.username = ""

if "role" not in st.session_state:
    st.session_state.role = ""


# =========================================================
# 3. FILE PATHS
# =========================================================

# Login background image
image_path = "login_bg.png"

# Excel file
excel_file_path = r"D:\As-Built\Dashboard\North As-Built Tracker.xlsx"

# Transworld logo
logo_path = r"D:\As-Built\Dashboard\transworld_logo.png"


# =========================================================
# 4. LOGIN PAGE
# =========================================================

if not st.session_state.logged_in:

    # =====================================================
    # LOAD LOGIN BACKGROUND
    # =====================================================

    if os.path.exists(login_bg_path):

        with open(login_bg_path, "rb") as image_file:
            encoded_image = base64.b64encode(image_file.read()).decode("utf-8")

        # =================================================
        # LOGIN PAGE CSS
        # =================================================

        st.markdown(
            f"""
            <style>

            /* Full Screen Background Image */
            .stApp {{
                background-image: url("data:image/png;base64,{encoded_image}");
                background-size: cover;
                background-position: center;
                background-repeat: no-repeat;
                background-attachment: fixed;
            }}

            /* Remove Streamlit Header & Adjust Top Padding */
            header, div[data-testid="stHeader"] {{
                display: none !important;
            }}

            .block-container {{
                padding-top: 5rem !important;
            }}

            /* TARGET STREAMLIT FORM AS THE OUTER DARK TRANSPARENT CARD */
            div[data-testid="stForm"] {{
                background: rgba(0, 0, 0, 0.65) !important; /* Semi-transparent dark overlay */
                border: 2px solid rgba(255, 255, 255, 0.3) !important;
                border-radius: 16px !important;
                padding: 30px 25px !important;
                box-shadow: 0px 10px 30px rgba(0, 0, 0, 0.7) !important;
                backdrop-filter: blur(8px) !important;
            }}

            /* Header Titles */
            .main-title {{
                color: #FFFFFF !important;
                font-size: 32px !important;
                font-weight: 700 !important;
                text-align: center !important;
                margin-bottom: 2px !important;
                font-family: sans-serif;
            }}

            .sub-title {{
                color: #E2E8F0 !important;
                font-size: 20px !important;
                font-weight: 500 !important;
                text-align: center !important;
                margin-bottom: 25px !important;
                font-family: sans-serif;
            }}

            /* Input Fields - Solid White Rectangles */
            div[data-testid="stTextInput"] input {{
                background-color: #FFFFFF !important;
                color: #000000 !important;
                border: none !important;
                border-radius: 4px !important;
                height: 42px !important;
                font-size: 16px !important;
            }}

            /* Bright Blue Button */
            div[data-testid="stFormSubmitButton"] > button {{
                background-color: #007bff !important;
                color: #FFFFFF !important;
                border: none !important;
                border-radius: 6px !important;
                height: 44px !important;
                font-size: 18px !important;
                font-weight: 600 !important;
                margin-top: 10px !important;
                box-shadow: 0px 4px 12px rgba(0, 123, 255, 0.4) !important;
            }}

            div[data-testid="stFormSubmitButton"] > button:hover {{
                background-color: #0056b3 !important;
                color: #FFFFFF !important;
            }}

            </style>
            """,
            unsafe_allow_html=True,
        )

    else:
        st.error(f"Login background image not found: {login_bg_path}")

    # =====================================================
    # LOGIN FORM LAYOUT
    # =====================================================

    col_left, col_center, col_right = st.columns([1, 1.3, 1])

    with col_center:
        # Native Streamlit Form encapsulates all elements inside a single div
        with st.form(key="login_form", clear_on_submit=False):

            # Titles
            st.markdown(
                "<div class='main-title'>Transworld Home</div>",
                unsafe_allow_html=True,
            )
            st.markdown(
                "<div class='sub-title'>Portal Login</div>",
                unsafe_allow_html=True,
            )

            # Username Row (Label Left, Input Right)
            u_label, u_input = st.columns([1, 2])
            with u_label:
                st.markdown(
                    "<p style='color: white; font-size: 19px; font-weight: 500;"
                    " margin-top: 8px;'>Username</p>",
                    unsafe_allow_html=True,
                )
            with u_input:
                username_input = st.text_input(
                    "Username",
                    label_visibility="collapsed",
                    key="login_username",
                )

            # Password Row (Label Left, Input Right)
            p_label, p_input = st.columns([1, 2])
            with p_label:
                st.markdown(
                    "<p style='color: white; font-size: 19px; font-weight: 500;"
                    " margin-top: 8px;'>Password</p>",
                    unsafe_allow_html=True,
                )
            with p_input:
                password_input = st.text_input(
                    "Password",
                    type="password",
                    label_visibility="collapsed",
                    key="login_password",
                )

            # Centered Blue Login Button
            b_left, b_center, b_right = st.columns([0.8, 1.4, 0.8])
            with b_center:
                submitted = st.form_submit_button(
                    "Login", use_container_width=True
                )

            if submitted:
                if (
                    username_input == "admin"
                    and password_input == "transworld123"
                ):
                    st.session_state.logged_in = True
                    st.session_state.username = "admin"
                    st.session_state.role = "admin"
                    st.rerun()

                elif (
                    username_input == "viewer"
                    and password_input == "transworldview"
                ):
                    st.session_state.logged_in = True
                    st.session_state.username = "viewer"
                    st.session_state.role = "viewer"
                    st.rerun()

                else:
                    st.error("❌ Invalid Username or Password")


# =========================================================
# 5. DASHBOARD
# =========================================================

else:

    # =====================================================
    # DASHBOARD CSS
    # =====================================================

    st.markdown(
        """
        <style>

        /* MAIN APP */
        .stApp {
            background-color: #FFFFFF !important;
        }

        /* REMOVE HEADER */
        header, div[data-testid="stHeader"] {
            background: transparent !important;
            height: 0px !important;
            display: none !important;
        }

        /* MAIN CONTAINER */
        .block-container {
            padding-top: 0.5rem !important;
            margin-top: -20px !important;
            padding-bottom: 1rem !important;
        }

        /* ANIMATED TITLE CSS */
        @keyframes slideInDown {
            0% {
                opacity: 0;
                transform: translateY(-20px);
            }
            100% {
                opacity: 1;
                transform: translateY(0);
            }
        }

        .header-title-animated {
            font-size: 26px !important;
            font-weight: 700 !important;
            color: #234263 !important;
            text-align: center !important;
            margin: 0 !important;
            padding-top: 5px !important;
            animation: slideInDown 0.8s ease-out forwards;
            white-space: nowrap !important;
        }

        /* TABLE HEADER */
        th {
            background-color: #234263 !important;
            color: white !important;
            font-weight: bold !important;
            text-align: center !important;
            padding: 10px !important;
        }

        /* TABLE CELLS */
        td {
            background-color: #FFFFFF !important;
            color: #31333F !important;
            padding: 8px !important;
            border: 1px solid #E6E6E6 !important;
        }

        /* HIDE TOOLBAR */
        div[data-testid="stToolbar"] {
            display: none !important;
        }

        /* SUMMARY CARDS */
        div[data-testid="stMetric"] {
            background-color: #FFFFFF !important;
            border: 1px solid #D9D9D9 !important;
            border-radius: 10px !important;
            padding: 20px !important;
            box-shadow: 0px 2px 8px rgba(0, 0, 0, 0.08) !important;
        }

        div[data-testid="stMetricLabel"] {
            color: #555555 !important;
            font-size: 16px !important;
            font-weight: 600 !important;
        }

        div[data-testid="stMetricValue"] {
            color: #234263 !important;
            font-size: 30px !important;
            font-weight: bold !important;
        }

        /* LOGOUT BUTTON */
        div.stButton > button[kind="primary"] {
            background-color: #234263 !important;
            color: #FFFFFF !important;
            border: 1px solid #234263 !important;
            border-radius: 6px !important;
            font-weight: 600 !important;
        }

        div.stButton > button[kind="primary"]:hover {
            background-color: #1A334D !important;
            color: #FFFFFF !important;
        }

        </style>
        """,
        unsafe_allow_html=True,
    )

    # =====================================================
    # HEADER
    # =====================================================

    col_logo, col_title, col_search, col_logout = st.columns([2, 4, 3, 1.2])

    with col_logo:
        if os.path.exists(logo_path):
            st.image(logo_path, width=180)

    with col_title:
        st.markdown(
            '<div class="header-title-animated">North As-Built Tracker</div>',
            unsafe_allow_html=True,
        )

    with col_search:
        search_query = st.text_input(
            "Search",
            placeholder="Search Phase / Project...",
            label_visibility="collapsed",
        )

    with col_logout:
        if st.button("Log Out", type="primary", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.username = ""
            st.session_state.role = ""
            st.rerun()

    # =====================================================
    # LOAD SHEET DATA
    # =====================================================

    @st.cache_data(ttl=10)
    def load_sheet_data(sheet_name):
        if not os.path.exists(excel_file_path):
            return pd.DataFrame()

        try:
            xls = pd.ExcelFile(excel_file_path)
            matched_sheet = None

            for name in xls.sheet_names:
                if name.strip().lower() == sheet_name.lower():
                    matched_sheet = name
                    break

            if matched_sheet:
                return pd.read_excel(xls, sheet_name=matched_sheet)

            return pd.DataFrame()
        except Exception:
            return pd.DataFrame()

    # =====================================================
    # SAVE DATA
    # =====================================================

    def save_sheet_data(df_to_save, target_name):
        if not os.path.exists(excel_file_path):
            return

        xls = pd.ExcelFile(excel_file_path)
        actual_name = target_name

        for name in xls.sheet_names:
            if name.strip().lower() == target_name.lower():
                actual_name = name
                break

        with pd.ExcelWriter(
            excel_file_path,
            mode="a",
            engine="openpyxl",
            if_sheet_exists="replace",
        ) as writer:
            df_to_save.to_excel(writer, sheet_name=actual_name, index=False)

    # =====================================================
    # LOAD DATA
    # =====================================================

    df_feeder = load_sheet_data("feeder")
    df_dist = load_sheet_data("distribution")

    all_projects = []
    if not df_feeder.empty:
        all_projects.append(df_feeder)
    if not df_dist.empty:
        all_projects.append(df_dist)

    combined_df = (
        pd.concat(all_projects, ignore_index=True)
        if all_projects
        else pd.DataFrame()
    )

    # =====================================================
    # PROJECT SUMMARY
    # =====================================================

    total_projects = len(combined_df)
    completed_projects = 0
    pending_projects = 0

    percentage_column = None

    if not combined_df.empty:
        for column in combined_df.columns:
            column_name = str(column).strip().lower()
            if any(
                term in column_name
                for term in ["progress", "completion", "%", "percent"]
            ):
                percentage_column = column
                break

    if percentage_column is not None:
        progress_values = (
            combined_df[percentage_column]
            .astype(str)
            .str.strip()
            .str.replace("%", "", regex=False)
        )
        progress_values = pd.to_numeric(progress_values, errors="coerce")

        completed_projects = int((progress_values >= 100).sum())
        pending_projects = int((progress_values < 100).sum())

    # =====================================================
    # SUMMARY SECTION
    # =====================================================

    st.markdown(
        """
        <h3 style="color: #234263; margin-top: 10px; margin-bottom: 15px;">
            Project Summary
        </h3>
        """,
        unsafe_allow_html=True,
    )

    card1, card2, card3 = st.columns(3)

    with card1:
        st.metric("Total Projects", total_projects)

    with card2:
        st.metric("Completed Projects", completed_projects)

    with card3:
        st.metric("Pending Projects", pending_projects)

    if os.path.exists(excel_file_path):
        last_updated = datetime.fromtimestamp(
            os.path.getmtime(excel_file_path)
        ).strftime("%d %B %Y")
    else:
        last_updated = "N/A"

    st.markdown(
        f"""
        <div style="text-align: center; font-size: 14px; color: #666666; margin-top: 5px; margin-bottom: 15px;">
            🕒 Last Updated: <b>{last_updated}</b>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        "<hr style='margin-top: 5px; margin-bottom: 15px;'>",
        unsafe_allow_html=True,
    )

    # =====================================================
    # TABS & TABLES
    # =====================================================

    tab_feeder, tab_distribution = st.tabs(
        ["📋 Feeder Data", "📋 Distribution Data"]
    )

    # --- FEEDER TAB ---
    with tab_feeder:
        if not df_feeder.empty:
            if search_query:
                mask = df_feeder.astype(str).apply(
                    lambda x: x.str.contains(
                        search_query, case=False, na=False
                    )
                ).any(axis=1)
                filtered_feeder = df_feeder[mask]
            else:
                filtered_feeder = df_feeder

            if st.session_state.role == "admin":
                edited_feeder = st.data_editor(
                    filtered_feeder,
                    num_rows="dynamic",
                    use_container_width=True,
                    hide_index=True,
                    key="edit_feeder",
                )
                if st.button("Save Feeder Changes"):
                    save_sheet_data(edited_feeder, "feeder")
                    st.success("Feeder sheet updated successfully!")
                    st.cache_data.clear()
                    st.rerun()
            else:
                st.table(filtered_feeder)
        else:
            st.info(
                "No data found matching the 'feeder' sheet name inside"
                " North As-Built Tracker.xlsx."
            )

    # --- DISTRIBUTION TAB ---
    with tab_distribution:
        if not df_dist.empty:
            if search_query:
                mask = df_dist.astype(str).apply(
                    lambda x: x.str.contains(
                        search_query, case=False, na=False
                    )
                ).any(axis=1)
                filtered_dist = df_dist[mask]
            else:
                filtered_dist = df_dist

            if st.session_state.role == "admin":
                edited_dist = st.data_editor(
                    filtered_dist,
                    num_rows="dynamic",
                    use_container_width=True,
                    hide_index=True,
                    key="edit_dist",
                )
                if st.button("Save Distribution Changes"):
                    save_sheet_data(edited_dist, "distribution")
                    st.success("Distribution sheet updated successfully!")
                    st.cache_data.clear()
                    st.rerun()
            else:
                st.table(filtered_dist)
        else:
            st.info(
                "No data found matching the 'distribution' sheet name inside"
                " North As-Built Tracker.xlsx."
            )
