import base64
import os
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import pandas as pd
import streamlit as st

# Set page layout to wide
st.set_page_config(page_title="As-Built Tracker North", layout="wide")

# Paths for files (relative for Streamlit Cloud deployment)
excel_file_path = "North As-Built Tracker.xlsx"
login_bg_path = "login_bg.png"


# --- FUNCTION TO SET BACKGROUND IMAGE FOR LOGIN ---
def set_bg_image(image_file):
    if os.path.exists(image_file):
        with open(image_file, "rb") as f:
            encoded_string = base64.b64encode(f.read()).decode()
        st.markdown(
            f"""
            <style>
            .stApp {{
                background-image: url("data:image/png;base64,{encoded_string}");
                background-size: cover;
                background-position: center;
                background-repeat: no-repeat;
                background-attachment: fixed;
            }}
            </style>
            """,
            unsafe_allow_html=True,
        )


# --- CUSTOM CSS STYLING ---
st.markdown(
    """
    <style>
    /* Hide 'Press Enter to submit form' hint text in login inputs */
    div[data-testid="InputInstructions"] {
        display: none !important;
    }
    
    /* Table Header Custom Styling */
    th {
        background-color: #234263 !important;
        color: white !important;
        font-weight: bold !important;
        text-align: center !important;
    }
    
    /* Form Subtitle Text Styling */
    .login-subtitle {
        color: #d1d5db !important;
        text-align: center;
        font-size: 1.1rem;
        margin-bottom: 1.5rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Authentication credentials
USERS = {"admin": "transworld123", "viewer": "transworldview"}

# Initialize session state for login
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
    st.session_state["username"] = ""


# --- EXCEL SAVE FUNCTION WITH BLUE HEADER STYLING ---
def save_sheet_data(df_to_save, target_name):
    if not os.path.exists(excel_file_path):
        st.error("Excel file not found!")
        return

    xls = pd.ExcelFile(excel_file_path)
    actual_name = target_name

    for name in xls.sheet_names:
        if name.strip().lower() == target_name.lower():
            actual_name = name
            break

    with pd.ExcelWriter(
        excel_file_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
        df_to_save.to_excel(writer, sheet_name=actual_name, index=False)

        # Apply Blue Header Formatting to Excel
        workbook = writer.book
        worksheet = workbook[actual_name]

        header_fill = PatternFill(
            start_color="4682B4", end_color="4682B4", fill_type="solid"
        )
        header_font = Font(color="FFFFFF", bold=True)

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")


# --- LOGIN PAGE ---
if not st.session_state["logged_in"]:
    set_bg_image(login_bg_path)

    col1, col2, col3 = st.columns([1, 1.2, 1])

    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        with st.form("login_form"):
            st.markdown(
                "<h1 style='text-align: center; color: white;'>As-Built Tracker North</h1>",
                unsafe_allow_html=True,
            )
            st.markdown(
                "<p class='login-subtitle'>Portal Login</p>",
                unsafe_allow_html=True,
            )

            username_input = st.text_input("Username")
            password_input = st.text_input("Password", type="password")

            submit = st.form_submit_button("Login", use_container_width=True)

            if submit:
                if (
                    username_input in USERS
                    and USERS[username_input] == password_input
                ):
                    st.session_state["logged_in"] = True
                    st.session_state["username"] = username_input
                    st.rerun()
                else:
                    st.error("Invalid Username or Password")

# --- MAIN DASHBOARD PAGE ---
else:
    # Sidebar Navigation & Logout
    st.sidebar.title(
        f"Welcome, {st.session_state['username'].capitalize()} 👋"
    )

    if st.sidebar.button("Logout", use_container_width=True):
        st.session_state["logged_in"] = False
        st.session_state["username"] = ""
        st.rerun()

    st.title("📊 As-Built Tracker North Dashboard")

    if os.path.exists(excel_file_path):
        xls = pd.ExcelFile(excel_file_path)
        sheet_names = xls.sheet_names

        selected_sheet = st.sidebar.selectbox("Select Sheet / Region", sheet_names)

        if selected_sheet:
            df = pd.read_excel(excel_file_path, sheet_name=selected_sheet)

            # Search Filter by Project Name
            search_query = st.text_input(
                "🔍 Search by Project", placeholder="Type project name here..."
            )

            if search_query:
                # Find matching column for project searching
                proj_col = [
                    c for c in df.columns if "project" in str(c).lower()
                ]
                if proj_col:
                    df_filtered = df[
                        df[proj_col[0]]
                        .astype(str)
                        .str.contains(search_query, case=False, na=False)
                    ]
                else:
                    df_filtered = df
            else:
                df_filtered = df

            # Table View / Admin Data Editor
            if st.session_state["username"] == "admin":
                st.subheader(f"Edit Data - {selected_sheet}")
                edited_df = st.data_editor(
                    df_filtered, num_rows="dynamic", hide_index=True
                )

                if st.button("Save Changes"):
                    save_sheet_data(edited_df, selected_sheet)
                    st.success("Changes saved successfully!")
            else:
                st.subheader(f"View Data - {selected_sheet}")
                st.dataframe(df_filtered, hide_index=True, use_container_width=True)
    else:
        st.error(
            f"Excel file '{excel_file_path}' not found! Please check the repository files."
        )