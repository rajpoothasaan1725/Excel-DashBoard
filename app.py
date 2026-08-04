import sys
import os
import asyncio
from datetime import datetime
import pandas as pd
import streamlit as st

# Windows Asyncio Connection Reset Error Suppression Fix
if sys.platform == 'win32':
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    except Exception:
        pass

st.set_page_config(
    page_title="GIS Mapping Services - Geo Academy Tracker",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    /* Global Styling */
    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }
    
    /* Centering helpers for Login */
    div[data-testid="stImage"] {
        display: flex;
        justify-content: center;
        align-items: center;
    }
    .login-title {
        text-align: center;
        margin: 5px 0 2px 0;
        font-size: 26px;
        font-weight: 700;
        color: #1e293b;
    }
    .login-sub {
        text-align: center;
        margin: 0 0 15px 0;
        font-size: 14px;
        color: #64748b;
    }

    /* Company Header Card */
    .header-card {
        background: linear-gradient(135deg, #0f2027, #203a43, #2c5364);
        padding: 22px 28px;
        border-radius: 14px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.12);
        display: flex;
        align-items: center;
        gap: 20px;
    }
    .header-card h1 {
        margin: 0;
        font-size: 26px;
        font-weight: 700;
        color: #ffffff;
    }
    .header-card p {
        margin: 4px 0 0 0;
        color: #cbd5e1;
        font-size: 14px;
    }

    /* Financial Alert Banners */
    .loss-banner {
        background-color: #fef2f2;
        border-left: 6px solid #ef4444;
        color: #991b1b;
        padding: 16px 20px;
        border-radius: 8px;
        margin-bottom: 20px;
        font-size: 15px;
        box-shadow: 0 2px 8px rgba(239, 68, 68, 0.08);
    }
    .profit-banner {
        background-color: #f0fdf4;
        border-left: 6px solid #22c55e;
        color: #166534;
        padding: 16px 20px;
        border-radius: 8px;
        margin-bottom: 20px;
        font-size: 15px;
        box-shadow: 0 2px 8px rgba(34, 197, 94, 0.08);
    }

    /* Metric Card Styling */
    div[data-testid="stMetricValue"] {
        font-size: 24px !important;
        font-weight: 700 !important;
    }
    
    /* Login Box */
    .login-container {
        max-width: 420px;
        margin: 40px auto;
        padding: 30px;
        background-color: #ffffff;
        border-radius: 12px;
        box-shadow: 0 6px 20px rgba(0,0,0,0.08);
        border: 1px solid #e2e8f0;
    }
</style>
""", unsafe_allow_html=True)

EXCEL_FILE = "gis_academy_tracker.xlsx"

def clean_numeric_val(val):
    """Clean and convert text values like '10K', '5k', '20,000' to floating numbers."""
    if pd.isna(val) or val is None:
        return 0.0
    val_str = str(val).strip().upper().replace(",", "").replace("RS", "").replace("PKR", "").replace("$", "")
    if not val_str or val_str == "NONE" or val_str == "NAN":
        return 0.0
    try:
        if val_str.endswith("K"):
            return float(val_str[:-1].strip()) * 1000.0
        elif val_str.endswith("M"):
            return float(val_str[:-1].strip()) * 1000000.0
        else:
            return float(val_str)
    except ValueError:
        return 0.0

def init_excel():
    """Create default excel sheet if not already existing."""
    if not os.path.exists(EXCEL_FILE):
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            # 1. Students Sheet
            df_std = pd.DataFrame({
                "S.No": [1, 2],
                "Name": ["Sample Student 1", "Sample Student 2"],
                "Total Fee (Pkr)": ["15K", "20K"],
                "Country": ["Pakistan", "USA"],
                "Fee Received (Pkr)": ["10K", "20K"],
                "Fee Pending (Pkr)": ["5K", "0"],
                "Comments": ["Paid Partial", "Full Paid"],
                "Month": ["Aug-26", "Aug-26"]
            })
            df_std.to_excel(writer, sheet_name="Students", index=False)

            # 2. Expenses & Investment Sheet
            df_exp = pd.DataFrame({
                "Date": [datetime.now().strftime("%Y-%m-%d"), datetime.now().strftime("%Y-%m-%d")],
                "Type": ["Investment", "Expense"],
                "Category / Description": ["Initial Capital", "Domain & Hosting"],
                "Amount (Pkr)": ["50K", "10K"],
                "Added By": ["admin1", "admin1"],
                "Notes": ["Initial pool", "Annual domain bill"]
            })
            df_exp.to_excel(writer, sheet_name="Expenses & Investment", index=False)

            # 3. Projects Sheet
            df_proj = pd.DataFrame({
                "Project ID": ["PRJ-001"],
                "Project Name": ["GIS Mapping System"],
                "Client": ["Geo Corp"],
                "Budget (Pkr)": ["100K"],
                "Status": ["In Progress"],
                "Deadline": ["2026-09-30"]
            })
            df_proj.to_excel(writer, sheet_name="Projects", index=False)

            # 4. Publications Sheet
            df_pub = pd.DataFrame({
                "Title": ["Remote Sensing in GIS"],
                "Authors": ["Geo Research Team"],
                "Journal/Conference": ["International GIS Journal"],
                "Status": ["Published"],
                "Year": [2026]
            })
            df_pub.to_excel(writer, sheet_name="Publications", index=False)

            # 5. Physical Academy Sheet
            df_phy = pd.DataFrame({
                "Module / Plan": ["Physical Classroom Setup", "Lab Computers"],
                "Target Date": ["Coming Soon (Q4 2026)", "Coming Soon"],
                "Estimated Cost (Pkr)": ["200K", "150K"],
                "Status": ["Planned", "Planned"]
            })
            df_phy.to_excel(writer, sheet_name="Physical Academy", index=False)

init_excel()

def read_sheet(sheet_name):
    """Read a specific sheet from Excel safely."""
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()

def get_all_sheet_names():
    """Get list of all sheet names present in the Excel file."""
    try:
        if os.path.exists(EXCEL_FILE):
            xl = pd.ExcelFile(EXCEL_FILE)
            return xl.sheet_names
    except Exception:
        pass
    return ["Students", "Expenses & Investment", "Projects", "Publications", "Physical Academy"]

def write_sheet(df, sheet_name):
    """Write DataFrame to a specific sheet without overwriting other sheets."""
    try:
        excel_dict = pd.read_excel(EXCEL_FILE, sheet_name=None)
        excel_dict[sheet_name] = df
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            for sname, sdata in excel_dict.items():
                sdata.to_excel(writer, sheet_name=sname, index=False)
        return True
    except PermissionError:
        st.error("⚠️ **File Permission Error**: 'gis_academy_tracker.xlsx' is currently open in Microsoft Excel or another program. Please **close the Excel file** on your PC and try saving again!")
        return False
    except Exception as e:
        st.error(f"Error saving to Excel: {e}")
        return False

ADMIN_CREDENTIALS = {
    "admin1": "geo123",
    "admin2": "gis456"
}

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "user" not in st.session_state:
    st.session_state["user"] = ""

if not st.session_state["logged_in"]:
    # Custom CSS for GIS Satellite Glassmorphism Login Screen
    st.markdown("""
        <style>
            [data-testid="stSidebar"] { display: none; }
            header { display: none !important; }
            footer { display: none !important; }

            /* Full screen background */
            .stApp {
                background: linear-gradient(rgba(10, 15, 26, 0.70), rgba(10, 15, 26, 0.80)), 
                            url('https://images.unsplash.com/photo-1524661135-423995f22d0b?q=80&w=1920&auto=format&fit=crop') no-repeat center center fixed;
                background-size: cover;
            }

            .main .block-container { 
                max-width: 460px !important; 
                padding-top: 10vh !important; 
                padding-bottom: 0rem !important; 
                margin: 0 auto !important;
            }
            body { 
                overflow: hidden !important; 
            }

            /* Single Centered Glassmorphism Card Container */
            div[data-testid="stForm"] {
                max-width: 460px !important;
                margin: 0 auto !important;
                background: rgba(15, 23, 42, 0.88) !important;
                backdrop-filter: blur(16px) !important;
                -webkit-backdrop-filter: blur(16px) !important;
                border: 1px solid rgba(255, 255, 255, 0.18) !important;
                border-radius: 16px !important;
                padding: 24px 26px 20px 26px !important;
                box-shadow: 0 15px 35px rgba(0, 0, 0, 0.7), 0 0 20px rgba(0, 122, 255, 0.2) !important;
            }

            /* Labels */
            .stTextInput label p {
                color: #f8fafc !important;
                font-weight: 600 !important;
                font-size: 13px !important;
            }

            /* Input Fields - Compact white rounded style */
            .stTextInput input {
                background-color: #ffffff !important;
                color: #0f172a !important;
                border-radius: 8px !important;
                padding: 8px 12px !important;
                border: 1px solid #cbd5e1 !important;
                font-weight: 500 !important;
                font-size: 14px !important;
            }
            .stTextInput input:focus {
                border-color: #007bff !important;
                box-shadow: 0 0 0 3px rgba(0, 123, 255, 0.3) !important;
            }

            /* Action Button Styling */
            div[data-testid="stFormSubmitButton"] > button {
                background: linear-gradient(135deg, #007bff, #0056b3) !important;
                color: #ffffff !important;
                font-weight: 700 !important;
                font-size: 14px !important;
                border: none !important;
                border-radius: 8px !important;
                padding: 8px 0 !important;
                margin-top: 6px !important;
                box-shadow: 0 4px 12px rgba(0, 123, 255, 0.4) !important;
                transition: all 0.3s ease !important;
            }
            div[data-testid="stFormSubmitButton"] > button:hover {
                background: linear-gradient(135deg, #0056b3, #004085) !important;
                box-shadow: 0 6px 18px rgba(0, 123, 255, 0.6) !important;
                transform: translateY(-1px);
            }
        </style>
    """, unsafe_allow_html=True)
    
    with st.form("login_form"):
        # Highlighted Company Header Inside Card
        st.markdown("""
            <div style='text-align: center; margin-bottom: 16px;'>
                <div style='background: linear-gradient(135deg, #38bdf8, #60a5fa); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-size: 21px; font-weight: 800; letter-spacing: 0.5px;'>
                    GIS MAPPING SERVICES
                </div>
                <div style='color: #94a3b8; font-size: 12px; margin-top: 2px; font-weight: 500;'>
                    Geo Academy Portal Login
                </div>
            </div>
        """, unsafe_allow_html=True)

        username_input = st.text_input("Username", placeholder="Enter admin1 or admin2")
        password_input = st.text_input("Password", type="password", placeholder="Enter password")
        submit_login = st.form_submit_button("Login", use_container_width=True)
        
        if submit_login:
            user = username_input.strip()
            pwd = password_input.strip()
            if user in ADMIN_CREDENTIALS and ADMIN_CREDENTIALS[user] == pwd:
                st.session_state["logged_in"] = True
                st.session_state["user"] = user
                st.success("Login Successful! Redirecting...")
                st.rerun()
            else:
                st.error("❌ Invalid Username or Password")

    st.markdown("""
        <div style='max-width: 460px; margin: 12px auto 0 auto; background-color: rgba(15, 23, 42, 0.85); border: 1px solid rgba(255, 255, 255, 0.15); padding: 8px 12px; border-radius: 8px; font-size: 12px; color: #cbd5e1; text-align: center; backdrop-filter: blur(8px);'>
            📌 <strong>Admin 1:</strong> <code>admin1</code> / <code>geo123</code> &nbsp;|&nbsp; 
            <strong>Admin 2:</strong> <code>admin2</code> / <code>gis456</code>
        </div>
    """, unsafe_allow_html=True)
    st.stop()

header_col1, header_col2 = st.columns([0.15, 0.85])

with header_col1:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=90)
    else:
        st.title("🗺️")

with header_col2:
    st.markdown("""
        <div style='padding-top: 5px;'>
            <h1 style='margin:0; font-size: 28px; color: #1e293b;'>GIS Mapping Services</h1>
            <p style='margin:0; font-size: 15px; color: #64748b;'>Geo Academy - Management & Financial Tracker</p>
        </div>
    """, unsafe_allow_html=True)

st.markdown("<hr style='margin: 15px 0 25px 0;'>", unsafe_allow_html=True)

with st.sidebar:
    st.subheader("⚙️ Settings & Backup")
    
    if st.button("🔄 Refresh Data from Excel", use_container_width=True, type="primary"):
        st.rerun()

    st.markdown("---")
    
    # Upload Logo Option
    uploaded_logo = st.file_uploader("🖼️ Upload / Change Logo", type=["png", "jpg", "jpeg"])
    if uploaded_logo:
        with open("logo.png", "wb") as f:
            f.write(uploaded_logo.getbuffer())
        st.success("Logo updated successfully!")
        st.rerun()

    st.markdown("---")
    
    # Excel Download Button
    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            st.download_button(
                label="📥 Download Excel Backup",
                data=f,
                file_name=f"GIS_Academy_Tracker_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    # Replace Excel File
    uploaded_excel = st.file_uploader("📤 Replace Excel File", type=["xlsx"])
    if uploaded_excel:
        with open(EXCEL_FILE, "wb") as f:
            f.write(uploaded_excel.getbuffer())
        st.success("Excel data replaced!")
        st.rerun()
        
    st.markdown("---")
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state["logged_in"] = False
        st.session_state["user"] = ""
        st.rerun()

# =========================================================
# GLOBAL FINANCIAL CALCULATIONS & TOP MAIN DASHBOARD METRICS
# =========================================================
df_std_global = read_sheet("Students")
df_exp_global = read_sheet("Expenses & Investment")
df_proj_global = read_sheet("Projects")

# 1. Student Revenue Calculation
std_rec_col = None
for c in df_std_global.columns:
    if "received" in c.lower():
        std_rec_col = c
        break
total_std_rev = df_std_global[std_rec_col].apply(clean_numeric_val).sum() if not df_std_global.empty and std_rec_col else 0.0

# 2. Project Revenue Calculation
proj_budg_col = None
for c in df_proj_global.columns:
    if any(k in c.lower() for k in ["budget", "total budget", "amount", "price"]):
        proj_budg_col = c
        break
total_proj_rev = df_proj_global[proj_budg_col].apply(clean_numeric_val).sum() if not df_proj_global.empty and proj_budg_col else 0.0

total_gross_rev = total_std_rev + total_proj_rev

# 3. Investments & Expenses Calculation
amt_col = None
for c in df_exp_global.columns:
    if "amount" in c.lower():
        amt_col = c
        break

total_investment = 0.0
total_expenses = 0.0

if not df_exp_global.empty and amt_col:
    df_exp_calc = df_exp_global.copy()
    df_exp_calc["_numeric_amt"] = df_exp_calc[amt_col].apply(clean_numeric_val)
    type_col = [c for c in df_exp_calc.columns if "type" in c.lower()]
    if type_col:
        is_invest = df_exp_calc[type_col[0]].astype(str).str.lower().str.contains("invest")
        total_investment = df_exp_calc[is_invest]["_numeric_amt"].sum()
        total_expenses = df_exp_calc[~is_invest]["_numeric_amt"].sum()
    else:
        total_expenses = df_exp_calc["_numeric_amt"].sum()

net_cash_balance = (total_gross_rev + total_investment) - total_expenses

# ---------------------------------------------------------
# MAIN OVERALL FINANCIAL BANNER & KPI METRICS (TOP DISPLAY)
# ---------------------------------------------------------
if net_cash_balance < 0:
    st.markdown(f"""
        <div class='loss-banner'>
            ⚠️ <strong>Financial Warning: Overall Business in LOSS / Deficit</strong><br>
            Total Expenses (PKR {total_expenses:,.0f}) exceed Gross Revenue + Investments (PKR {(total_gross_rev + total_investment):,.0f}). Net Deficit: <strong>PKR {abs(net_cash_balance):,.0f}</strong>
        </div>
    """, unsafe_allow_html=True)
else:
    st.markdown(f"""
        <div class='profit-banner'>
            ✅ <strong>Financial Status: POSITIVE Cash Flow / Profit</strong><br>
            Net Cash Balance in Hand: <strong>PKR {net_cash_balance:,.0f}</strong>
        </div>
    """, unsafe_allow_html=True)

# Display 5 Key Metrics at the top of the main page
m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("🎓 Student Revenue", f"PKR {total_std_rev:,.0f}")
m2.metric("📁 Project Earnings", f"PKR {total_proj_rev:,.0f}")
m3.metric("💵 Total Investment", f"PKR {total_investment:,.0f}")
m4.metric("💸 Total Expenses", f"PKR {total_expenses:,.0f}")
m5.metric("⚖️ Net Cash Balance", f"PKR {net_cash_balance:,.0f}")

st.markdown("<br>", unsafe_allow_html=True)

# Dynamically fetch all sheet names present in the Excel file
all_sheet_names = get_all_sheet_names()

# Dynamic Emoji mapping for tabs
tab_icons = {
    "Students": "🎓 ",
    "Expenses & Investment": "💰 ",
    "Projects": "📁 ",
    "Publications": "📚 ",
    "Physical Academy": "🚀 "
}

tab_titles = [f"{tab_icons.get(s, '📄 ')}{s}" for s in all_sheet_names]
tabs = st.tabs(tab_titles)

for tab, sheet_name in zip(tabs, all_sheet_names):
    with tab:
        df_sheet = read_sheet(sheet_name)
        
        # 1. Custom Specialized UI for "Students" Sheet
        if sheet_name == "Students":
            st.subheader("🎓 Online Students Fee Tracker")
            with st.expander("➕ Add New Student Entry"):
                with st.form("add_student_form"):
                    c1, c2, c3 = st.columns(3)
                    s_name = c1.text_input("Student Name")
                    s_total_fee = c2.text_input("Total Fee (PKR)", value="15K")
                    s_received = c3.text_input("Fee Received (PKR)", value="10K")
                    
                    c4, c5, c6 = st.columns(3)
                    s_country = c4.text_input("Country", value="Pakistan")
                    s_month = c5.text_input("Month / Batch", value=datetime.now().strftime("%b-%y"))
                    s_comments = c6.text_input("Comments", value="Paid Partial")

                    if st.form_submit_button("Save Student to Excel"):
                        if s_name:
                            tot_num = clean_numeric_val(s_total_fee)
                            rec_num = clean_numeric_val(s_received)
                            pending_num = max(0.0, tot_num - rec_num)
                            
                            new_row = {
                                "S.No": len(df_sheet) + 1,
                                "Name": s_name,
                                "Total Fee (Pkr)": s_total_fee,
                                "Country": s_country,
                                "Fee Received (Pkr)": s_received,
                                "Fee Pending (Pkr)": f"{pending_num:,.0f}" if pending_num > 0 else "0",
                                "Comments": s_comments,
                                "Month": s_month
                            }
                            df_sheet = pd.concat([df_sheet, pd.DataFrame([new_row])], ignore_index=True)
                            if write_sheet(df_sheet, "Students"):
                                st.success(f"Student '{s_name}' added successfully!")
                                st.rerun()

            st.markdown("### 📋 Student List Editor")
            edited_df = st.data_editor(df_sheet, num_rows="dynamic", use_container_width=True, key=f"editor_{sheet_name}")

            if st.button("💾 Save All Student Changes to Excel", type="primary", key=f"btn_save_{sheet_name}"):
                for idx, row in edited_df.iterrows():
                    tot_col = [c for c in edited_df.columns if "total" in c.lower()]
                    rec_col = [c for c in edited_df.columns if "received" in c.lower()]
                    pend_col = [c for c in edited_df.columns if "pending" in c.lower()]
                    
                    if tot_col and rec_col and pend_col:
                        tot_v = clean_numeric_val(row[tot_col[0]])
                        rec_v = clean_numeric_val(row[rec_col[0]])
                        edited_df.at[idx, pend_col[0]] = f"{max(0.0, tot_v - rec_v):,.0f}"

                if write_sheet(edited_df, "Students"):
                    st.success("Student records updated successfully in Excel!")
                    st.rerun()

        # 2. Custom Specialized UI for "Expenses & Investment" Sheet
        elif sheet_name == "Expenses & Investment":
            st.subheader("💰 Expenses & Investment Tracker")
            with st.expander("➕ Add New Expense or Investment"):
                with st.form("add_exp_form"):
                    e1, e2, e3 = st.columns(3)
                    e_type = e1.selectbox("Type", ["Expense", "Investment", "Salary", "Website & Domain", "Marketing", "Other"])
                    e_desc = e2.text_input("Category / Description", value="Marketing Ads")
                    e_amt = e3.text_input("Amount (PKR)", value="10K")
                    
                    e4, e5 = st.columns(2)
                    e_date = e4.date_input("Date", value=datetime.now())
                    e_notes = e5.text_input("Notes", value="Spent for campaign")

                    if st.form_submit_button("Record Financial Entry"):
                        new_exp = {
                            "Date": e_date.strftime("%Y-%m-%d"),
                            "Type": e_type,
                            "Category / Description": e_desc,
                            "Amount (Pkr)": e_amt,
                            "Added By": st.session_state["user"],
                            "Notes": e_notes
                        }
                        df_exp_save = df_sheet.drop(columns=["_numeric_amt"], errors="ignore")
                        df_exp_save = pd.concat([df_exp_save, pd.DataFrame([new_exp])], ignore_index=True)
                        if write_sheet(df_exp_save, "Expenses & Investment"):
                            st.success("Financial entry saved successfully!")
                            st.rerun()

            st.markdown("### 📋 Expenses & Investment Grid Editor")
            df_exp_clean = df_sheet.drop(columns=["_numeric_amt"], errors="ignore")
            edited_df = st.data_editor(df_exp_clean, num_rows="dynamic", use_container_width=True, key=f"editor_{sheet_name}")

            if st.button("💾 Save All Financial Changes to Excel", type="primary", key=f"btn_save_{sheet_name}"):
                if write_sheet(edited_df, "Expenses & Investment"):
                    st.success("Expenses and Investment records updated successfully!")
                    st.rerun()

        # 3. Custom Specialized UI for "Projects" Sheet
        elif sheet_name == "Projects":
            st.subheader("📁 GIS Projects Tracker")
            with st.expander("➕ Add New Project"):
                with st.form("add_project_form"):
                    p1, p2 = st.columns(2)
                    p_id = p1.text_input("Project ID", value=f"PRJ-00{len(df_sheet)+1}")
                    p_name = p2.text_input("Project Name")
                    p3, p4 = st.columns(2)
                    p_client = p3.text_input("Client Name")
                    p_budget = p4.text_input("Budget (PKR)", value="50K")
                    p5, p6 = st.columns(2)
                    p_status = p5.selectbox("Status", ["Planning", "In Progress", "Completed", "On Hold"])
                    p_deadline = p6.date_input("Deadline", value=datetime.now())

                    if st.form_submit_button("Save Project"):
                        if p_name:
                            new_p = {
                                "Project ID": p_id,
                                "Project Name": p_name,
                                "Client": p_client,
                                "Budget (Pkr)": p_budget,
                                "Status": p_status,
                                "Deadline": p_deadline.strftime("%Y-%m-%d")
                            }
                            df_sheet = pd.concat([df_sheet, pd.DataFrame([new_p])], ignore_index=True)
                            if write_sheet(df_sheet, "Projects"):
                                st.success("Project added successfully!")
                                st.rerun()

            st.markdown("### 📋 Projects Grid Editor")
            edited_df = st.data_editor(df_sheet, num_rows="dynamic", use_container_width=True, key=f"editor_{sheet_name}")
            if st.button("💾 Save Projects Changes to Excel", type="primary", key=f"btn_save_{sheet_name}"):
                if write_sheet(edited_df, "Projects"):
                    st.success("Projects sheet updated!")
                    st.rerun()

        # 4. Custom Specialized UI for "Publications" Sheet
        elif sheet_name == "Publications":
            st.subheader("📚 Research & Publications Tracker")
            with st.expander("➕ Add Publication"):
                with st.form("add_pub_form"):
                    pu1 = st.text_input("Publication Title")
                    pu2 = st.text_input("Authors")
                    pu3, pu4, pu5 = st.columns(3)
                    pu_journal = pu3.text_input("Journal / Conference")
                    pu_status = pu4.selectbox("Status", ["Draft", "Under Review", "Accepted", "Published"])
                    pu_year = pu5.number_input("Year", value=2026, step=1)

                    if st.form_submit_button("Save Publication"):
                        if pu1:
                            new_pub = {
                                "Title": pu1,
                                "Authors": pu2,
                                "Journal/Conference": pu_journal,
                                "Status": pu_status,
                                "Year": int(pu_year)
                            }
                            df_sheet = pd.concat([df_sheet, pd.DataFrame([new_pub])], ignore_index=True)
                            if write_sheet(df_sheet, "Publications"):
                                st.success("Publication record added!")
                                st.rerun()

            st.markdown("### 📋 Publications Grid Editor")
            edited_df = st.data_editor(df_sheet, num_rows="dynamic", use_container_width=True, key=f"editor_{sheet_name}")
            if st.button("💾 Save Publications Changes to Excel", type="primary", key=f"btn_save_{sheet_name}"):
                if write_sheet(edited_df, "Publications"):
                    st.success("Publications sheet updated!")
                    st.rerun()

        # 5. Dynamic UI for ANY New / Custom Sheet added in Excel
        else:
            st.subheader(f"📄 Sheet: {sheet_name}")
            st.info(f"📌 Custom Excel Sheet detected. You can edit, add rows, or delete entries directly below.")
            
            edited_df = st.data_editor(
                df_sheet,
                num_rows="dynamic",
                use_container_width=True,
                key=f"editor_{sheet_name}"
            )
            
            if st.button(f"💾 Save '{sheet_name}' Changes to Excel", type="primary", key=f"btn_save_{sheet_name}"):
                if write_sheet(edited_df, sheet_name):
                    st.success(f"Changes saved to sheet '{sheet_name}' successfully!")
                    st.rerun()