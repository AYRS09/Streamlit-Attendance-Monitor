# =============================
# 🚀 Employee Attendance Dashboard - Streamlit App
# =============================

# === Imports ===
import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
import io
import smtplib
from email.message import EmailMessage
from PIL import Image
from datetime import datetime
import os

# =============================
# 🌗 Theme Toggle
# =============================
theme = st.sidebar.radio("🌓 Choose Theme", ["Light", "Dark"])
if theme == "Dark":
    st.markdown("""
        <style>
        body { background-color: #0E1117; color: white; }
        .stApp { background-color: #0E1117; color: white; }
        </style>
    """, unsafe_allow_html=True)

# =============================
# ⚙️ Streamlit Config
# =============================
st.set_page_config(page_title="Employee Attendance Dashboard", layout="wide")

# =============================
# 🖼️ Load Logo
# =============================
image_path = os.path.join(os.path.dirname(__file__), "download.jpeg")
if os.path.exists(image_path):
    st.sidebar.image(image_path, width=120)
    st.sidebar.markdown("### 👋 Welcome to the Dashboard")
else:
    st.sidebar.warning("⚠️ Logo image not found.")

# =============================
# ⏱️ Last Updated Timestamp
# =============================
from datetime import datetime
import pytz  # Add this import at the top of your script

# Get current time in IST
ist = pytz.timezone('Asia/Kolkata')
now = datetime.now(ist).strftime("%d %b %Y, %I:%M %p")

# Display Last Updated timestamp
st.markdown(
    f"<div style='text-align:right; color:gray; font-size:0.85rem;'>🕒 Last updated: {now}</div>",
    unsafe_allow_html=True
)

# =============================
# 🧭 Title & Style
# =============================
st.markdown("<h1 style='text-align: center; color: #4B8BBE;'>|🚀 Employee Productivity Dashboard 🚀| Diverse Infotech Pvt Ltd</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; color: gray;'>Punctuality & Productivity Analysis Based on Daily Hours Worked</h4>", unsafe_allow_html=True)

# === Custom Dark Theme Style ===
st.markdown("""
    <style>
    .stApp { background-color: #000; color: #E0E0E0; font-family: 'Segoe UI'; }
    section[data-testid="stSidebar"] { background-color: #111; color: white; }
    h1, h2, h3, h4 { color: #00CED1; font-weight: 600; }
    .element-container { background-color: #1a1a1a; padding: 20px; border-radius: 12px; margin-bottom: 20px; box-shadow: 0 0 10px rgba(0,255,255,0.05); }
    .stButton>button, .stDownloadButton>button { border-radius: 8px; font-weight: 600; }
    .stDownloadButton>button { background-color: #FFD700; color: black; }
    .stButton>button { background-color: #00CED1; color: black; }
    .stTextInput>div>div>input, .stSelectbox>div>div>div>div { background-color: #1e1e1e; color: white; }
    </style>
""", unsafe_allow_html=True)

# =============================
# 📤 File Upload & Instructions
# =============================
st.sidebar.markdown("---")
st.sidebar.subheader("📄 Upload Attendance Sheet")
file = st.sidebar.file_uploader("Upload Excel/CSV File", type=["xlsx", "xls", "csv"])

st.sidebar.markdown("### 📝 Format Instructions")
st.sidebar.info("""
Your Excel/CSV file must contain these columns:
- `employee_id`
- `employee_gender`
- `employee_resident`
- `employee_department`
- `in_1`, `out_1`, `in_2`, `out_2`

🛑 Time format: `HH:MM AM/PM`
""")
st.sidebar.markdown("---")

# =============================
# 📦 Data Preprocessing
# =============================
if file is not None:
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    elif file.name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(file)
    else:
        st.error("❌ Please upload a valid .csv or Excel file.")
        st.stop()
else:
    st.warning("📂 Please upload your attendance file to proceed.")
    st.stop()

# --- Calculate hours from In/Out columns ---
in_cols = [col for col in df.columns if col.startswith('in_')]
out_cols = [col for col in df.columns if col.startswith('out_')]

for in_col, out_col in zip(in_cols, out_cols):
    hours_col = in_col.replace('in_', 'hours_')
    df[hours_col] = (
        pd.to_datetime(df[out_col], format='%I:%M %p', errors='coerce') -
        pd.to_datetime(df[in_col], format='%I:%M %p', errors='coerce')
    ).dt.total_seconds() / 3600
    df[hours_col] = df[hours_col].round(2)

# =============================
# 🧹 Clean & Prepare Data
# =============================

# --- Handle Duplicates ---
df.drop_duplicates(inplace=True)
duplicate_ids = df['employee_id'].value_counts()
duplicate_ids = duplicate_ids[duplicate_ids > 1]

if not duplicate_ids.empty:
    st.warning("⚠️ Found duplicate entries for these employee IDs:")
    st.dataframe(df[df['employee_id'].isin(duplicate_ids.index)])

df['total_hours'] = df[[col for col in df.columns if col.startswith('hours_')]].sum(axis=1)
df = df.sort_values('total_hours', ascending=False).drop_duplicates(subset=['employee_id'], keep='first')
df.drop(columns='total_hours', inplace=True)

# =============================
# 🔄 Reshape Data
# =============================
# 📅 Select Attendance Start Date
st.sidebar.markdown("### 📅 Select Attendance Start Date")
start_date_input = st.sidebar.date_input("Start Date", value=pd.to_datetime("2025-06-01"))
day_cols = sorted([col for col in df.columns if col.startswith('hours_')], key=lambda x: int(x.split('_')[1]))
df_long = df.melt(
    id_vars=['employee_id', 'employee_gender', 'employee_resident', 'employee_department'],
    value_vars=day_cols,
    var_name='day',
    value_name='hours_worked'
)
df_long['day_num'] = df_long['day'].str.extract(r'(\d+)').astype(int)
df_long['date'] = pd.to_datetime(start_date_input) + pd.to_timedelta(df_long['day_num'] - 1, unit='D')
df_long['is_punctual'] = df_long['hours_worked'] >= 8

# =============================
# 🎛️ Sidebar Filters
# =============================
st.sidebar.header("🔍 Filter Options")

# Employee filter
employees = sorted(df_long['employee_id'].dropna().unique())
selected_employees = st.sidebar.selectbox("👤 Select Employee", options=["All"] + employees)

# Residency filter
residency = st.sidebar.selectbox("🏩 Resident Type", options=["All", "Local", "Non-local"])

# Department filter
departments = sorted(df_long['employee_department'].dropna().unique())
selected_departments = st.sidebar.multiselect("🏢 Select Department(s)", options=departments, default=departments)

# Extended Date Range filter
min_date = pd.to_datetime("2020-01-01")
max_date = pd.to_datetime("2030-12-31")
default_start = df_long['date'].min()
default_end = df_long['date'].max()

date_range = st.sidebar.date_input(
    "🗓️ Select Date Range",
    [default_start, default_end],
    min_value=min_date,
    max_value=max_date
)

# --- Apply Filters ---
filtered_df = df_long[
    (df_long['date'] >= pd.to_datetime(date_range[0])) &
    (df_long['date'] <= pd.to_datetime(date_range[1]))
].copy()

if selected_employees != "All":
    filtered_df = filtered_df[filtered_df['employee_id'] == selected_employees]
if residency != "All":
    filtered_df = filtered_df[filtered_df['employee_resident'].str.lower() == residency.lower()]
if selected_departments:
    filtered_df = filtered_df[filtered_df['employee_department'].isin(selected_departments)]

# =============================
# 📊 KPIs
# =============================
total_employees = filtered_df['employee_id'].nunique()
total_days = len(filtered_df)
total_punctual = filtered_df[filtered_df['is_punctual']].shape[0]
avg_hours = round(filtered_df['hours_worked'].mean(), 2)
punctuality_rate = round((total_punctual / total_days) * 100, 2) if total_days else 0.0

st.markdown("<h2 style='text-align: center; color: white;'>📊 Key Metrics</h2>", unsafe_allow_html=True)
kpi1, kpi2, kpi3 = st.columns(3)
kpi1.metric("👥 Total Employees", total_employees)
kpi2.metric("✅ Punctuality Rate", f"{punctuality_rate}%")
kpi3.metric("⏱️ Average Hours Worked", f"{avg_hours} hrs")
st.markdown("---")

# =============================
# 🧭 Tabs: Visualization | Summary | Download | Email
# =============================
tab1, tab2, tab3, tab4 = st.tabs(["📊 Visualizations", "📋 Summary", "📅 Download", "📬 Email Summary"])

# --- Tab 1: Visualizations ---
with tab1:
    # Centered title
    st.markdown("<h2 style='text-align: center; color: white;'>📊 Employee Attendance Visualizations</h2>", unsafe_allow_html=True)

    row1_col1, row1_col2 = st.columns(2)
    with row1_col1:
        st.subheader("⏱️ Total Hours Worked per Employee")
        fig1 = px.bar(
            filtered_df.groupby('employee_id')['hours_worked'].sum().reset_index(),
            x='employee_id', y='hours_worked', color='hours_worked', color_continuous_scale='Greens')
        st.plotly_chart(fig1, use_container_width=True)

    with row1_col2:
        st.subheader("⏰ Punctuality Ratio per Employee")
        punctual_summary = filtered_df.groupby(['employee_id', 'is_punctual']).size().reset_index(name='Count')
        punctual_summary['Status'] = punctual_summary['is_punctual'].map({True: 'Met (≥8 hrs)', False: 'Not Met (<8 hrs)'})
        fig2 = px.bar(punctual_summary, x='employee_id', y='Count', color='Status', barmode='stack')
        st.plotly_chart(fig2, use_container_width=True)

    row2_col1, row2_col2 = st.columns(2)
    with row2_col1:
        st.subheader("🗓️ Daily Productivity Heatmap")
        heatmap_data = filtered_df.pivot_table(index='employee_id', columns='day_num', values='hours_worked')
        fig3 = px.imshow(heatmap_data.astype(np.float32), aspect="auto", color_continuous_scale='YlGnBu')
        st.plotly_chart(fig3, use_container_width=True)

    with row2_col2:
        st.subheader("📈 Overall Punctuality Ratio")
        overall = filtered_df['is_punctual'].value_counts().rename({True: 'Met ≥8 hrs', False: 'Not Met <8 hrs'})
        fig4 = px.pie(names=overall.index, values=overall.values)
        st.plotly_chart(fig4, use_container_width=True)

    row3_col1, row3_col2 = st.columns(2)
    with row3_col1:
        st.subheader("🏘️ Resident Type vs Hours Worked")
        fig5 = px.box(filtered_df, x='employee_resident', y='hours_worked', color='employee_resident')
        st.plotly_chart(fig5, use_container_width=True)

    with row3_col2:
        st.subheader("🏅 Top 5 Most Punctual Employees")
        top5 = filtered_df[filtered_df['is_punctual'] == True]['employee_id'].value_counts().nlargest(5).reset_index()
        top5.columns = ['Employee ID', 'Punctual Days']
        fig_top5 = px.bar(top5, x='Employee ID', y='Punctual Days', color='Employee ID', text='Punctual Days')
        fig_top5.update_layout(showlegend=False)
        st.plotly_chart(fig_top5, use_container_width=True)

    row4_col1, row4_col2 = st.columns(2)
    with row4_col1:
        st.subheader("🚨 Top 5 Late Comers (Hours < 8)")
        bottom5 = filtered_df[filtered_df['is_punctual'] == False]['employee_id'].value_counts().nlargest(5).reset_index()
        bottom5.columns = ['Employee ID', 'Late Days']
        fig_bottom5 = px.bar(bottom5, x='Employee ID', y='Late Days', color='Employee ID', text='Late Days')
        fig_bottom5.update_layout(showlegend=False)
        st.plotly_chart(fig_bottom5, use_container_width=True)

    with row4_col2:
        st.subheader("⚖️ Punctuality vs Late Days Comparison")
        top_late_ids = bottom5['Employee ID'].tolist()
        compare_df = filtered_df[filtered_df['employee_id'].isin(top_late_ids)]
        compare_summary = compare_df.groupby(['employee_id', 'is_punctual']).size().reset_index(name='Count')
        compare_summary['Status'] = compare_summary['is_punctual'].map({True: 'Punctual Days', False: 'Late Days'})
        fig_compare = px.bar(compare_summary, x='employee_id', y='Count', color='Status', barmode='group')
        st.plotly_chart(fig_compare, use_container_width=True)

# --- Tab 2: Summary ---
with tab2:
    st.subheader("📄 Executive Summary")

    # Core Metrics
    total_employees = filtered_df['employee_id'].nunique()
    total_days = filtered_df['date'].nunique()
    total_records = len(filtered_df)
    total_punctual = filtered_df[filtered_df['is_punctual']].shape[0]
    avg_hours_worked = round(filtered_df['hours_worked'].mean(), 2)
    punctuality_rate = round((total_punctual / total_records) * 100, 2) if total_records else 0.0

    # Insights
    st.markdown(f"""
    #### 🧾 Attendance Overview
    - **Total Employees Analysed:** {total_employees}
    - **Total Attendance Records:** {total_records}
    - **Total Unique Days Tracked:** {total_days}

    #### ⏱️ Work Hours Insights
    - **Average Daily Hours Worked:** {avg_hours_worked:.2f} hours
    - **Employees Meeting 8+ Hours Criteria:** {total_punctual} entries

    #### ✅ Punctuality Insights
    - **Overall Punctuality Rate:** {punctuality_rate:.2f}%
    - **Standard Threshold:** Employees are considered **punctual** if they work **≥ 8 hours per day**
    """)

    st.success("This detailed summary provides a managerial view of employee attendance patterns, average productivity, and organizational discipline levels.")

# --- Tab 3: Download ---
with tab3:
    st.subheader("📥 Download Processed Data")

    # 📌 Display Month-Year from selected start date
    month_label = pd.to_datetime(start_date_input).strftime("%B %Y")
    st.markdown(f"### 🗓️ Download Monthly Punctuality Summary for **{month_label}**")

    # Add Month-Year column
    download_df = filtered_df.copy()
    download_df['month_year'] = download_df['date'].dt.to_period('M').astype(str)

    # Monthly Summary
    monthly_summary_df = download_df.groupby(['employee_id', 'month_year']).agg(
        Total_Days=('date', 'count'),
        Punctual_Days=('is_punctual', 'sum'),
        Late_Days=('is_punctual', lambda x: (~x).sum()),
        Punctuality_Rate=('is_punctual', lambda x: round(x.mean() * 100, 2)),
        Avg_Hours_Worked=('hours_worked', lambda x: round(x.mean(), 2))
    ).reset_index()

    # ➕ Add Punctual_Status column
    monthly_summary_df['Punctual_Status'] = monthly_summary_df['Punctuality_Rate'].apply(
        lambda x: "Yes" if x >= 90 else "No"
    )

    # Provide CSV download
    csv_data = monthly_summary_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📄 Download Monthly Summary CSV",
        data=csv_data,
        file_name=f'monthly_punctuality_summary_{month_label.replace(" ", "_")}.csv',
        mime='text/csv'
    )

# --- Tab 4: Email Summary ---
with tab4:
    st.subheader("📬 Email Summary to Manager")
    sender_email = st.text_input("📧 Enter your Gmail address")
    sender_password = st.text_input("🔐 Enter App Password", type="password")
    recipient_email = st.text_input("📨 Enter Manager's Email")
    send_email = st.button("📧 Send Summary")

    if send_email and sender_email and sender_password and recipient_email:
        with st.spinner("📤 Sending email..."):
            try:
                msg = EmailMessage()
                msg['Subject'] = "Employee Attendance Summary"
                msg['From'] = sender_email
                msg['To'] = recipient_email
                msg.set_content(
                    "Hi,\n\nPlease find attached the daily and monthly employee attendance summary.\n\nRegards,\nDashboard System")

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    monthly_summary_df.to_excel(writer, index=False, sheet_name='Monthly Summary')
                output.seek(0)
                msg.add_attachment(
                    output.read(),
                    maintype='application',
                    subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    filename='EmployeeAttendanceSummary.xlsx'
                )

                with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                    smtp.login(sender_email, sender_password)
                    smtp.send_message(msg)

                st.success("✉️ Email sent successfully!")
            except Exception as e:
                st.error(f"❌ Something went wrong: {e}")
    elif send_email:
        st.warning("⚠️ Please enter all email credentials correctly.")

# =============================
# 📎 Footer
# =============================
st.markdown("---")
st.markdown("© 2025 Diverse Infotech Pvt Ltd | Built by AYRS")














