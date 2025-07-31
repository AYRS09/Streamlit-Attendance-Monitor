import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
import io
import smtplib
from email.message import EmailMessage
from PIL import Image

# Theme toggle
theme = st.sidebar.radio("🌓 Choose Theme", ["Light", "Dark"])
if theme == "Dark":
    st.markdown(
        """
        <style>
        body { background-color: #0E1117; color: white; }
        .stApp { background-color: #0E1117; color: white; }
        </style>
        """,
        unsafe_allow_html=True
    )

# --- Streamlit Config ---
st.set_page_config(page_title="Employee Punctuality Dashboard", layout="wide")

# --- Load Logo ---
import os

curr_dir = os.path.dirname(__file__)
image_path = os.path.join(curr_dir, "download.jpeg")

if os.path.exists(image_path):
    st.sidebar.image(image_path, width=120)
    st.sidebar.markdown("### 👋 Welcome to the Dashboard")
    
else:
    st.sidebar.warning("⚠️ Logo image not found.")

from datetime import datetime

# Show Last Updated Timestamp on top-right
now = datetime.now().strftime("%d %b %Y, %I:%M %p")
st.markdown(
    f"<div style='text-align:right; color:gray; font-size:0.85rem;'>🕒 Last updated: {now}</div>",
    unsafe_allow_html=True
)

# --- Title ---
st.markdown("<h1 style='text-align: center; color: #4B8BBE;'>📊 Employee Productivity Dashboard | Diverse Infotech Pvt Ltd</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; color: gray;'>Punctuality & Productivity Analysis Based on Daily Hours Worked</h4>", unsafe_allow_html=True)

# --- File Upload ---
st.sidebar.markdown("---")
st.sidebar.subheader("📄 Upload Attendance Sheet")
file = st.sidebar.file_uploader("Upload Excel/CSV File", type=["xlsx", "xls", "csv"])
st.sidebar.info("ℹ️ Upload your Excel or CSV attendance file to view the dashboard.")
st.sidebar.markdown("---")

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
    try:
        df[hours_col] = (
            pd.to_datetime(df[out_col], format='%I:%M %p', errors='coerce') -
            pd.to_datetime(df[in_col], format='%I:%M %p', errors='coerce')
        ).dt.total_seconds() / 3600
        df[hours_col] = df[hours_col].round(2)
    except Exception as e:
        st.warning(f"⚠️ Error calculating hours for {in_col} & {out_col}: {e}")

# Step 1: Drop perfect duplicate rows, if any
df.drop_duplicates(inplace=True)

# Step 2: Check for employees with multiple rows
duplicate_ids = df['employee_id'].value_counts()
duplicate_ids = duplicate_ids[duplicate_ids > 1]

# Optional Debug Info
if not duplicate_ids.empty:
    st.warning("⚠️ Found duplicate entries for these employee IDs:")
    st.dataframe(df[df['employee_id'].isin(duplicate_ids.index)])

# Step 3: Combine duplicates by taking the row with max total hours
# Sum hours across all 'hours_' columns
df['total_hours'] = df[[col for col in df.columns if col.startswith('hours_')]].sum(axis=1)

# Keep only the row with max total hours per employee_id
df = df.sort_values('total_hours', ascending=False).drop_duplicates(subset=['employee_id'], keep='first')

# Drop the helper column
df.drop(columns='total_hours', inplace=True)

# --- Day Columns ---
day_cols = sorted([col for col in df.columns if col.startswith('hours_')], key=lambda x: int(x.split('_')[1]))

# --- Melt for long format ---
df_long = df.melt(
    id_vars=[
        'employee_id',
        'employee_gender',
        'employee_resident',
        'employee_department'
    ],
    value_vars=day_cols,
    var_name='day',
    value_name='hours_worked'
)

# Extract day number and convert to date
df_long['day_num'] = df_long['day'].str.extract(r'(\d+)').astype(int)
df_long['date'] = pd.to_datetime('2025-06-01') + pd.to_timedelta(df_long['day_num'] - 1, unit='D')

# Add punctuality flag
df_long['is_punctual'] = df_long['hours_worked'] >= 8

# --- Sidebar Filters ---
st.sidebar.header("🔍 Filter Options")
employees = sorted(df_long['employee_id'].dropna().unique())
selected_employees = st.sidebar.selectbox("👤 Select Employee", options=["All"] + list(employees))

residency = st.sidebar.selectbox("🏩 Resident Type", options=["All", "Local", "Non-local"])
departments = sorted(df_long['employee_department'].dropna().unique())
selected_departments = st.sidebar.multiselect("🏢 Select Department(s)", options=departments, default=departments)

# --- Date Range Filter ---
st.sidebar.markdown("🗓️ **Date Range Filter**")
min_date = df_long['date'].min()
max_date = df_long['date'].max()
date_range = st.sidebar.date_input("Select Date Range", [min_date, max_date], min_value=min_date, max_value=max_date)

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

# --- KPIs ---
total_employees = filtered_df['employee_id'].nunique()
total_days = len(filtered_df)
total_punctual = filtered_df[filtered_df['is_punctual']].shape[0]
avg_hours = round(filtered_df['hours_worked'].mean(), 2)
punctuality_rate = round((total_punctual / total_days) * 100, 2) if total_days else 0.0

st.markdown("## 📌 Key Metrics")
kpi1, kpi2, kpi3 = st.columns(3)
kpi1.metric("👥 Total Employees", total_employees)
kpi2.metric("✅ Punctuality Rate", f"{punctuality_rate}%")
kpi3.metric("⏱️ Average Hours Worked", f"{avg_hours} hrs")

st.markdown("---")

# --- Tabs ---
tab1, tab2, tab3, tab4 = st.tabs(["📊 Visualizations", "📋 Summary", "📅 Download", "📬 Email Summary"])

# --- Tab 1: Visualizations ---
with tab1:
    row1_col1, row1_col2 = st.columns(2)
    with row1_col1:
        st.subheader("📌 Total Hours Worked per Employee")
        fig1 = px.bar(
            filtered_df.groupby('employee_id')['hours_worked'].sum().reset_index(),
            x='employee_id', y='hours_worked', color='hours_worked', color_continuous_scale='Greens')
        st.plotly_chart(fig1, use_container_width=True)

    with row1_col2:
        st.subheader("📌 Punctuality Ratio per Employee")
        punctual_summary = filtered_df.groupby(['employee_id', 'is_punctual']).size().reset_index(name='Count')
        punctual_summary['Status'] = punctual_summary['is_punctual'].map({True: 'Met (≥8 hrs)', False: 'Not Met (<8 hrs)'})
        fig2 = px.bar(punctual_summary, x='employee_id', y='Count', color='Status', barmode='stack')
        st.plotly_chart(fig2, use_container_width=True)

    row2_col1, row2_col2 = st.columns(2)
    with row2_col1:
        st.subheader("📌 Daily Productivity Heatmap")
        heatmap_data = filtered_df.pivot_table(index='employee_id', columns='day_num', values='hours_worked')
        fig3 = px.imshow(heatmap_data.astype(np.float32), aspect="auto", color_continuous_scale='YlGnBu')
        st.plotly_chart(fig3, use_container_width=True)

    with row2_col2:
        st.subheader("📌 Overall Punctuality Ratio")
        overall = filtered_df['is_punctual'].value_counts().rename({True: 'Met ≥8 hrs', False: 'Not Met <8 hrs'})
        fig4 = px.pie(names=overall.index, values=overall.values)
        st.plotly_chart(fig4, use_container_width=True)

    row3_col1, row3_col2 = st.columns(2)
    with row3_col1:
        st.subheader("📌 Resident Type vs Hours Worked")
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

    st.write("Columns available:", df.columns.tolist())

    total_employees = filtered_df['employee_id'].nunique()

    punctuality_rate = round(filtered_df['is_punctual'].mean() * 100, 2)
    avg_hours_worked = round(filtered_df['hours_worked'].mean(), 2)

    st.markdown(f"""
    - **Total Employees:** `{total_employees}`
    - **Punctuality Rate:** `{punctuality_rate:.2f}%`
    - **Average Hours Worked:** `{avg_hours_worked:.2f} hrs`
    """)

    st.success("This summary gives a quick snapshot of overall team attendance and productivity.")
                  
# --- Tab 3: Download ---
with tab3:
    st.subheader("📥 Download Processed Data")
    
    # Optional: preview filtered data
    st.dataframe(filtered_df)

    # Prepare CSV
    csv = filtered_df.to_csv(index=False).encode('utf-8')

    # Download button
    st.download_button(
        label="📄 Download CSV",
        data=csv,
        file_name='processed_attendance.csv',
        mime='text/csv'
    )

    st.caption("Click the button above to download the filtered attendance data.")

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
                    "Hi,\n\nPlease find attached the latest employee attendance summary.\n\nRegards,\nDashboard System")

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='Summary')
                output.seek(0)
                msg.add_attachment(
                    output.read(),
                    maintype='application',
                    subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    filename='EmployeeSummary.xlsx'
                )

                with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                    smtp.login(sender_email, sender_password)
                    smtp.send_message(msg)

                st.success("✉️ Email sent successfully!")
            except Exception as e:
                st.error(f"❌ Something went wrong: {e}")
    elif send_email:
        st.warning("⚠️ Please enter all email credentials correctly.")

# --- Footer ---
st.markdown("---")
st.markdown("© 2025 Diverse Infotech Pvt Ltd | Built by AYRS")
