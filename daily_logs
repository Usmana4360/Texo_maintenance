import streamlit as st
import pandas as pd
import requests
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from datetime import datetime

# ---------------------- Constants ----------------------
EXCEL_FILE = "generated_reports.xlsx"
PANEL_FILE = "lt_panel_checklist.xlsx"
HF_API_KEY = os.getenv("HF_API_KEY")  # Or use st.secrets
HF_API_URL = "https://api-inference.huggingface.co/models/mistralai/Mistral-7B-Instruct-v0.1"

# st.set_page_config(page_title="Maintenance Tools", layout="wide")
# st.sidebar.title("🔧 Maintenance Dashboard")
# selected_tool = st.sidebar.radio("Select Function", ["Report Generator", "LT Panel Entry", "Compressor Readings"])


# ---------------------- Maintenance Report Generator ----------------------
def maintenance_report_ui():
    st.title("📝 AI-Powered Maintenance Report Generator (Mistral)")

    col1, col2 = st.columns(2)
    with col1:
        unit = st.text_input("Enter Unit:", placeholder="e.g., Unit 5")
        technician_name = st.text_input("Enter Technician Name:", placeholder="e.g., John Doe")
    with col2:
        machine = st.text_input("Enter Machine Name:", placeholder="e.g., Compressor A1")
        issue = st.text_input("Enter Issue Reported:", placeholder="e.g., High temperature issue")

    if st.button("Generate Report"):
        if unit and machine and technician_name and issue:
            with st.spinner("Generating report using Mistral..."):
                report = generate_report(unit, machine, technician_name, issue)
                save_to_excel(unit, machine, technician_name, issue, report)
                st.success("✅ Report generated and saved!")
                st.subheader("Generated Report:")
                st.write(report)
        else:
            st.warning("⚠ Please fill in all fields before generating the report.")

    st.subheader("📁 Report History")
    df_reports = load_reports()
    st.dataframe(df_reports, use_container_width=True)

    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as file:
            st.download_button("📥 Download Reports", file, file_name="generated_reports.xlsx")


def generate_report(unit, machine, technician_name, issue):
    prompt = f"""
You are an expert electrical maintenance engineer. Generate a **concise and professional** one-line report based on the following details:
Unit: {unit}
Machine: {machine}
Technician Name: {technician_name}
Issue Reported: {issue}
- Keep it one line only.
"""
    headers = {"Authorization": f"Bearer {HF_API_KEY}", "Content-Type": "application/json"}
    payload = {"inputs": prompt, "parameters": {"max_new_tokens": 100, "temperature": 0.7, "return_full_text": False}}

    try:
        response = requests.post(HF_API_URL, headers=headers, json=payload)
        if response.status_code == 200:
            return response.json()[0]["generated_text"].strip()
        else:
            return f"Issue reported and solved by {technician_name}: {issue}"
    except Exception:
        return f"Issue reported and solved by {technician_name}: {issue}"


def save_to_excel(unit, machine, technician_name, issue, report):
    new_data = pd.DataFrame([{
        "Date": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Unit": unit, "Machine": machine,
        "Technician Name": technician_name, "Issue": issue,
        "Generated Report": report
    }])
    if os.path.exists(EXCEL_FILE):
        try:
            existing_data = pd.read_excel(EXCEL_FILE, engine="openpyxl")
            updated_data = pd.concat([existing_data, new_data], ignore_index=True)
        except:
            updated_data = new_data
    else:
        updated_data = new_data

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="w") as writer:
        updated_data.to_excel(writer, index=False, sheet_name="Work Orders")

    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                         top=Side(style="thin"), bottom=Side(style="thin"))
    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(wrap_text=True, vertical="center")
    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col if cell.value)
        ws.column_dimensions[col[0].column_letter].width = max_length + 2
    for cell in ws[1]:
        cell.font = Font(bold=True)
    wb.save(EXCEL_FILE)


def load_reports():
    if os.path.exists(EXCEL_FILE):
        return pd.read_excel(EXCEL_FILE)
    return pd.DataFrame(columns=["Unit", "Machine", "Technician", "Issue", "Report"])


# ---------------------- LT Panel Readings Entry ----------------------
LT_PANEL_FILE = "lt_panel_log.xlsx"


def lt_panel_ui():
    st.title("📊 LT Panel Readings Entry")

    with st.form("lt_panel_form"):
        st.subheader("General Information")
        col1, col2 = st.columns(2)
        with col1:
            date = st.date_input("Date", value=datetime.today())
            shift = st.selectbox("Shift", ["A", "B", "C"])
        with col2:
            time = st.time_input("Time", value=datetime.now().time())
            reader = st.text_input("Technician Name / Reader")

        st.subheader("Panel Readings")
        readings = {}
        panels = ["LT Panel 1", "LT Panel 2", "LT Panel 3", "LT Panel 4", "Tapline", "Looms Panel"]
        for panel in panels:
            readings[panel] = lt_panel_input(panel)

        submitted = st.form_submit_button("Submit LT Panel Data")
        if submitted:
            save_lt_panel_data(date, shift, time, reader, readings)
            st.success("✅ LT Panel data saved successfully!")

    if os.path.exists(LT_PANEL_FILE):
        with open(LT_PANEL_FILE, "rb") as file:
            st.download_button("📥 Download LT Panel Sheet", file, file_name="lt_panel_log.xlsx")



def lt_panel_input(panel_name):
    st.markdown(f"### {panel_name}")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        volt = st.number_input(f"{panel_name} - Volt", key=f"{panel_name}_volt", value=0.0)
    with col2:
        amp = st.number_input(f"{panel_name} - Amp", key=f"{panel_name}_amp", value=0.0)
    with col3:
        pf = st.number_input(f"{panel_name} - PF", key=f"{panel_name}_pf", value=0.98, step=0.01)
    with col4:
        temp = st.number_input(f"{panel_name} - Temp", key=f"{panel_name}_temp", value=0.0)
    return volt, amp, pf, temp


def save_lt_panel_data(date, shift, time, reader, readings):
    from openpyxl import Workbook, load_workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.styles import Font, Border, Side, Alignment
    from openpyxl.worksheet.table import Table, TableStyleInfo
    import os

    if os.path.exists(LT_PANEL_FILE):
        wb = load_workbook(LT_PANEL_FILE)
    else:
        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

    for panel_name, values in readings.items():
        row_data = {
            "Date": date.strftime("%Y-%m-%d"),
            "Shift": shift,
            "Time": time.strftime("%H:%M"),
            "Technician": reader,
            "Volt": values[0],
            "Amp": values[1],
            "PF": values[2],
            "Temp": values[3]
        }

        df = pd.DataFrame([row_data])

        if panel_name in wb.sheetnames:
            ws = wb[panel_name]
            for r in dataframe_to_rows(df, index=False, header=False):
                ws.append(r)
        else:
            ws = wb.create_sheet(panel_name)
            for r in dataframe_to_rows(df, index=False, header=True):
                ws.append(r)

            # Format header
            for cell in ws[1]:
                cell.font = Font(bold=True)

            # Create table
            ref = f"A1:{chr(64 + ws.max_column)}{ws.max_row}"
            table = Table(displayName=f"{panel_name.replace(' ', '')}Table", ref=ref)
            style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
            table.tableStyleInfo = style
            ws.add_table(table)

        # Apply formatting
        border = Border(left=Side(style="thin"), right=Side(style="thin"),
                        top=Side(style="thin"), bottom=Side(style="thin"))
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for col in ws.columns:
            max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max_len + 2

    wb.save(LT_PANEL_FILE)

# ---------------------- Compressor Readings Logger ----------------------

def compressor_excel_logger():
    st.header("🧾 Compressor Data Logger")

    EXCEL_PATH = "compressor_log.xlsx"

    compressor_units = {
        "Unit 1 (37 KW)": {},
        "Unit 2 (37 KW)": {},
        "Unit 3 (55 KW)": {},
        "Unit 4 (55 KW)": {},
        "Unit 5 (22 KW)": {}
    }

    with st.form("compressor_log_form"):
        st.subheader("Log Details")
        c1, c2 = st.columns(2)
        with c1:
            log_date = st.date_input("Log Date", value=datetime.today(), key="log_date")
        with c2:
            shift_code = st.selectbox("Shift", ["A", "B", "C"], key="log_shift")

        st.markdown("### Enter Compressor Readings")

        for unit in compressor_units.keys():
            st.markdown(f"#### {unit}")
            c1, c2, c3 = st.columns(3)
            with c1:
                amps = st.number_input(f"{unit} - Amperes", min_value=0.0, key=f"{unit}_amps")
            with c2:
                temperature = st.number_input(f"{unit} - Temperature (°C)", min_value=0.0, key=f"{unit}_temp")
            with c3:
                pressure = st.number_input(f"{unit} - Pressure (bar)", min_value=0.0, key=f"{unit}_pres")

            compressor_units[unit] = {
                "Amps": amps,
                "TempC": temperature,
                "PressBar": pressure
            }

        save_log = st.form_submit_button("➕ Save Log")

        if save_log:
            try:
                if os.path.exists(EXCEL_PATH):
                    try:
                        workbook = load_workbook(EXCEL_PATH)
                    except Exception:
                        os.remove(EXCEL_PATH)
                        workbook = Workbook()
                        workbook.remove(workbook.active)
                else:
                    workbook = Workbook()
                    workbook.remove(workbook.active)

                for unit, data in compressor_units.items():
                    sheet = unit.replace(" ", "_").replace("(", "").replace(")", "")
                    if sheet not in workbook.sheetnames:
                        ws = workbook.create_sheet(title=sheet)
                        ws.append(["Date", "Shift", "Amps", "Temp (°C)", "Pressure (bar)"])
                    else:
                        ws = workbook[sheet]

                    ws.append([
                        log_date.strftime("%Y-%m-%d"),
                        shift_code,
                        data["Amps"],
                        data["TempC"],
                        data["PressBar"]
                    ])

                    border = Border(left=Side(style="thin"), right=Side(style="thin"),
                                    top=Side(style="thin"), bottom=Side(style="thin"))
                    for row in ws.iter_rows():
                        for cell in row:
                            cell.border = border
                            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    for cell in ws[1]:
                        cell.font = Font(bold=True)
                    for col in ws.columns:
                        max_len = max(len(str(cell.value)) for cell in col if cell.value)
                        ws.column_dimensions[col[0].column_letter].width = max_len + 2

                workbook.save(EXCEL_PATH)
                st.success("✅ Compressor log saved successfully!")

            except Exception as err:
                st.error(f"❌ Failed to log data: {err}")

    if os.path.exists(EXCEL_PATH):
        with open(EXCEL_PATH, "rb") as f:
            st.download_button("📥 Download Compressor Log File", f, file_name="compressor_log.xlsx")

CHILLER_EXCEL_FILE = "chiller_readings.xlsx"

CHILLER_NAMES = {
    "CHILLER NO 1 (UNIT NO 1)": "Chiller 1",
    "CHILLER NO 2 (BACKUP FOR UNIT NO 1)": "Chiller 2",
    "CHILLER NO 3 (UNIT NO 2)": "Chiller 3",
    "CHILLER NO 4 (UNIT NO 3)": "Chiller 4",
    "CHILLER NO 5 (BACKUP FOR UNIT NO 3)": "Chiller 5",
    "CHILLER NO 6 (UNIT NO 3)": "Chiller 6",
    "CHILLER NO 7 (UNIT NO 4)": "Chiller 7",
    "CHILLER NO 8 (BACKUP FOR UNIT NO 6)": "Chiller 8",
    "CHILLER NO 9 (UNIT NO 6)": "Chiller 9",
    "CHILLER NO 10 (UNIT NO 8)": "Chiller 10",
    "CHILLER NO 11 (UNIT NO 8)": "Chiller 11", 
}

def save_chiller_data_separate_sheets(shift, time, chiller_readings):
    time_str = time.strftime("%I:%M %p")

    if os.path.exists(CHILLER_EXCEL_FILE):
        wb = load_workbook(CHILLER_EXCEL_FILE)
    else:
        wb = Workbook()
        wb.remove(wb.active)

    for chiller_name, readings in chiller_readings.items():
        sheet_name = CHILLER_NAMES[chiller_name]
        if sheet_name not in wb.sheetnames:
            ws = wb.create_sheet(title=sheet_name)
            ws.append(["SHIFT", "TIME", "AMP", "COOLING TEMP", "PRESSURE", "OIL LEVEL"])
        else:
            ws = wb[sheet_name]

        ws.append([shift, time_str] + readings)

        # Format: borders, fonts, column widths
        border = Border(left=Side(style="thin"), right=Side(style="thin"),
                        top=Side(style="thin"), bottom=Side(style="thin"))

        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for cell in ws[1]:
            cell.font = Font(bold=True)

        for col in ws.columns:
            max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max_len + 2

    wb.save(CHILLER_EXCEL_FILE)


def chiller_ui():
    st.title("🌬️ Daily Chillers Checklist")

    with st.form("chiller_form"):
        shift = st.selectbox("Shift", ["A", "B", "C"])
        time = st.time_input("Time", value=datetime.now().time())

        chiller_readings = {}

        for name in CHILLER_NAMES:
            st.markdown(f"### {name}")
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                amp = st.text_input(f"{name} - Amp", key=f"{name}_amp")
            with c2:
                temp = st.text_input(f"{name} - Cooling Temp", key=f"{name}_temp")
            with c3:
                pressure = st.text_input(f"{name} - Pressure", key=f"{name}_pressure")
            with c4:
                oil = st.text_input(f"{name} - Oil Level", key=f"{name}_oil")

            chiller_readings[name] = [amp, temp, pressure, oil]

        submitted = st.form_submit_button("Submit Chiller Readings")
        if submitted:
            save_chiller_data_separate_sheets(shift, time, chiller_readings)
            st.success("✅ Chiller readings saved to separate sheets!")

    if os.path.exists(CHILLER_EXCEL_FILE):
        with open(CHILLER_EXCEL_FILE, "rb") as file:
            st.download_button("📥 Download Chiller Excel File", file, file_name=CHILLER_EXCEL_FILE)



def run_app():
    st.set_page_config(page_title="Maintenance Tools", layout="wide")
    st.sidebar.title("🔧 Maintenance Dashboard")

    selected_tool = st.sidebar.radio(
        "Select Function", ["Report Generator", "LT Panel Entry", "Compressor Readings","Chiller Readings"]
    )

    if selected_tool == "Report Generator":
        maintenance_report_ui()
    elif selected_tool == "LT Panel Entry":
        lt_panel_ui()
    elif selected_tool == "Compressor Readings":
        compressor_excel_logger()
    elif selected_tool == "Chiller Readings":
        chiller_ui()

# Call this at the end of your script
if __name__ == "__main__":
    run_app()
