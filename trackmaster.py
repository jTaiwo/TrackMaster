import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import random
import pandas as pd
from io import BytesIO
import json
import os
from google.oauth2 import service_account

# Function to authenticate and connect to Google Sheets
def connect_to_sheets():
    # Load the service account key from Streamlit secrets
    service_account_info = json.loads(st.secrets["gcp"]["key"])
    
    # Create the credentials using the service account info
    creds = service_account.Credentials.from_service_account_info(service_account_info)

    # Authorize the gspread client with the credentials
    client = gspread.authorize(creds)
    return client

# Function to generate a unique Repair ID
def generate_repair_id():
    timestamp = int(time.time())
    random_number = random.randint(1000, 9999)
    return f"R{timestamp}{random_number}"

# Function to submit the repair record
def submit_repair_record(asset_id, repair_date, technician_name, diagnosis_report, recommended_solutions, repair_actions_taken, notes):
    client = connect_to_sheets()
    spreadsheet = client.open_by_key("1vPT_3GSVaM3AIcpsL6sTfgl_GyoNRnqAwaswzTzY6h8")
    repair_records_sheet = spreadsheet.worksheet("Repair Records")
    
    # Generate a unique Repair ID
    repair_id = generate_repair_id()
    
    # Get asset details from Asset Information sheet
    asset_info_sheet = spreadsheet.worksheet("Asset Information")
    asset_cell = asset_info_sheet.find(asset_id)
    
    if not asset_cell:
        st.error("Asset ID not found. Please enter a valid Asset ID.")
        return
    
    asset_row = asset_cell.row
    asset_details = asset_info_sheet.row_values(asset_row)[1:]  # Exclude Asset ID
    
    # Append the new repair record
    repair_records_sheet.append_row([
        repair_id, asset_id, *asset_details, repair_date, technician_name, diagnosis_report, 
        recommended_solutions, repair_actions_taken, notes
    ])
    
    # Update the Asset Information sheet
    update_asset_info(asset_id)
    st.success(f"Repair record submitted with Repair ID: {repair_id} and asset information updated.")

def update_asset_info(asset_id):
    client = connect_to_sheets()
    spreadsheet = client.open_by_key("1vPT_3GSVaM3AIcpsL6sTfgl_GyoNRnqAwaswzTzY6h8")
    asset_info_sheet = spreadsheet.worksheet("Asset Information")
    repair_records_sheet = spreadsheet.worksheet("Repair Records")
    
    cell = asset_info_sheet.find(asset_id)
    if not cell:
        st.error("Asset ID not found.")
        return
    
    asset_row = cell.row
    repair_records = repair_records_sheet.findall(asset_id)
    repair_count = len(repair_records)
    last_repair_date = max([repair_records_sheet.cell(r.row, repair_records_sheet.find("Date of Repair").col).value for r in repair_records], default="")
    
    if repair_count == 0 or repair_count == 1:
        status = "Good"
    elif repair_count == 2:
        status = "Fair"
    else:
        status = "Recommended for Replacement"
    
    asset_info_sheet.update_cell(asset_row, asset_info_sheet.find("Repair Count").col, repair_count)
    asset_info_sheet.update_cell(asset_row, asset_info_sheet.find("Date of Last Repair").col, last_repair_date)
    asset_info_sheet.update_cell(asset_row, asset_info_sheet.find("Status").col, status)

# Function to retrieve asset information and repair records
def retrieve_asset_info(asset_id):
    client = connect_to_sheets()
    spreadsheet = client.open_by_key("1vPT_3GSVaM3AIcpsL6sTfgl_GyoNRnqAwaswzTzY6h8")
    asset_info_sheet = spreadsheet.worksheet("Asset Information")
    repair_records_sheet = spreadsheet.worksheet("Repair Records")
    
    # Get asset details
    asset_cell = asset_info_sheet.find(asset_id)
    if not asset_cell:
        return None, None

    asset_row = asset_cell.row
    asset_details = asset_info_sheet.row_values(asset_row)
    
    # Get all repair records
    repair_records = repair_records_sheet.get_all_values()
    
    # Convert to DataFrame for easier manipulation
    headers = repair_records[0]
    records_df = pd.DataFrame(repair_records[1:], columns=headers)
    
    # Filter records for the given Asset ID
    asset_repair_records = records_df[records_df['Asset ID'] == asset_id]
    
    return asset_details, asset_repair_records.to_dict('records')

# Streamlit UI
st.title("Asset Management System")

# Tabs for switching between sections
tabs = st.tabs(["Submit Repair Record", "Retrieve Asset Information"])

with tabs[0]:
    st.header("Submit Repair Record")

    asset_id = st.text_input("Asset ID")
    repair_date = st.date_input("Date of Repair")
    technician_name = st.text_input("Technician Name")
    diagnosis_report = st.text_area("Diagnosis Report")
    recommended_solutions = st.text_area("Recommended Solutions")
    repair_actions_taken = st.text_area("Repair Actions Taken")
    notes = st.text_area("Notes")

    if st.button("Submit"):
        if asset_id and technician_name:
            submit_repair_record(asset_id, str(repair_date), technician_name, diagnosis_report, recommended_solutions, repair_actions_taken, notes)
        else:
            st.error("Please fill in all required fields.")

with tabs[1]:
    st.header("Retrieve Asset Information")

    retrieve_asset_id = st.text_input("Enter Asset ID to Retrieve Information")

    if st.button("Retrieve Info"):
        if retrieve_asset_id:
            asset_details, asset_repair_records = retrieve_asset_info(retrieve_asset_id)
            
            if asset_details:
                st.subheader("Asset Information")
                st.write({
                    'Asset ID': asset_details[0],
                    'Asset Name': asset_details[1],
                    'Asset Type': asset_details[2],
                    'Location': asset_details[3],
                    'Purchase Date': asset_details[4],
                    'Manufacturer': asset_details[5],
                    'Serial Number': asset_details[6],
                    'Repair Count': asset_details[7],
                    'Date of Last Repair': asset_details[8],
                    'Status': asset_details[9]
                })
                
                st.subheader("Repair Records")
                df = pd.DataFrame(asset_repair_records)
                if not df.empty:
                    st.write(df)
                    
                    # Create a downloadable Excel file
                    excel_file = BytesIO()
                    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, sheet_name='Repair Records')
                    
                    st.download_button(
                        label="Download Repair Records as Excel",
                        data=excel_file.getvalue(),
                        file_name=f"{retrieve_asset_id}_repair_records.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("No repair records found for this Asset ID.")
            else:
                st.error("Asset ID not found.")
        else:
            st.error("Please enter an Asset ID.")
