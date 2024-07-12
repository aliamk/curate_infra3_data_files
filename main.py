import streamlit as st
import pandas as pd
import tempfile
import os
import re
from openpyxl.utils import get_column_letter
import pytz
from datetime import datetime

def process_transaction_sheet(transaction_df):
    # Define a function to extract numerical value from a string
    def extract_numerical_value(text):
        match = re.search(r'[\d,.]+', str(text))
        return match.group() if match else ''

    # Map columns for the 'Transaction' sheet with transformations
    return pd.DataFrame({
        "Transaction Upload ID": transaction_df["Transaction Upload ID"],
        "Transaction Name": transaction_df["Transaction Name"],
        "Transaction Asset Class": "Infrastructure",  # Static value for all rows
        "Transaction Status": transaction_df.get("Current Status", ""),  # Using .get to avoid KeyError if not present
        "Finance Type": transaction_df.get("Finance Type", ""),
        "Transaction Type": transaction_df.get("Type", ""),
        "Unknown Asset": "",
        "Underlying Asset Configuration": "",
        "Transaction Local Currency": transaction_df["Transaction Currency"].apply(extract_numerical_value),  # Extract numerical value
        "Transaction Value (Local Currency)": transaction_df.get("Transaction size (m)", ""),
        "Transaction Debt (Local Currency)": "",
        "Transaction Equity (Local Currency)": "",
        "Debt/Equity Ratio": "",
        "Underlying Number of Assets": "",
        "Region - Country": transaction_df.get("Geography", ""),
        "Region - State": "",
        "Region - City": "",
        "Any Level Sectors": transaction_df.get("Sector", "") + ", " + transaction_df.get("Sub-Sector", ""),
        "PPP": transaction_df.get("PPP", ""),
        "Concession Period": transaction_df.get("Duration", ""),
        "Contract": transaction_df.get("Delivery Model", ""),
        "SPV": transaction_df.get("SPV", ""),
        "Active": "TRUE"
    })

def process_events_sheet(transaction_df):
    # Define the event types with their respective date columns and static or dynamic labels
    event_details = [
        ("Current status date", "Current status"),  # Dynamic label copied from the source file
        ("Financial close", "Financial Close"),  # Static label
        ("Transaction Launch", "Announced"),
        ("RFP returned", "Request for Proposals"),
        ("Preferred Proponents", "Preferred Bidder"),
        ("Expressions of Interest", "Expression of Interest"),
        ("RFQ returned", "Request for Qualifications"),
        ("Shortlisted proponents", "Shortlist")
    ]
    
    # Create an empty DataFrame to store event data
    events_df = pd.DataFrame(columns=["Transaction Upload ID", "Event Date", "Event Type", "Event Title"])

    for date_column, event_type in event_details:
        # Filter out rows with empty dates
        filtered_df = transaction_df.dropna(subset=[date_column])

        # Determine the event type: static or from a source column
        if event_type == "Current status":
            event_type_series = filtered_df[event_type]  # Copy dynamically from the 'Current status' column
        else:
            event_type_series = pd.Series([event_type] * len(filtered_df), index=filtered_df.index)  # Use static label
        
        # Create a temporary DataFrame for the current event type
        temp_df = pd.DataFrame({
            "Transaction Upload ID": filtered_df["Transaction Upload ID"],
            "Event Date": pd.to_datetime(filtered_df[date_column]).dt.date,  # Convert to date only
            "Event Type": event_type_series,
            "Event Title": ""  # Assuming this field is empty or to be populated later
        })

        # Append the temporary DataFrame to the main events DataFrame
        events_df = pd.concat([events_df, temp_df], ignore_index=True)

    return events_df

def process_bidders_any_sheet(transaction_df):
    sources = {
        "Legal Advisors": "Adviser",
        "Technical Advisors": "Adviser",
        "Financial Advisors": "Adviser",
        "Vendors": "Divestor",
        "Grantors": "Awarding Authority"
    }

    # Prepare a list to store all entries before converting to DataFrame
    entries = []

    # Process each source column
    for source_column, role_type in sources.items():
        for _, row in transaction_df.iterrows():
            if pd.notna(row[source_column]):
                companies = row[source_column].split(';')  # Split by semi-colon
                for company in companies:
                    company = company.strip()
                    if company:  # Ensure company is not empty
                        entries.append({
                            "Transaction Upload ID": row["Transaction Upload ID"],
                            "Role Type": role_type,
                            "Role Subtype": "",
                            "Company": company,
                            "Fund": "",
                            "Bidder Status": "",
                            "Client Counterparty": "",
                            "Client Company Name": "",
                            "Fund Name": ""
                        })

    # Create DataFrame from list of dictionaries
    return pd.DataFrame(entries, columns=[
        "Transaction Upload ID", "Role Type", "Role Subtype", "Company", "Fund", 
        "Bidder Status", "Client Counterparty", "Client Company Name", "Fund Name"])

def process_tranches_sheet(transaction_df):
    # Prepare a list to store all entries before converting to DataFrame
    entries = []

    # Process tranches up to a maximum of 20
    for i in range(1, 21):
        tranche_type_column = f'Loan Debt Tranche {i} Type'
        tranche_tenor_column = f'Tranche {i} Tenor'
        tranche_value_column = f'Tranche {i} Volume USD (m)'
        
        # Check if the tranche columns exist in the dataframe
        if all(column in transaction_df.columns for column in [tranche_type_column, tranche_tenor_column, tranche_value_column]):
            for _, row in transaction_df.iterrows():
                if pd.notna(row[tranche_type_column]) or pd.notna(row[tranche_tenor_column]) or pd.notna(row[tranche_value_column]):
                    # Append data to entries
                    entries.append({
                        "Transaction Upload ID": row["Transaction Upload ID"],
                        "Tranche Upload ID": f'{row["Transaction Upload ID"]}-L{i}',
                        "Tranche Primary Type": "",  
                        "Tranche Secondary Type": "",  
                        "Tranche Tertiary Type": row.get(tranche_type_column, ""),
                        "Value": "",  
                        "Maturity Start Date": "",  
                        "Maturity End Date": "",  
                        "Tenor": row.get(tranche_tenor_column, ""),
                        "Tranche ESG Type": "",  
                        "Helper_Tranche Value USD m": row.get(tranche_value_column, "")
                    })

    # Create DataFrame from list of dictionaries
    tranches_df = pd.DataFrame(entries, columns=[
        "Transaction Upload ID", "Tranche Upload ID", "Tranche Primary Type", 
        "Tranche Secondary Type", "Tranche Tertiary Type", "Value", 
        "Maturity Start Date", "Maturity End Date", "Tenor", 
        "Tranche ESG Type", "Helper_Tranche Value USD m"])
    
    # Remove rows where 'Tranche Tertiary Type' is empty or invalid
    tranches_df = tranches_df[tranches_df["Tranche Tertiary Type"].astype(str).str.strip() != ""]

    return tranches_df

def populate_additional_tranches(transaction_df, tranches_df):
    # Process capital market tranches up to a maximum of 20
    for i in range(1, 21):
        cap_market_debt_column = f'Capital Market Debt {i} Volume USD (m)'
        
        if cap_market_debt_column in transaction_df.columns:
            for _, row in transaction_df.iterrows():
                if pd.notna(row[cap_market_debt_column]):
                    volume_usd = row[cap_market_debt_column]
                    transaction_upload_id = row["Transaction Upload ID"]
                    tranche_upload_id = f'{transaction_upload_id}-CM{i}'
                    
                    # Create a temporary DataFrame for the new tranche
                    temp_df = pd.DataFrame({
                        "Transaction Upload ID": [transaction_upload_id],
                        "Tranche Upload ID": [tranche_upload_id],
                        "Tranche Primary Type": [""],  
                        "Tranche Secondary Type": [""],  
                        "Tranche Tertiary Type": [""],  
                        "Value": [""],  
                        "Maturity Start Date": [""],  
                        "Maturity End Date": [""],  
                        "Tenor": [""],  
                        "Tranche ESG Type": [""],  
                        "Helper_Tranche Value USD m": [volume_usd]
                    })
                    
                    # Append the temporary DataFrame to the main tranches DataFrame
                    tranches_df = pd.concat([tranches_df, temp_df], ignore_index=True)

    # Remove rows where 'Helper_Tranche Value USD m' (Column K) is empty
    tranches_df = tranches_df[tranches_df["Helper_Tranche Value USD m"].astype(str).str.strip() != ""]
    
    return tranches_df

# Autofit columns
def autofit_columns(writer):
    for sheetname in writer.sheets:
        worksheet = writer.sheets[sheetname]
        for col in worksheet.columns:
            max_length = 0
            column = get_column_letter(col[0].column)  # Get the column name
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column].width = adjusted_width

def create_destination_file(source_file):
    # Load the source Excel file and automatically select the first sheet
    xls = pd.ExcelFile(source_file)
    first_sheet_name = xls.sheet_names[0]
    transaction_df = xls.parse(first_sheet_name)

    # Process each required sheet
    transaction_mapped_df = process_transaction_sheet(transaction_df)
    events_df = process_events_sheet(transaction_df)
    bidders_any_df = process_bidders_any_sheet(transaction_df)  # Process the 'Bidders_Any' sheet
    tranches_df = process_tranches_sheet(transaction_df)  # Process the 'Tranches' sheet
    
    # Populate additional tranches
    tranches_df = populate_additional_tranches(transaction_df, tranches_df)
    
    # Get the current date and time in London timezone
    london_tz = pytz.timezone('Europe/London')
    current_time = datetime.now(london_tz)
    formatted_time = current_time.strftime('%Y%m%d_%H%M')
    
    # Create destination file name
    destination_file_name = f'curated_INFRA3_{formatted_time}.xlsx'
    
    # Save to new Excel file
    with pd.ExcelWriter(destination_file_name, engine='openpyxl') as writer:
        transaction_mapped_df.to_excel(writer, sheet_name='Transaction', index=False)
        underlying_asset_df = pd.DataFrame(columns=["Transaction Upload ID", "Asset Upload ID"])
        underlying_asset_df.to_excel(writer, sheet_name='Underlying_Asset', index=False)
        events_df.to_excel(writer, sheet_name='Events', index=False)
        bidders_any_df.to_excel(writer, sheet_name='Bidders_Any', index=False)
        tranches_df.to_excel(writer, sheet_name='Tranches', index=False)
        tranche_pricings_df = pd.DataFrame(columns=[
            "Tranche Upload ID", "Tranche Benchmark", "Basis Point From", "Basis Point To", "Period From", "Period To", "Period Duration", "Comment"])
        tranche_pricings_df.to_excel(writer, sheet_name='Tranche_Pricings', index=False)
        tranche_roles_any_df = pd.DataFrame(columns=[
                    "Transaction Upload ID", "Tranche Upload ID", "Tranche Role Type",	"Company",	"Fund",	"Value", "Percentage", "Comment", "Helper_Tranche Primary Type", "Helper_Tranche Value $", "Helper_Transaction Value (USD m)", "Helper_LT Accredited Value ($m)", "Helper_Sponsor Equity USD m" ])
        tranche_roles_any_df.to_excel(writer, sheet_name='Tranche_Roles_Any', index=False)        

        # Autofit columns for all sheets
        autofit_columns(writer)
    
    return destination_file_name


# Streamlit app
st.title('Curating INFRA 3 Data Files')

uploaded_file = st.file_uploader("Choose a source file", type=["xlsx"])

if uploaded_file is not None:
    # Save the uploaded file to a temporary directory
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_file.write(uploaded_file.getbuffer())
    temp_file_path = temp_file.name
    temp_file.close()  # Ensure file is closed before processing

    try:
        with st.spinner("Processing the file..."):
            destination_path = create_destination_file(temp_file_path)
        st.success("File processed successfully!")

        # Provide a download button for the processed file
        with open(destination_path, "rb") as file:
            st.download_button(
                label="Download Processed File",
                data=file,
                file_name=os.path.basename(destination_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"An error occurred: {e}")

    finally:
        # Clean up temporary files
        try:
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
        except PermissionError:
            st.warning("Temporary file could not be deleted immediately, please try again later.")
        if os.path.exists(destination_path):
            os.remove(destination_path)

else:
    st.info("Please upload an Excel file to start processing.")
