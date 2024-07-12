import streamlit as st
import pandas as pd
import tempfile
import os
import openpyxl

def process_transaction_sheet(transaction_df):
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
        "Transaction Local Currency": transaction_df.get("Transaction Currency", ""),
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




def create_destination_file(source_file):
    # Load the source Excel file and automatically select the first sheet
    xls = pd.ExcelFile(source_file)
    first_sheet_name = xls.sheet_names[0]
    transaction_df = xls.parse(first_sheet_name)

    # Process each required sheet
    transaction_mapped_df = process_transaction_sheet(transaction_df)
    events_df = process_events_sheet(transaction_df)
    bidders_any_df = process_bidders_any_sheet(transaction_df)  # Process the 'Bidders_Any' sheet
    
    # Save to new Excel file
    with pd.ExcelWriter('processed_file.xlsx', engine='openpyxl') as writer:
        transaction_mapped_df.to_excel(writer, sheet_name='Transaction', index=False)
        events_df.to_excel(writer, sheet_name='Events', index=False)
        bidders_any_df.to_excel(writer, sheet_name='Bidders_Any', index=False)
    
    return 'processed_file.xlsx'


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
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)
        if os.path.exists(destination_path):
            os.remove(destination_path)

else:
    st.info("Please upload an Excel file to start processing.")
