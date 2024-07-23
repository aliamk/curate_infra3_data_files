import streamlit as st
import pandas as pd
import tempfile
import os
import re
from openpyxl.utils import get_column_letter
import pytz
from datetime import datetime

# Define a function to extract numerical value from a string globally
def extract_numerical_value(text):
    match = re.search(r'[\d,.]+', str(text))
    return match.group() if match else ''

def process_transaction_sheet(transaction_df):
    # Define a function to replace specific words in the 'Transaction Status' column
    def replace_transaction_status(status):
        replacements = {
            'Binding Bids': 'Preparation',
            'Expressions of Interest': 'Preparation',
            'Indicative Bids': 'Preparation',
            'No Private Financing': '',
            'On Hold': 'Preparation',
            'Preferred Proponent': 'Financing',
            'Pre-Launch': 'Preparation',
            'Pre-Qualified Proponents': 'Preparation',
            'RFP Returned': 'Preparation',
            'RFQ returned': 'Preparation',
            'Shortlisted Proponents': 'Preparation',
            'Transaction Launch': 'Preparation'
        }
        return replacements.get(status, status)
    
    # Define a function to replace specific words in the 'Transaction Type' column
    def replace_transaction_type(type):
        replacements = {
            'Additional Financing': 'Additional Financing',
            'Greenfield': 'Primary Financing',
            'M&A': 'Acquisition',
            'Nationalisation': '',
            'Privatisation': 'Privatisation',
            'Privatisation,M&A': 'Privatisation',
            'Public Offering': '',
            'Refinancing': 'Refinancing',
            'Take Private': ''
        }
        return replacements.get(type, type)
    
    # Define a function to replace specific words in the 'region - country' column
    def replace_region_country(region):
        replacements = {
            'AFGHANISTAN': 'Afghanistan',
            'ALBANIA': 'Albania',
            'ALGERIA': 'Algeria',
            'ANDORRA': 'Andorra',
            'ANGOLA': 'Angola',
            'ARGENTINA': 'Argentina',
            'ARMENIA': 'Armenia',
            'ARUBA': 'Aruba',
            'AUSTRALIA': 'Australia',
            'AUSTRIA': 'Austria',
            'AZERBAIJAN': 'Azerbaijan',
            'BAHAMAS': 'Bahamas',
            'BAHRAIN': 'Bahrain',
            'BANGLADESH': 'Bangladesh',
            'BARBADOS': 'Barbados',
            'BELARUS': 'Belarus',
            'BELGIUM': 'Belgium',
            'BENIN': 'Benin',
            'BERMUDA': 'Bermuda',
            'BOLIVIA': 'Bolivia',
            'BOSNIA': 'Bosnia & Herzegovina',
            'BOTSWANA': 'Botswana',
            'BRAZIL': 'Brazil',
            'BRUNEI': 'Brunei',
            'BULGARIA': 'Bulgaria',
            'BURKINA FASO': 'Burkina Faso',
            'BURUNDI': 'Burundi',
            'CAMBODIA': 'Cambodia',
            'CAMEROON': 'Cameroon',
            'CANADA': 'Canada',
            'CAPE VERDE': 'Cape Verde',
            'CAYMAN ISLANDS': 'Cayman Islands',
            'CHAD': 'Chad',
            'CHILE': 'Chile',
            'CHINA': 'China',
            'COLOMBIA': 'Colombia',
            'CONGO - REPUBLIC OF THE': 'Republic of the Congo',
            'COSTA RICA': 'Costa Rica',
            'CROATIA': 'Croatia',
            'CURACAO': 'Curaçao',
            'CYPRUS': 'Cyprus',
            'CZECH REPUBLIC': 'Czech Republic',
            'DENMARK': 'Denmark',
            'DJIBOUTI': 'Djibouti',
            'DOMINICAN REPUBLIC': 'Dominican Republic',
            'DR CONGO': 'Democratic Republic of Congo',
            'EAST TIMOR': 'Timor-Leste',
            'ECUADOR': 'Ecuador',
            'EGYPT': 'Egypt',
            'EL SALVADOR': 'El Salvador',
            'ESTONIA': 'Estonia',
            'ETHIOPIA': 'Ethiopia',
            'FINLAND': 'Finland',
            'FRANCE': 'France',
            'FRENCH GUIANA': 'French Guiana',
            'FRENCH POLYNESIA': 'French Polynesia',
            'GABON': 'Gabon',
            'GAMBIA': 'Gambia',
            'GEORGIA': 'Georgia',
            'GERMANY': 'Germany',
            'GHANA': 'Ghana',
            'GIBRALTAR': 'Gibraltar',
            'GREECE': 'Greece',
            'GUATEMALA': 'Guatemala',
            'GUINEA': 'Guinea',
            'GUYANA': 'Guyana',
            'HONDURAS': 'Honduras',
            'HONG KONG (CHINA)': 'Hong Kong',
            'HUNGARY': 'Hungary',
            'ICELAND': 'Iceland',
            'INDIA': 'India',
            'INDONESIA': 'Indonesia',
            'IRAQ': 'Iraq',
            'IRELAND': 'Ireland',
            'ISRAEL': 'Israel',
            'ITALY': 'Italy',
            'IVORY COAST': 'Ivory Coast',
            'JAMAICA': 'Jamaica',
            'JAPAN': 'Japan',
            'JORDAN': 'Jordan',
            'KAZAKHSTAN': 'Kazakhstan',
            'KENYA': 'Kenya',
            'KOSOVO': 'Kosovo',
            'KUWAIT': 'Kuwait',
            'KYRGYZSTAN': 'Kyrgyzstan',
            'LAOS': 'Laos',
            'LATVIA': 'Latvia',
            'LIBERIA': 'Liberia',
            'LIBYA': 'Libya',
            'LITHUANIA': 'Lithuania',
            'LUXEMBOURG': 'Luxembourg',
            'MADAGASCAR': 'Madagascar',
            'MALAWI': 'Malawi',
            'MALAYSIA': 'Malaysia',
            'MALDIVES': 'Maldives',
            'MALI': 'Mali',
            'MAURITIUS': 'Mauritius',
            'MEXICO': 'Mexico',
            'MOLDOVA': 'Moldova',
            'MONACO': 'Monaco',
            'MONGOLIA': 'Mongolia',
            'MONTENEGRO': 'Montenegro',
            'MONTSERRAT': 'Montserrat',
            'MOROCCO': 'Morocco',
            'MOZAMBIQUE': 'Mozambique',
            'MYANMAR': 'Myanmar',
            'NAMIBIA': 'Namibia',
            'NEPAL': 'Nepal',
            'NETHERLANDS': 'Netherlands',
            'NETHERLANDS ANTILLES': '',
            'NEW ZEALAND': 'New Zealand',
            'NICARAGUA': 'Nicaragua',
            'NIGER': 'Niger',
            'NIGERIA': 'Nigeria',
            'NORTH MACEDONIA': 'North Macedonia',
            'NORWAY': 'Norway',
            'OMAN': 'Oman',
            'PAKISTAN': 'Pakistan',
            'PALESTINE': 'Palestine',
            'PANAMA': 'Panama',
            'PAPUA NEW GUINEA': 'Papua New Guinea',
            'PARAGUAY': 'Paraguay',
            'PERU': 'Peru',
            'PHILIPPINES': 'Philippines',
            'POLAND': 'Poland',
            'PORTUGAL': 'Portugal',
            'QATAR': 'Qatar',
            'REUNION': 'Reunion',
            'ROMANIA': 'Romania',
            'RUSSIA': 'Russia',
            'RWANDA': 'Rwanda',
            'SAUDI ARABIA': 'Saudi Arabia',
            'SENEGAL': 'Senegal',
            'SERBIA': 'Serbia',
            'SEYCHELLES': 'Seychelles',
            'SINGAPORE': 'Singapore',
            'SLOVAKIA': 'Slovakia',
            'SLOVENIA': 'Slovenia',
            'SOUTH AFRICA': 'South Africa',
            'SOUTH KOREA': 'South Korea',
            'SPAIN': 'Spain',
            'SRI LANKA': 'Sri Lanka',
            'SWEDEN': 'Sweden',
            'SWITZERLAND': 'Switzerland',
            'SYRIA': 'Syria',
            'TAIWAN (CHINA)': 'Taiwan',
            'TAJIKISTAN': 'Tajikistan',
            'TANZANIA': 'Tanzania',
            'THAILAND': 'Thailand',
            'TOGO': 'Togo',
            'TRINIDAD & TOBAGO': 'Trinidad and Tobago',
            'TUNISIA': 'Tunisia',
            'TURKEY': 'Turkey',
            'UGANDA': 'Uganda',
            'UKRAINE': 'Ukraine',
            'UNITED ARAB EMIRATES': 'United Arab Emirates',
            'UNITED KINGDOM': 'United Kingdom',
            'URUGUAY': 'Uruguay',
            'USA': 'United States',
            'UZBEKISTAN': 'Uzbekistan',
            'VIETNAM': 'Vietnam',
            'VIRGIN ISLANDS (US)': 'US Virgin Islands',
            'ZAMBIA': 'Zambia',
            'ZIMBABWE': 'Zimbabwe'
        }
        return replacements.get(region, region)
    
    # Define a function to replace specific words in the 'contract' column
    def replace_contract(contract):
        replacements = {
            'DBFOM': 'DBFOM',
            'DBFM': 'DBFM',
            'DBFO': 'DBFO',
            'DBF': 'DBF',
            'BF': '',
            'BFOM': '',
            'DBOM': '',
            'BFO': '',
            'BO': '',
            'OM': '',
            'DBO': '',
            'DB': '',
            'FOM': '',
            'BOM': '',
            'DFOM': '',
            'DBM': '',
            'BM': '',
            'DOM': '',
            'DO': '',
            'DFO': '',
            'O': '',
        }
        return replacements.get(contract, contract)

    # Define a function to replace words based on the replacement list for 'Any Level Sectors'
    def populate_any_level_sectors(sectors):
        replacements = {
            'Accommodation': 'Social Infrastructure',
            'Airports': 'Transport, Airport',
            'Battery Storage': 'Renewable Energy, Energy Storage',
            'Biofuels': 'Renewable Energy, Biofuels/Biomass',
            'Biogas': 'Renewable Energy, Biofuels/Biomass',
            'Biomass': 'Renewable Energy, Biofuels/Biomass',
            'Bridges and Tunnels': 'Transport',
            'Broadband': 'Digital Infrastructure, Internet',
            'Car Parks': 'Transport, Car Park',
            'Carbon Capture': 'Renewable Energy, Carbon Capture & Storage',
            'Coal fired': 'Conventional Energy, Coal-Fired Power',
            'Co-generation': 'Conventional Energy, Cogeneration Power',
            'Courthouses': 'Social Infrastructure, Justice',
            'Data Centre': 'Digital Infrastructure, Data Centre',
            'Defence': 'Social Infrastructure',
            'Desalination': 'Water, Desalination',
            'District Heating & Cooling': 'Social Infrastructure, Heat Network',
            'Education': 'Social Infrastructure, Education',
            'Electricity Distribution': 'Conventional Energy, Transmission',
            'Electricity Smart Meter': 'Conventional Energy, Transmission',
            'Electricity Transmission': 'Conventional Energy, Transmission',
            'Energy from waste': 'Renewable Energy, Waste to Energy',
            'Energy Other': 'Conventional Energy',
            'EV Infrastructure': 'Renewable Energy, EV Charging',
            'Exploration & Production': 'Oil & Gas, Upstream',
            'Ferries': 'Transport, Waterway',
            'Fibre Optic': 'Digital Infrastructure, Internet',
            'Floating Solar PV': 'Renewable Energy, Solar (Floating PV)',
            'Gas Distribution': 'Oil & Gas, Downstream',
            'Gas fired': 'Conventional Energy, Gas-Fired Power',
            'Gas Pipeline': 'Oil & Gas, Midstream',
            'Gas Smart Meter': 'Conventional Energy',
            'Geothermal': 'Renewable Energy, Geothermal',
            'Healthcare': 'Social Infrastructure, Healthcare',
            'High-speed Rail': 'Transport, Heavy Rail',
            'Hydroelectric': 'Renewable Energy, Hydro',
            'Hydrogen': 'Renewable Energy, Hydrogen',
            'IWPP': 'Conventional Energy',
            'Leisure': 'Social Infrastructure, Leisure',
            'LNG export terminal': 'Oil & Gas, LNG',
            'Microgrids': 'Conventional Energy, Transmission',
            'Mining': 'Mining',
            'Nuclear': 'Conventional Energy, Nuclear Power',
            'Offshore wind': 'Renewable Energy, Wind (Offshore)',
            'Oil & Gas Storage': 'Oil & Gas, Midstream',
            'Oil & gas transportation': 'Oil & Gas, Midstream',
            'Oil fired': 'Conventional Energy, Oil-Fired Power',
            'Oil Pipeline': 'Oil & Gas, Midstream',
            'Onshore wind': 'Renewable Energy, Wind (Onshore)',
            'Petrochemical plants': 'Oil & Gas, Petrochemical',
            'Police Facilities': 'Social Infrastructure, Justice',
            'Ports': 'Transport, Port',
            'Power Other': 'Conventional Energy',
            'Prisons': 'Social Infrastructure, Justice',
            'Rail': 'Transport, Heavy Rail',
            'Refineries': 'Oil & Gas',
            'Renewables Other': 'Renewable Energy',
            'Roads': 'Transport, Road',
            'Rolling Stock': 'Transport, Heavy Rail',
            'Social Housing': 'Social Infrastructure, Social Housing',
            'Social Infrastructure Other': 'Social Infrastructure',
            'Solar CSP': 'Renewable Energy, Solar (Thermal)',
            'Solar PV': 'Renewable Energy, Solar (Land-Based Solar)',
            'Subsea Cable': 'Digital Infrastructure',
            'Telecommunications Other': 'Digital Infrastructure',
            'Tidal': 'Renewable Energy, Marine',
            'Transport Other': 'Transport',
            'Urban Rail Transit': 'Transport, Light Transport',
            'Waste': 'Waste',
            'Water': 'Water',
            'Wireless Transmission': 'Digital Infrastructure'
        }
        sector_list = [sector.strip() for sector in sectors.split(',')]
        replacement_list = []
        for sector in sector_list:
            if sector in replacements:
                replacement_list.append(replacements[sector])
            else:
                replacement_list.append(sector)
        return ', '.join(replacement_list)

    # Ensure 'Helper_Any Level Sectors' exists
    if 'Helper_Any Level Sectors' not in transaction_df.columns:
        transaction_df['Helper_Any Level Sectors'] = ''

    # Map columns for the 'Transaction' sheet with transformations
    transaction_df["Helper_Any Level Sectors"] = transaction_df["Helper_Any Level Sectors"].fillna('')
    return pd.DataFrame({
        "Transaction Upload ID": transaction_df["Transaction Upload ID"],
        "Transaction Name": transaction_df["Transaction Name"],
        "Transaction Asset Class": "Infrastructure",  # Static value for all rows
        "Transaction Status": transaction_df.get("Current status", "").apply(replace_transaction_status),  # Using .get to avoid KeyError if not present
        "Finance Type": "",
        "Transaction Type": transaction_df.get("Type", "").apply(replace_transaction_type),
        "Unknown Asset": "",
        "Underlying Asset Configuration": "",
        "Transaction Local Currency": transaction_df["Transaction Currency"].apply(extract_numerical_value),  # Extract numerical value
        "Transaction Value (Local Currency)": transaction_df.get("Transaction size (m)", ""),
        "Transaction Debt (Local Currency)": "",
        "Transaction Equity (Local Currency)": "",
        "Debt/Equity Ratio": "",
        "Underlying Number of Assets": "",
        "Region - Country": transaction_df.get("Geography", "").apply(replace_region_country),
        "Region - State": "",
        "Region - City": "",
        "Any Level Sectors": transaction_df.get("Sector", "") + ", " + transaction_df.get("Sub-Sector", "").apply(populate_any_level_sectors),
        "PPP": transaction_df.get("PPP", ""),
        "Concession Period": transaction_df.get("Duration", ""),
        "Contract": transaction_df.get("Delivery Model", "").apply(replace_contract),
        "SPV": transaction_df.get("SPV", ""),
        "Active": "TRUE",
        "Helper_Any Level Sectors": transaction_df.get("Sector", "") + ", " + transaction_df.get("Sub-Sector", "")
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
    
    # Define a function to replace specific words in the 'Event Type' column
    def replace_event_type(eventtype):
        replacements = {
            'Binding Bids': '',
            'Cancelled': 'Cancelled',
            'Expressions of Interest': 'Expression of Interest',
            'Financial Close': 'Financial Close',
            'Indicative Bids': '',
            'No Private Financing': '',
            'On Hold': '',
            'Preferred Proponent': 'Preferred Bidder',
            'Pre-Launch': '',
            'Pre-Qualified Proponents': '',
            'RFP Returned': 'Request for Proposals',
            'RFQ returned': 'Request for Qualifications',
            'Shortlisted Proponents': 'Shortlist',
            'Transaction Launch': 'Announced',
        }
        return replacements.get(eventtype, eventtype)
    
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
            "Event Type": event_type_series.apply(replace_event_type),
            "Event Title": ""  # Assuming this field is empty or to be populated later
        })

        # Append the temporary DataFrame to the main events DataFrame
        events_df = pd.concat([events_df, temp_df], ignore_index=True)

    # Remove rows with blank or "N/A" or "n/a" in 'Event Date'
    events_df = events_df[~events_df["Event Date"].isin(["", "N/A", "n/a"])]

    # Remove duplicate rows
    events_df = events_df.drop_duplicates()

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
                        # Determine Client Counterparty based on the content in parentheses
                        client_counterparty = ''
                        if '(Funders)' in company:
                            client_counterparty = 'Debt Provider'
                        elif '(Acquirer)' in company or '(Acquiror)' in company:
                            client_counterparty = 'Acquirer'
                        elif '(SPV)' in company:
                            client_counterparty = 'SPV'
                        elif '(Seller)' in company:
                            client_counterparty = 'Divestor'
                        elif '(Grantor)' in company:
                            client_counterparty = 'Awarding Authority'
                        elif '(Target)' in company or '(Target Company)' in company:
                            client_counterparty = 'Target'
                        elif '(Lenders)' in company:
                            client_counterparty = 'Debt Provider'

                        # Remove parentheses and their content from the company name
                        company_cleaned = re.sub(r'\s*\(.*?\)\s*', '', company).strip()

                        entries.append({
                            "Transaction Upload ID": row["Transaction Upload ID"],
                            "Role Type": role_type,
                            "Role Subtype": "",
                            "Company": company_cleaned,
                            "Fund": "",
                            "Bidder Status": "Successful",
                            "Client Counterparty": client_counterparty,
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
                        "Tranche ESG Type": row.get("Tranche ESG Type", ""),  
                        "Helper_Tranche Value USD m": row.get(tranche_value_column, ""),
                        "Helper_Transaction Value USD m": row.get("Transaction size USD(m)", ""),
                        "Helper_Transaction Value LC": row.get("Transaction size (m)", "")
                    })

    # Create DataFrame from list of dictionaries
    tranches_df = pd.DataFrame(entries, columns=[
        "Transaction Upload ID", "Tranche Upload ID", "Tranche Primary Type", 
        "Tranche Secondary Type", "Tranche Tertiary Type", "Value", 
        "Maturity Start Date", "Maturity End Date", "Tenor", 
        "Tranche ESG Type", "Helper_Tranche Value USD m",
        "Helper_Transaction Value USD m", "Helper_Transaction Value LC"])
    
    # Helper function to safely convert values to float
    def safe_float_conversion(value):
        if isinstance(value, str):
            value = value.replace(',', '').strip()
        try:
            return float(value)
        except ValueError:
            return 0

    # Add the new column 'Helper_Tranche Value USD m as % of Helper_Transaction Value USD m'
    tranches_df["Helper_Tranche Value USD m as % of Helper_Transaction Value USD m"] = tranches_df.apply(
    lambda row: safe_float_conversion(row["Helper_Tranche Value USD m"]) / safe_float_conversion(row["Helper_Transaction Value USD m"]) 
    if safe_float_conversion(row["Helper_Transaction Value USD m"]) != 0 else 0, axis=1)

    # Populate column F "Value" with results of multiplying columns "Helper_Tranche Value USD m as % of Helper_Transaction Value USD m" by "Helper_Transaction Value LC"
    tranches_df["Value"] = tranches_df.apply(
    lambda row: safe_float_conversion(row["Helper_Tranche Value USD m as % of Helper_Transaction Value USD m"]) * safe_float_conversion(row["Helper_Transaction Value LC"]) 
    if safe_float_conversion(row["Helper_Tranche Value USD m as % of Helper_Transaction Value USD m"]) and safe_float_conversion(row["Helper_Transaction Value LC"]) else 0, axis=1)
    
    # Update 'Tranche ESG Type' if 'Tranche Tertiary Type' contains 'Islamic'
    tranches_df["Tranche ESG Type"] = tranches_df.apply(
        lambda row: f'{row["Tranche ESG Type"]}, Tranche ESG Type' if "Islamic" in row["Tranche Tertiary Type"] else row["Tranche ESG Type"],
        axis=1
    )

    # Replace words in 'Tranche Tertiary Type' based on the provided list
    replacements = {
        'Capex Facility': '',
        'Change-in-Law Facility': '',
        'Equity Bridge Loan': '',
        'Export Credit': 'Export Credit Facility',
        'Government Grant': '',
        'Government Loan': 'State Loan',
        'Islamic Financing': 'Term Loan',
        'Multilateral': 'Multilateral Loan',
        'Other': '',
        'Standby/Contigency Facility': 'Standby Facility'
    }

    tranches_df["Tranche Tertiary Type"] = tranches_df["Tranche Tertiary Type"].replace(replacements)

    
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
                        "Helper_Tranche Value USD m": [volume_usd],
                        "Helper_Transaction Value USD m": [row.get("Transaction size USD(m)", "")],
                        "Helper_Transaction Value LC": [row.get("Transaction size (m)", "")]
                    })
                    
                    # Append the temporary DataFrame to the main tranches_df DataFrame
                    tranches_df = pd.concat([tranches_df, temp_df], ignore_index=True)
    
    # Append additional data based on 'Equity Providers at FC'
    if 'Equity Providers at FC' in transaction_df.columns:
        equity_providers_df = transaction_df.dropna(subset=['Equity Providers at FC'])
        for _, row in equity_providers_df.iterrows():
            transaction_upload_id = row["Transaction Upload ID"]
            tranche_upload_id = f'{transaction_upload_id}-E'
            equity_value = extract_numerical_value(row.get('Equity at FC USD(m)', ''))

            temp_df = pd.DataFrame({
                "Transaction Upload ID": [transaction_upload_id],
                "Tranche Upload ID": [tranche_upload_id],
                "Tranche Primary Type": [""],
                "Tranche Secondary Type": [""],
                "Tranche Tertiary Type": ["Equity"],
                "Value": [""],
                "Maturity Start Date": [""],
                "Maturity End Date": [""],
                "Tenor": [""],
                "Tranche ESG Type": [""],
                "Helper_Tranche Value USD m": [equity_value],
                "Helper_Transaction Value USD m": [row.get("Transaction size USD(m)", "")],
                "Helper_Transaction Value LC": [row.get("Transaction size (m)", "")]
            })
            
            tranches_df = pd.concat([tranches_df, temp_df], ignore_index=True)
    
    # Remove rows where 'Tranche Tertiary Type' is empty and 'Tranche Upload ID' includes "-L1" to "-L20"
    tranches_df = tranches_df[~((tranches_df['Tranche Tertiary Type'].astype(str).str.strip() == '') & 
                                (tranches_df['Tranche Upload ID'].str.contains('-L[1-9]$|-L1[0-9]$|-L20$', regex=True)))]

    return tranches_df

def populate_tranche_roles_any(transaction_df, tranche_roles_any_df):
    entries = []

    # Process 'Tranche 1 Lenders' to 'Tranche 20 Lenders' first
    for i in range(1, 21):
        lenders_column = f'Tranche {i} Lenders'

        if lenders_column in transaction_df.columns:
            for _, row in transaction_df.iterrows():
                if pd.notna(row[lenders_column]):
                    lenders = re.split(r',\s*(?![^()]*\))', row[lenders_column])  # Split by comma unless within parentheses
                    for lender in lenders:
                        lender = lender.strip()
                        percentage = ''
                        match = re.search(r'(\d+)%\)', lender)
                        if match:
                            percentage = match.group(1)
                        if lender:
                            entries.append({
                                "Transaction Upload ID": row["Transaction Upload ID"],
                                "Tranche Upload ID": f'{row["Transaction Upload ID"]}-L{i}',
                                "Tranche Role Type": "",
                                "Company": lender,
                                "Fund": "",
                                "Value": "",
                                "Percentage": percentage,
                                "Comment": ""
                            })

    # Process 'Capital Market Debt 1 Underwriters' to 'Capital Market Debt 20 Underwriters' next
    for i in range(1, 21):
        cm1_column = f'Capital Market Debt {i} Underwriters'
        cm2_column = f'Capital Market Debt 2{i} Underwriters'

        if cm1_column in transaction_df.columns:
            for _, row in transaction_df.iterrows():
                if pd.notna(row[cm1_column]):
                    underwriters_cm1 = re.split(r',\s*(?![^()]*\))', row[cm1_column])  # Split by comma unless within parentheses
                    for underwriter in underwriters_cm1:
                        underwriter = underwriter.strip()
                        percentage = ''
                        match = re.search(r'(\d+)%\)', underwriter)
                        if match:
                            percentage = match.group(1)
                        if underwriter:
                            entries.append({
                                "Transaction Upload ID": row["Transaction Upload ID"],
                                "Tranche Upload ID": f'{row["Transaction Upload ID"]}-CM{i}',
                                "Tranche Role Type": "",
                                "Company": underwriter,
                                "Fund": "",
                                "Value": "",
                                "Percentage": percentage,
                                "Comment": ""
                            })

        if cm2_column in transaction_df.columns:
            for _, row in transaction_df.iterrows():
                if pd.notna(row[cm2_column]):
                    underwriters_cm2 = re.split(r',\s*(?![^()]*\))', row[cm2_column])  # Split by comma unless within parentheses
                    for underwriter in underwriters_cm2:
                        underwriter = underwriter.strip()
                        percentage = ''
                        match = re.search(r'(\d+)%\)', underwriter)
                        if match:
                            percentage = match.group(1)
                        if underwriter:
                            entries.append({
                                "Transaction Upload ID": row["Transaction Upload ID"],
                                "Tranche Upload ID": f'{row["Transaction Upload ID"]}-CM2{i}',
                                "Tranche Role Type": "",
                                "Company": underwriter,
                                "Fund": "",
                                "Value": "",
                                "Percentage": percentage,
                                "Comment": ""
                            })

    # Append additional data based on 'Equity Providers at FC'
    if 'Equity Providers at FC' in transaction_df.columns:
        equity_providers_df = transaction_df.dropna(subset=['Equity Providers at FC'])
        for _, row in equity_providers_df.iterrows():
            equity_providers = re.split(r',\s*(?![^()]*\))', row['Equity Providers at FC'])  # Split by comma unless within parentheses
            for provider in equity_providers:
                provider = provider.strip()
                percentage = ''
                match = re.search(r'(\d+)%\)', provider)
                if match:
                    percentage = match.group(1)
                if provider:
                    entries.append({
                        "Transaction Upload ID": row["Transaction Upload ID"],
                        "Tranche Upload ID": f'{row["Transaction Upload ID"]}-E',
                        "Tranche Role Type": "",
                        "Company": provider,
                        "Fund": "",
                        "Value": "",
                        "Percentage": percentage,
                        "Comment": ""
                    })

    # Create DataFrame from list of dictionaries
    tranche_roles_any_df = pd.DataFrame(entries, columns=[
        "Transaction Upload ID", "Tranche Upload ID", "Tranche Role Type", "Company", "Fund", 
        "Value", "Percentage", "Comment"
    ])

    return tranche_roles_any_df

def clean_company_names(tranche_roles_any_df):
    def clean_company_name(name):
        # a) Delete content within parenthesis and delete parenthesis
        name = re.sub(r'\s*\(.*?\)\s*', '', name)
        # b) Delete all trailing spaces
        name = name.strip()
        # c) Delete two or more spaces in between words
        name = re.sub(r'\s{2,}', ' ', name)
        return name
    
    tranche_roles_any_df['Company'] = tranche_roles_any_df['Company'].apply(clean_company_name)
    return tranche_roles_any_df

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

# Clean up transaction names
def clean_transaction_name(df):
    df['Transaction Name'] = df['Transaction Name'].str.strip()  # Remove leading/trailing spaces
    df['Transaction Name'] = df['Transaction Name'].replace(r'\s+', ' ', regex=True)  # Replace multiple spaces with single space
    return df

# Replace " and " with " & "
def replace_and_with_ampersand(df):
    df['Transaction Name'] = df['Transaction Name'].str.replace(' and ', ' & ')
    return df

def create_destination_file(source_file):
    # Load the source Excel file and automatically select the first sheet
    xls = pd.ExcelFile(source_file)
    first_sheet_name = xls.sheet_names[0]
    transaction_df = xls.parse(first_sheet_name)

    # Process each required sheet
    transaction_mapped_df = process_transaction_sheet(transaction_df)
    transaction_mapped_df = clean_transaction_name(transaction_mapped_df)  # Clean transaction names
    transaction_mapped_df = replace_and_with_ampersand(transaction_mapped_df)  # Replace " and " with " & "
    events_df = process_events_sheet(transaction_df)
    bidders_any_df = process_bidders_any_sheet(transaction_df)  # Process the 'Bidders_Any' sheet
    tranches_df = process_tranches_sheet(transaction_df)  # Process the 'Tranches' sheet
    
    # Populate additional tranches
    tranches_df = populate_additional_tranches(transaction_df, tranches_df)
    
    # Update 'Tranche Primary Type', 'Tranche Secondary Type' and 'Tranche Tertiary Type' based on 'Tranche Upload ID'
    tranches_df['Tranche Primary Type'] = tranches_df['Tranche Upload ID'].apply(
        lambda x: 'Debt' if any(x.endswith(suffix) for suffix in ['L1', 'L2', 'L3', 'CM1', 'CM2', 'CM3']) else 'Equity'
    )
    tranches_df['Tranche Secondary Type'] = tranches_df['Tranche Upload ID'].apply(
        lambda x: 'Loan' if any(x.endswith(suffix) for suffix in ['L1', 'L2', 'L3']) else ('Bond' if any(x.endswith(suffix) for suffix in ['CM1', 'CM2', 'CM3']) else 'Equity')
    )
    tranches_df['Tranche Tertiary Type'] = tranches_df.apply(
        lambda row: 'Commercial Bond' if any(row['Tranche Upload ID'].endswith(suffix) for suffix in ['CM1', 'CM2', 'CM3']) else row['Tranche Tertiary Type'],
        axis=1
    )    

    # Get the current date and time in London timezone
    london_tz = pytz.timezone('Europe/London')
    current_time = datetime.now(london_tz)
    formatted_time = current_time.strftime('%Y%m%d_%H%M')
    
    # Create destination file name
    destination_file_name = f'curated_INFRA3_{formatted_time}.xlsx'
    
    # Populate tranche roles
    tranche_roles_any_df = pd.DataFrame(columns=[
        "Transaction Upload ID", "Tranche Upload ID", "Role Type", "Company", "Fund", 
        "Value", "Percentage", "Comment"])
    tranche_roles_any_df = populate_tranche_roles_any(transaction_df, tranche_roles_any_df)

    # New logic to update 'Tranche Role Type' based on conditions
    for i, row in tranche_roles_any_df.iterrows():
        tranche_upload_id = row['Tranche Upload ID']
        tranche_info = tranches_df[tranches_df['Tranche Upload ID'] == tranche_upload_id]
        if not tranche_info.empty:
            if tranche_info.iloc[0]['Tranche Primary Type'] == 'Equity':
                tranche_roles_any_df.at[i, 'Tranche Role Type'] = 'Sponsor'
            elif tranche_info.iloc[0]['Tranche Secondary Type'] == 'Bond':
                tranche_roles_any_df.at[i, 'Tranche Role Type'] = 'Bond Arranger'
            elif tranche_info.iloc[0]['Tranche Secondary Type'] == 'Loan':
                tranche_roles_any_df.at[i, 'Tranche Role Type'] = 'Debt Provider'
            elif (tranche_info.iloc[0]['Tranche Primary Type'] == 'Debt' and 
                  tranche_info.iloc[0]['Tranche Secondary Type'] == 'Non-Commercial Instrument'):
                tranche_roles_any_df.at[i, 'Tranche Role Type'] = 'Debt Provider'

    # Clean the 'Company' column in the 'Tranche_Roles_Any' tab
    tranche_roles_any_df = clean_company_names(tranche_roles_any_df)
    
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

    destination_path = None  # Initialize destination_path

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
        if destination_path and os.path.exists(destination_path):
            os.remove(destination_path)

else:
    st.info("Please upload an Excel file to start processing.")