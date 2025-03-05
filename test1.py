import streamlit as st
import sys
import subprocess

# Attempt to install dependencies if not present
def install_dependencies():
    """Attempt to install required dependencies"""
    dependencies = ['openpyxl', 'pandas']
    for dep in dependencies:
        try:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', dep])
        except Exception as e:
            st.error(f"Error installing {dep}: {e}")
            return False
    return True

# Check and install dependencies
def check_dependencies():
    """Check and install missing dependencies"""
    try:
        import openpyxl
        import pandas
        return True
    except ImportError:
        st.warning("Missing required dependencies. Attempting to install...")
        return install_dependencies()

# Only proceed if dependencies are installed
if not check_dependencies():
    st.error("Could not install required dependencies. Please check your internet connection.")
    st.stop()

# Now import other required libraries
import os
import pandas as pd
from datetime import datetime
import base64
import uuid

SURVEY_RESULTS_DIR = "survey_results"

def flatten_nested_dict(nested_dict, parent_key='', sep='_'):
    """
    Flatten a nested dictionary with comprehensive handling.
    Converts None or empty values to 'N/A' for better Excel representation.
    """
    items = []
    for k, v in nested_dict.items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        
        if isinstance(v, dict):
            # Recursively flatten nested dictionaries
            items.extend(flatten_nested_dict(v, new_key, sep=sep).items())
        else:
            # Convert None or empty values to 'N/A'
            if v is None or (isinstance(v, str) and v.strip() == ''):
                v = 'N/A'
            items.append((new_key, v))
    
    return dict(items)

def sanitize_filename(filename):
    """
    Sanitize the filename to remove any invalid characters
    """
    import re
    return re.sub(r'[<>:"/\\|?*]', '_', filename).strip()

def save_survey_data(survey_data):
    """
    Save survey responses to an Excel file with comprehensive data handling.
    Each survey creates a unique file named after the ashram.
    """
    # Ensure results directory exists
    os.makedirs(SURVEY_RESULTS_DIR, exist_ok=True)
    
    # Add timestamp
    survey_data['timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Get ashram name, use a default if not provided
    ashram_name = survey_data.get('ashram_name', 'Unknown_Ashram')
    
    # Sanitize ashram name for filename
    safe_ashram_name = sanitize_filename(ashram_name)
    
    # Create a unique filename with timestamp to prevent overwriting
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{safe_ashram_name}_{timestamp}.xlsx"
    file_path = os.path.join(SURVEY_RESULTS_DIR, filename)

    try:
        # Flatten the survey data
        flat_entry = flatten_nested_dict(survey_data)

        # Create DataFrame
        df = pd.DataFrame([flat_entry])

        # Reorder columns to make it more readable
        column_order = ['timestamp']
        
        # Add sections in a specific order
        sections = [
            'ashram_name',
            'Property Ownership and Legal Documents',
            'Trust/Society Details & Documents',
            'Institutions Details & Documents'
        ]
        
        for section in sections:
            section_columns = [col for col in df.columns if section in col]
            column_order.extend(sorted(section_columns))

        # Reorder columns, adding any remaining columns at the end
        remaining_columns = [col for col in df.columns if col not in column_order]
        column_order.extend(remaining_columns)

        # Reorder DataFrame
        df = df[column_order]

        # Save to Excel
        df.to_excel(file_path, index=False)
        
        # Create download link
        with open(file_path, 'rb') as f:
            bytes_data = f.read()
        b64 = base64.b64encode(bytes_data).decode()
        
        # Display success message with download link
        download_link = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download {filename}</a>'
        st.markdown(f"Survey saved successfully! {download_link}", unsafe_allow_html=True)
        
        return True
    except Exception as e:
        st.error(f"Error saving survey: {e}")
        return False

def start_page():
    """Start page with a title and a start button."""
    # Center the main title with the largest font size
    st.markdown("<h1 style='text-align: center; font-size: 3em;'>ORDER OF THE IMITATION OF CHRIST</h1>", unsafe_allow_html=True)
    
    # Subheading, centered with slightly smaller font
    st.markdown("<h2 style='text-align: center; font-size: 2em;'>BETHANY NAVAJYOTHY PROVINCE</h2>", unsafe_allow_html=True)
    
    # Document verification text centered
    st.markdown("<h3 style='text-align: center;'>Document Verification</h3>", unsafe_allow_html=True)
    
    # Center the button
    col1, col2, col3 = st.columns([3,2,3])
    with col2:
        if st.button("Start Survey"):
            st.session_state["survey_started"] = True
            st.rerun()

def input_section(label, key, add_comment=True, add_upload=False):
    """Creates an input section with optional upload and comment boxes."""
    st.markdown(f"<h4 style='color: yellow;'>{label}</h4>", unsafe_allow_html=True)
    response = {}
    
    col1, col2 = st.columns([3,1])  # Create two columns
    
    with col1:
        response["number_of_documents"] = st.text_input("Number of Documents", key=key)
    
    with col2:
        # Small file upload
        uploaded_file = st.file_uploader("📄", key=f"{key}_upload", label_visibility="collapsed")
        if uploaded_file is not None:
            response["uploaded_file_name"] = uploaded_file.name
            # Optional: You could save the file or process it here
    
    if add_comment:
        response["additional_comments"] = st.text_area("Additional comments", key=f"{key}_comment", height=100)
    
    st.markdown("---", unsafe_allow_html=True)
    return response

def property_survey():
    """Survey page for Property Ownership and Legal Documents."""
    st.title("Property Ownership and Legal Documents")

    responses = {}
    
    # New first question for Ashram Name
    responses["ashram_name"] = st.text_input("Name of the Ashram", key="ashram_name")
    st.markdown("---")  # Add a divider after the name input
    
    responses["land_title"] = input_section("Land title deed (original and certified copies)", "land_title", add_upload=True)
    responses["property_tax"] = input_section("Property tax receipts (latest)", "property_tax", add_upload=True)
    responses["land_survey"] = input_section("Land survey records and maps", "land_survey", add_upload=True)
    responses["mutation_cert"] = input_section("Mutation Certificate (Pokkuvaravu)", "mutation_cert", add_upload=True)
    responses["land_na"] = input_section("Land N A Documents", "land_na", add_upload=True)
    responses["registration_details"] = input_section("Registration details of the property (if under a trust or society)", "registration_details", add_upload=True)

    st.subheader("Building and Construction Approvals")
    responses["building_permit"] = input_section("Building permit and approval documents from local authorities", "building_permit", add_upload=True)
    responses["occupancy_cert"] = input_section("Occupancy certificate (if applicable)", "occupancy_cert", add_upload=True)
    responses["approved_plans"] = input_section("Approved building plans and layout", "approved_plans", add_upload=True)
    responses["renovation_approvals"] = input_section("Any renovation or expansion approvals", "renovation_approvals", add_upload=True)
    
    return responses

def trust_society_survey():
    """Survey page for Trust/Society Details & Documents with file upload options."""
    st.title("Trust/Society Details & Documents")

    responses = {}
    
    responses["trust_deed"] = input_section("Original copy of Trust deed", "trust_deed", add_comment=False, add_upload=True)
    responses["minority_certificate"] = input_section("Minority Certificate", "minority_certificate", add_comment=False, add_upload=True)
    responses["deed_society_cert"] = input_section("Deed & Society Certificate", "deed_society_cert", add_comment=False, add_upload=True)
    responses["society_renewal"] = input_section("Society Renewal receipt", "society_renewal", add_comment=False, add_upload=True)
    responses["noc_government"] = input_section("NOC from Government", "noc_government", add_comment=False, add_upload=True)
    responses["pan_card"] = input_section("Pan card", "pan_card", add_comment=False, add_upload=True)
    responses["amc_mou"] = input_section("AMC/MOU if any (Refer Educational Manual)", "amc_mou", add_comment=False, add_upload=True)
    responses["income_tax"] = input_section("Income Tax Records", "income_tax", add_comment=False, add_upload=True)
    responses["bank_kyc_resolution"] = input_section("Resolution prepared for Bank KYC & list of office bearers", "bank_kyc_resolution", add_comment=False, add_upload=True)
    responses["darpan_cert"] = input_section("Darpan - 10 AC - ATG certificate if any", "darpan_cert", add_comment=False, add_upload=True)
    responses["trust_meeting_reports"] = input_section("Reports & Minutes of the Trust Meetings", "trust_meeting_reports", add_comment=False, add_upload=True)
    
    return responses

def institution_survey():
    """Survey page for Institutions Details & Documents with file upload options."""
    st.title("Institutions Details & Documents")

    responses = {}
    
    responses["institution_type"] = st.selectbox("Select Institution Type", ["Schools", "College", "Social Institutions", "Others"], key="institution_type")
    responses["institution_name"] = st.text_input("Enter Institution Name", key="institution_name")

    responses["gov_approvals"] = input_section("Government Approvals", "gov_approvals", add_comment=False, add_upload=True)
    responses["sanction_province"] = input_section("Sanction order by Province", "sanction_province", add_comment=False, add_upload=True)
    responses["electricity_sanction"] = input_section("Electricity Sanction Order", "electricity_sanction", add_comment=False, add_upload=True)
    responses["property_deeds"] = input_section("Copy of the Property Deeds", "property_deeds", add_comment=False, add_upload=True)
    responses["land_sketch"] = input_section("Land Sketch with Survey Measurement", "land_sketch", add_comment=False, add_upload=True)
    responses["land_tax"] = input_section("Land Tax Receipts (Latest)", "land_tax", add_comment=False, add_upload=True)
    responses["building_tax"] = input_section("Building Tax Receipts (Latest)", "building_tax", add_comment=False, add_upload=True)
    responses["land_na_inst"] = input_section("Land N A Documents", "land_na_inst", add_comment=False, add_upload=True)
    responses["mutation_inst"] = input_section("Mutation Certificate (Pokkuvaravu)", "mutation_inst", add_comment=False, add_upload=True)
    
    return responses

def main():
    # Check if survey has started
    if "survey_started" not in st.session_state:
        start_page()
        return

    # Initialize survey data if not exists
    if 'survey_data' not in st.session_state:
        st.session_state.survey_data = {
            'Property Ownership and Legal Documents': None,
            'Trust/Society Details & Documents': None,
            'Institutions Details & Documents': None
        }

    # Survey sections
    categories = {
        "Property Ownership and Legal Documents": property_survey,
        "Trust/Society Details & Documents": trust_society_survey,
        "Institutions Details & Documents": institution_survey
    }

    st.title("Survey Portal")
    st.write("Select a survey category to proceed:")

    # Allow user to choose section
    choice = st.radio("Choose a survey category:", list(categories.keys()))

    # Perform selected survey
    survey_data = categories[choice]()

    # Save the specific section data
    st.session_state.survey_data[choice] = survey_data

    # Prepare for final submission
    if st.button("Submit Entire Survey"):
        # Combine all sections
        final_survey_data = {
            'ashram_name': st.session_state.survey_data['Property Ownership and Legal Documents'].get('ashram_name', 'N/A'),
            'Property Ownership and Legal Documents': st.session_state.survey_data['Property Ownership and Legal Documents'],
            'Trust/Society Details & Documents': st.session_state.survey_data['Trust/Society Details & Documents'],
            'Institutions Details & Documents': st.session_state.survey_data['Institutions Details & Documents']
        }

        # Save the entire survey data
        save_survey_data(final_survey_data)

if __name__ == "__main__":
    main()