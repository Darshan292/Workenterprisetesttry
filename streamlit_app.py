import streamlit as st
import os
import pandas as pd
import multiprocessing
from Converting_Data_Create_EIB_Final import (
    creating_eib_files_with_parallel_processing,
    load_file_params,
    validate_required_fields_in_eib,
)

# ==============================
# CONFIGURATION
# ==============================

BASE_DIR = r"C:\Users\patel\Desktop\final code\\"
TEMPLATE_DIR = os.path.join(BASE_DIR, "EIB")
MAPPING_FILE_PATH = os.path.join(BASE_DIR, "EIB", "Combined_Mapping.xlsx")
LOAD_TEMPLATE_MAPPING = {
    "Location": "Put_Location_v46.0.xlsx",
    "Job Profile": "Submit_Job_Profile_v46.0.xlsx",
    "Cost Center": "Put_Cost_Center_v46.0.xlsx",
}

st.set_page_config(
    layout="wide",
    page_title="EIB Transformation Tool",
    page_icon="🏢"
)

# ==============================
# HEADER
# ==============================

st.markdown("""
# 🏢 Enterprise EIB Transformation Tool
Internal Data Processing Portal
""")

st.markdown("---")

# ==============================
# HELPER FUNCTIONS
# ==============================

def list_excel_files(folder):
    return [f for f in os.listdir(folder) if f.endswith(".xlsx")]

def save_uploaded_file(uploaded_file, save_path):
    full_path = os.path.join(save_path, uploaded_file.name)
    with open(full_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return full_path

def preview_excel(file_path, key_suffix):
    try:
        xls = pd.ExcelFile(file_path)
        sheet = st.selectbox(
            f"Select Sheet",
            xls.sheet_names,
            key=f"{file_path}_{key_suffix}"
        )
        df = pd.read_excel(file_path, sheet_name=sheet)

        st.caption(f"Rows: {df.shape[0]} | Columns: {df.shape[1]}")
        st.dataframe(df.head(10), use_container_width=True)

    except Exception as e:
        st.warning(f"Preview not available: {e}")

# ==============================
# SIDEBAR (Corporate Control Panel)
# ==============================

with st.sidebar:

    st.header("⚙ Configuration")

    load_type = st.selectbox(
        "Select Load Type",
        ["Location", "Job Profile", "Cost Center"]
    )
    load_list = [load_type]

    st.markdown("---")

    input_mode = st.radio(
        "Input File Source",
        ["Select from Folder", "Upload File"]
    )

    st.markdown("---")

    template_files = [
        f for f in list_excel_files(TEMPLATE_DIR)
        if f.lower() != "combined_mapping.xlsx"
    ]

    # selected_template = load_file_params(load_type, BASE_DIR)
    template_file_name = LOAD_TEMPLATE_MAPPING.get(load_type)
    template_preview_path = os.path.join(TEMPLATE_DIR, template_file_name)
    

# ==============================
# INPUT FILE HANDLING
# ==============================

input_file_path = None

if input_mode == "Select from Folder":
    input_files = list_excel_files(BASE_DIR)
    selected_input = st.selectbox(
        "Select Input Excel File",
        input_files
    )
    input_file_path = os.path.join(BASE_DIR, selected_input)

else:
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    if uploaded_file:
        input_file_path = save_uploaded_file(uploaded_file, BASE_DIR)
        st.success("File uploaded successfully.")

# ==============================
# PREVIEW SECTION
# ==============================

st.markdown("## 📄 Data Preview")
st.markdown("---")

left_col, right_col = st.columns(2)

with left_col:
    st.subheader("Input File")

    if input_file_path:
        preview_excel(input_file_path, "input")
    else:
        st.info("Select or upload an input file to preview.")

with right_col:
    if "generated_file" not in st.session_state:
        st.subheader("Template Preview")
        preview_excel(template_preview_path, "template")
    else:
        st.subheader("Generated Output")
        preview_excel(st.session_state["generated_file"], "output")

# ==============================
# TRANSFORM SECTION
# ==============================

st.markdown("---")

transform_disabled = input_file_path is None

if st.button(
    "🚀 Execute Transformation",
    disabled=transform_disabled,
    use_container_width=True
):

    try:
        multiprocessing.freeze_support()

        _, expected_output_path = load_file_params(load_type, BASE_DIR)

        previous_mtime = None
        if os.path.exists(expected_output_path):
            previous_mtime = os.path.getmtime(expected_output_path)

        with st.spinner("Processing... Please wait..."):
            try:
                creating_eib_files_with_parallel_processing(
                    load_list,
                    BASE_DIR,
                    MAPPING_FILE_PATH,
                    input_file_path
                )
            except Exception as e:
                st.warning("Backend returned an error, validating output...")
                st.write("Error:", e)
                validation_errors = None

        if os.path.exists(expected_output_path):
            # new_mtime = os.path.getmtime(expected_output_path)
            validation_errors = validate_required_fields_in_eib(expected_output_path)
            # to change to see output
            if not validation_errors: 
                st.error("Required field validation failed.")

                for column, count in validation_errors.items():
                    st.error(f"{column}: {count} rows missing")

            else:
                st.session_state["generated_file"] = expected_output_path
            
                # if validation_errors and len(validation_errors) > 0:
                #     st.error("❌ Required Field Validation Failed")
            
                #     error_df = pd.DataFrame(validation_errors)
            
                #     st.markdown("### Missing Required Fields")
                #     st.dataframe(error_df, use_container_width=True)
            
                # else:
                st.success("Transformation completed successfully. All required fields are populated.")           
        else:
            st.error("Output file not found.")

    except Exception as e:
        st.error(f"Critical error: {e}")

# ==============================
# DOWNLOAD SECTION
# ==============================

if "generated_file" in st.session_state:
    st.markdown("---")

    with open(st.session_state["generated_file"], "rb") as f:
        st.download_button(
            label="⬇ Download Generated EIB File",
            data=f,
            file_name=os.path.basename(st.session_state["generated_file"]),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )