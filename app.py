import streamlit as st
import tempfile
from generate_template import generate_output  # your function to generate output
import os

st.set_page_config(page_title="Event Template Generator")
st.title("Event Template Generator")

st.write("Upload Source and Staff Specialty files to generate the output template.")

source_file = st.file_uploader("Upload Source File", type=["xlsx"])
staff_file = st.file_uploader("Upload Staff Specialty File", type=["xlsx"])

if source_file and staff_file:
    # Save uploaded files to temporary files
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as src_tmp:
        src_path = src_tmp.name
        src_tmp.write(source_file.getbuffer())

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as staff_tmp:
        staff_path = staff_tmp.name
        staff_tmp.write(staff_file.getbuffer())

    # Output temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as out_tmp:
        output_path = out_tmp.name

    # Show loading spinner while processing
    with st.spinner("Generating output file... Please wait."):
        generate_output(src_path, staff_path, output_path)  # your function should save output to output_path

    st.success("Output generated successfully!")

    # Provide download button
    with open(output_path, "rb") as f:
        st.download_button(
            label="Download Output File",
            data=f,
            file_name="output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Optional: cleanup temp files
    os.remove(src_path)
    os.remove(staff_path)
