import streamlit as st
from pptx import Presentation
from pathlib import Path
from datetime import datetime
from main import (
    check_inconsistent_fonts,
    check_grammar_and_spelling,
    check_punctuation_issues,
    save_issues_to_csv,
)

# Streamlit Title
st.title("PPT Validator")
st.write("This is a Streamlit app for validating PowerPoint presentations.")

# Upload PPTX file
uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])

# Dropdown for default font selection
default_font = st.selectbox("Select the default font for validation", ["Arial", "Calibri", "Times New Roman"])

# Submit button
if st.button("Run Validation"):
    if uploaded_file is not None:
        # Save the uploaded file locally
        with open("uploaded_presentation.pptx", "wb") as f:
            f.write(uploaded_file.read())
        st.success("File uploaded successfully!")

        # Validate PowerPoint file
        try:
            st.info("Starting validation...")
            file_path = "uploaded_presentation.pptx"
            output_pptx_path = f"highlighted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"

            # Step 1: Check for inconsistent fonts
            font_issues = check_inconsistent_fonts(file_path, default_font, output_pptx_path)

            # Step 2: Check for grammar and spelling issues
            grammar_issues = check_grammar_and_spelling(file_path)

            # Step 3: Check for punctuation issues
            punctuation_issues = check_punctuation_issues(file_path)

            # Step 4: Save results to CSV
            csv_file = save_issues_to_csv(font_issues, grammar_issues, punctuation_issues, "Validation_Report")

            # Show results
            st.success("Validation completed!")
            st.write(f"Font issues: {len(font_issues)}")
            st.write(f"Grammar and spelling issues: {len(grammar_issues)}")
            st.write(f"Punctuation issues: {len(punctuation_issues)}")

            # Provide download links
            with open(csv_file, "rb") as f:
                st.download_button("Download Validation Report (CSV)", f, file_name="Validation_Report.csv")
            with open(output_pptx_path, "rb") as f:
                st.download_button("Download Highlighted PowerPoint", f, file_name="Highlighted_Presentation.pptx")
        except Exception as e:
            st.error(f"An error occurred during validation: {e}")
    else:
        st.error("Please upload a PowerPoint file.")
