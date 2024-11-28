import streamlit as st
from matplotlib import font_manager
from pptx import Presentation
from pptx.dml.color import RGBColor
import os
import tempfile

# Title and description
st.title("PPT Validator")
st.write("This is a Streamlit app for validating PowerPoint presentations.")

# File upload widget
uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])

# Get the list of available system fonts
available_fonts = sorted(set(f.name for f in font_manager.fontManager.ttflist))

# Add more fonts manually if required
additional_fonts = ["Arial", "Calibri", "Times New Roman"]
all_fonts = sorted(set(available_fonts + additional_fonts))

# Dropdown for selecting the default font
default_font = st.selectbox("Select the default font for validation", all_fonts)

# Function to validate PowerPoint fonts
def validate_fonts(file_path, selected_font):
    presentation = Presentation(file_path)
    issues = []

    for slide_number, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        current_font = run.font.name
                        if current_font and current_font != selected_font:
                            issues.append({
                                "slide": slide_number,
                                "text": run.text.strip(),
                                "font": current_font,
                            })
    return issues

# Check if file is uploaded
if uploaded_file:
    # Save the uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_file:
        temp_file.write(uploaded_file.read())
        temp_file_path = temp_file.name

    # Add a "Run Validation" button
    if st.button("Run Validation"):
        st.write("Processing the presentation...")
        font_issues = validate_fonts(temp_file_path, default_font)

        if font_issues:
            st.write("Found font issues:")
            for issue in font_issues:
                st.write(
                    f"Slide {issue['slide']}: '{issue['text']}' (Font: {issue['font']}, Expected: {default_font})"
                )
        else:
            st.success("No font issues found!")

        # Clean up temporary file
        os.unlink(temp_file_path)
