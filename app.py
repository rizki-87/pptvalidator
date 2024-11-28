import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor
from pathlib import Path
from datetime import datetime
import os
import csv
import tempfile
import re
from transformers import AutoModelForSeq2SeqLM, AutoTokenizer
from matplotlib import font_manager

# Initialize T5 model for grammar correction
MODEL_NAME = "vennify/t5-base-grammar-correction"
CACHE_DIR = "./model_cache"
tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME, cache_dir=CACHE_DIR)
model = AutoModelForSeq2SeqLM.from_pretrained(MODEL_NAME, cache_dir=CACHE_DIR)

# Function to correct grammar and spelling
def correct_grammar(text):
    input_text = f"grammar correction: {text}"
    input_ids = tokenizer.encode(input_text, return_tensors="pt")
    outputs = model.generate(input_ids, max_length=512)
    corrected_text = tokenizer.decode(outputs[0], skip_special_tokens=True)
    return corrected_text

# Function to validate fonts
def validate_fonts(ppt_path, selected_font, output_ppt_path):
    presentation = Presentation(ppt_path)
    font_issues = []

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        current_font = run.font.name
                        text = run.text.strip()
                        if current_font and current_font != selected_font and text:
                            font_issues.append({
                                "slide": slide_index,
                                "text": text,
                                "font": current_font,
                            })
                            # Highlight text with inconsistent font
                            run.font.color.rgb = RGBColor(255, 0, 0)

    presentation.save(output_ppt_path)
    return font_issues

# Function to validate grammar and spelling
def validate_grammar_and_spelling(ppt_path):
    presentation = Presentation(ppt_path)
    grammar_issues = []

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = " ".join(run.text for paragraph in shape.text_frame.paragraphs for run in paragraph.runs).strip()
                if text:
                    corrected_text = correct_grammar(text)
                    if text.lower().strip() != corrected_text.lower().strip():
                        grammar_issues.append({
                            "slide": slide_index,
                            "original": text,
                            "corrected": corrected_text,
                        })
    return grammar_issues

# Function to validate punctuation
def validate_punctuation(ppt_path):
    presentation = Presentation(ppt_path)
    punctuation_issues = []

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        text = run.text.strip()
                        if not text:
                            continue

                        # Detect excessive punctuation
                        excessive_punctuation = re.findall(r"[!?.:,;]{2,}", text)
                        for punct in excessive_punctuation:
                            punctuation_issues.append({
                                "slide": slide_index,
                                "text": text,
                                "issue": f"Excessive punctuation: {punct}",
                            })

                        # Detect repeated words
                        repeated_words = re.findall(r"\b(\w+)\s+\1\b", text, flags=re.IGNORECASE)
                        for word in repeated_words:
                            punctuation_issues.append({
                                "slide": slide_index,
                                "text": text,
                                "issue": f"Repeated word: {word}",
                            })
    return punctuation_issues

# Function to save issues to CSV
def save_issues_to_csv(font_issues, grammar_issues, punctuation_issues, output_name):
    downloads_path = Path.home() / "Downloads"
    csv_file_name = downloads_path / f"{output_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

    with open(csv_file_name, mode='w', newline='', encoding='utf-8-sig') as file:
        writer = csv.writer(file)
        writer.writerow(["Slide Number", "Issue Type", "Original Text", "Details"])

        for issue in font_issues:
            writer.writerow([issue["slide"], "Font Issue", issue["text"], f"Font: {issue['font']}"])

        for issue in grammar_issues:
            writer.writerow([issue["slide"], "Grammar Issue", issue["original"], issue["corrected"]])

        for issue in punctuation_issues:
            writer.writerow([issue["slide"], "Punctuation Issue", issue["text"], issue["issue"]])

    return csv_file_name

# Streamlit app starts here
st.title("PPT Validator")
st.write("This is a Streamlit app for validating PowerPoint presentations.")

# File upload widget
uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])

# Font selection widget
available_fonts = sorted(set(f.name for f in font_manager.fontManager.ttflist))
default_font = st.selectbox("Select the default font for validation", available_fonts)

# Process the file on button click
if uploaded_file and st.button("Run Validation"):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_ppt:
        temp_ppt.write(uploaded_file.read())
        temp_ppt_path = temp_ppt.name

    # Paths for output files
    downloads_path = Path.home() / "Downloads"
    highlighted_ppt_path = downloads_path / "highlighted_presentation.pptx"

    # Run validations
    font_issues = validate_fonts(temp_ppt_path, default_font, highlighted_ppt_path)
    grammar_issues = validate_grammar_and_spelling(temp_ppt_path)
    punctuation_issues = validate_punctuation(temp_ppt_path)

    # Save issues to CSV
    csv_file_path = save_issues_to_csv(font_issues, grammar_issues, punctuation_issues, "Validation_Report")

    # Display results
    st.write("### Validation Results:")
    st.write(f"Font Issues: {len(font_issues)}")
    st.write(f"Grammar Issues: {len(grammar_issues)}")
    st.write(f"Punctuation Issues: {len(punctuation_issues)}")

    st.write("### Download Results:")
    st.download_button(
        label="Download Highlighted PPTX",
        data=open(highlighted_ppt_path, "rb").read(),
        file_name="highlighted_presentation.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )

    st.download_button(
        label="Download CSV Report",
        data=open(csv_file_path, "rb").read(),
        file_name=csv_file_path.name,
        mime="text/csv",
    )

    # Clean up temporary file
    os.unlink(temp_ppt_path)
