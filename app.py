# import streamlit as st
# from pptx import Presentation
# from pptx.dml.color import RGBColor
# from pathlib import Path
# from datetime import datetime
# import os
# import csv
# import tempfile
# import re
# from transformers import AutoModelForSeq2SeqLM, AutoTokenizer

# # Initialize T5 model for grammar correction
# MODEL_NAME = "vennify/t5-base-grammar-correction"
# CACHE_DIR = "./model_cache"
# tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME, cache_dir=CACHE_DIR)
# model = AutoModelForSeq2SeqLM.from_pretrained(MODEL_NAME, cache_dir=CACHE_DIR)

# # Function to correct grammar and spelling
# def correct_grammar(text):
#     input_text = f"grammar correction: {text}"
#     input_ids = tokenizer.encode(input_text, return_tensors="pt")
#     outputs = model.generate(input_ids, max_length=512)
#     corrected_text = tokenizer.decode(outputs[0], skip_special_tokens=True)
#     return corrected_text

# # Function to validate fonts
# def validate_fonts(ppt_path, selected_font):
#     presentation = Presentation(ppt_path)
#     font_issues = []

#     for slide_index, slide in enumerate(presentation.slides, start=1):
#         for shape in slide.shapes:
#             if shape.has_text_frame:
#                 for paragraph in shape.text_frame.paragraphs:
#                     for run in paragraph.runs:
#                         current_font = run.font.name
#                         text = run.text.strip()
#                         if current_font and current_font != selected_font and text:
#                             font_issues.append({
#                                 "slide": slide_index,
#                                 "text": text,
#                                 "font": current_font,
#                             })
#                             # Highlight text with inconsistent font
#                             run.font.color.rgb = RGBColor(255, 0, 0)
#     # Save highlighted presentation to a temporary file
#     temp_highlighted_ppt = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
#     presentation.save(temp_highlighted_ppt.name)
#     return font_issues, temp_highlighted_ppt.name

# # Function to validate grammar and spelling
# def validate_grammar_and_spelling(ppt_path):
#     presentation = Presentation(ppt_path)
#     grammar_issues = []

#     for slide_index, slide in enumerate(presentation.slides, start=1):
#         for shape in slide.shapes:
#             if shape.has_text_frame:
#                 text = " ".join(run.text for paragraph in shape.text_frame.paragraphs for run in paragraph.runs).strip()
#                 if text:
#                     corrected_text = correct_grammar(text)
#                     if text.lower().strip() != corrected_text.lower().strip():
#                         grammar_issues.append({
#                             "slide": slide_index,
#                             "original": text,
#                             "corrected": corrected_text,
#                         })
#     return grammar_issues

# # Function to validate punctuation
# def validate_punctuation(ppt_path):
#     presentation = Presentation(ppt_path)
#     punctuation_issues = []

#     for slide_index, slide in enumerate(presentation.slides, start=1):
#         for shape in slide.shapes:
#             if shape.has_text_frame:
#                 for paragraph in shape.text_frame.paragraphs:
#                     for run in paragraph.runs:
#                         text = run.text.strip()
#                         if not text:
#                             continue

#                         # Detect excessive punctuation
#                         excessive_punctuation = re.findall(r"[!?.:,;]{2,}", text)
#                         for punct in excessive_punctuation:
#                             punctuation_issues.append({
#                                 "slide": slide_index,
#                                 "text": text,
#                                 "issue": f"Excessive punctuation: {punct}",
#                             })

#                         # Detect repeated words
#                         repeated_words = re.findall(r"\b(\w+)\s+\1\b", text, flags=re.IGNORECASE)
#                         for word in repeated_words:
#                             punctuation_issues.append({
#                                 "slide": slide_index,
#                                 "text": text,
#                                 "issue": f"Repeated word: {word}",
#                             })
#     return punctuation_issues

# # Function to save issues to CSV
# def save_issues_to_csv(font_issues, grammar_issues, punctuation_issues):
#     # Save issues to a temporary CSV file
#     temp_csv = tempfile.NamedTemporaryFile(delete=False, suffix=".csv", mode="w", encoding="utf-8-sig", newline="")
#     with open(temp_csv.name, mode="w", newline="", encoding="utf-8-sig") as file:
#         writer = csv.writer(file)
#         writer.writerow(["Slide Number", "Issue Type", "Original Text", "Details"])

#         for issue in font_issues:
#             writer.writerow([issue["slide"], "Font Issue", issue["text"], f"Font: {issue['font']}"])

#         for issue in grammar_issues:
#             writer.writerow([issue["slide"], "Grammar Issue", issue["original"], issue["corrected"]])

#         for issue in punctuation_issues:
#             writer.writerow([issue["slide"], "Punctuation Issue", issue["text"], issue["issue"]])
#     return temp_csv.name

# # Streamlit app starts here
# st.title("PPT Validator")
# st.write("This is a Streamlit app for validating PowerPoint presentations.")

# # File upload widget
# uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])

# # Font selection widget with static font list
# default_fonts = ["Arial", "Calibri", "Times New Roman", "Verdana", "Tahoma"]
# default_font = st.selectbox("Select the default font for validation", default_fonts)

# # Process the file on button click
# if uploaded_file and st.button("Run Validation"):
#     with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_ppt:
#         temp_ppt.write(uploaded_file.read())
#         temp_ppt_path = temp_ppt.name

#     # Run validations
#     font_issues, highlighted_ppt_path = validate_fonts(temp_ppt_path, default_font)
#     grammar_issues = validate_grammar_and_spelling(temp_ppt_path)
#     punctuation_issues = validate_punctuation(temp_ppt_path)

#     # Save issues to CSV
#     csv_file_path = save_issues_to_csv(font_issues, grammar_issues, punctuation_issues)

#     # Display results
#     st.write("### Validation Results:")
#     st.write(f"Font Issues: {len(font_issues)}")
#     st.write(f"Grammar Issues: {len(grammar_issues)}")
#     st.write(f"Punctuation Issues: {len(punctuation_issues)}")

#     st.write("### Download Results:")
#     st.download_button(
#         label="Download Highlighted PPTX",
#         data=open(highlighted_ppt_path, "rb").read(),
#         file_name="highlighted_presentation.pptx",
#         mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
#     )

#     st.download_button(
#         label="Download CSV Report",
#         data=open(csv_file_path, "rb").read(),
#         file_name="Validation_Report.csv",
#         mime="text/csv",
#     )

#     # Clean up temporary file
#     os.unlink(temp_ppt_path)
#     os.unlink(highlighted_ppt_path)
#     os.unlink(csv_file_path)

import streamlit as st
import tempfile
from pathlib import Path
from pptx import Presentation
from transformers import AutoTokenizer, AutoModelForSeq2SeqLM
from spellchecker import SpellChecker
import csv
import re
import string

# Initialize T5 model
MODEL_NAME = "t5-small"
tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
model = AutoModelForSeq2SeqLM.from_pretrained(MODEL_NAME)

# Function to detect and correct misspellings
def detect_misspellings(text):
    spell = SpellChecker()
    words = text.split()
    misspellings = {}

    for word in words:
        # Remove punctuation from the word for checking
        clean_word = word.strip(string.punctuation)
        
        # Check if the word is misspelled
        if clean_word and clean_word.lower() not in spell:
            correction = spell.correction(clean_word)
            if correction:  # Only suggest if there's a valid correction
                misspellings[word] = correction

    return misspellings


# Integrate into the main logic
def validate_spelling(input_ppt):
    presentation = Presentation(input_ppt)
    spelling_issues = []

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.text.strip():
                            misspellings = detect_misspellings(run.text)
                            for original_word, correction in misspellings.items():
                                spelling_issues.append({
                                    'slide': slide_index,
                                    'issue': 'Misspelling',
                                    'text': f"Original: {original_word}",
                                    'corrected': f"Suggestion: {correction}"
                                })

    return spelling_issues

# Function to validate fonts in a presentation
def validate_fonts(input_ppt, default_font):
    presentation = Presentation(input_ppt)
    issues = []

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.text.strip():  # Skip empty text
                            # Check for inconsistent fonts
                            if run.font.name != default_font:
                                issues.append({
                                    'slide': slide_index,
                                    'issue': 'Inconsistent Font',
                                    'text': run.text,
                                    'corrected': f"Expected font: {default_font}"
                                })
    return issues

# Function to detect punctuation issues
def validate_punctuation(input_ppt):
    presentation = Presentation(input_ppt)
    punctuation_issues = []

    # Define patterns for punctuation problems
    excessive_punctuation_pattern = r"([!?.:,;]{2,})"  # Two or more punctuation marks
    repeated_word_pattern = r"\b(\w+)\s+\1\b"  # Repeated words (e.g., "the the")

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.text.strip():
                            text = run.text

                            # Check excessive punctuation
                            match = re.search(excessive_punctuation_pattern, text)
                            if match:
                                punctuation_marks = match.group(1)  # Extract punctuation
                                punctuation_issues.append({
                                    'slide': slide_index,
                                    'issue': 'Punctuation Marks',
                                    'text': text,
                                    'corrected': f"Excessive punctuation marks detected ({punctuation_marks})"
                                })

                            # Check repeated words
                            if re.search(repeated_word_pattern, text, flags=re.IGNORECASE):
                                punctuation_issues.append({
                                    'slide': slide_index,
                                    'issue': 'Punctuation Marks',
                                    'text': text,
                                    'corrected': "Repeated words detected"
                                })

    return punctuation_issues

# Function to save issues to CSV
def save_to_csv(issues, output_csv):
    with open(output_csv, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=['slide', 'issue', 'text', 'corrected'])
        writer.writeheader()
        writer.writerows(issues)

# Main Streamlit app
def main():
    st.title("PPT Validator")

    uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])

    font_options = ["Arial", "Calibri", "Times New Roman", "Verdana", "Helvetica"]
    default_font = st.selectbox("Select the default font for validation", font_options)

    if uploaded_file and st.button("Run Validation"):
        with tempfile.TemporaryDirectory() as tmpdir:
            # Save uploaded file temporarily
            temp_ppt_path = Path(tmpdir) / "uploaded_ppt.pptx"
            with open(temp_ppt_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # Output path
            csv_output_path = Path(tmpdir) / "validation_report.csv"

            # Run validations
            font_issues = validate_fonts(temp_ppt_path, default_font)
            punctuation_issues = validate_punctuation(temp_ppt_path)
            spelling_issues = validate_spelling(temp_ppt_path)

            # Combine issues and save to CSV
            combined_issues = font_issues + punctuation_issues + spelling_issues
            save_to_csv(combined_issues, csv_output_path)

            # Display success and download link
            st.success("Validation completed!")
            st.download_button("Download Validation Report (CSV)", csv_output_path.read_bytes(),
                               file_name="validation_report.csv")

if __name__ == "__main__":
    main()

