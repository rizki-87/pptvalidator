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
from pptx import Presentation
from pptx.dml.color import RGBColor
from pathlib import Path
from transformers import AutoModelForSeq2SeqLM, AutoTokenizer
from concurrent.futures import ThreadPoolExecutor
import re
import tempfile
import csv
import time
import os

# Initialize the T5 model
MODEL_NAME = "t5-small"
CACHE_DIR = "./model_cache"

# Load model and tokenizer
@st.cache_resource
def load_model():
    tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME, cache_dir=CACHE_DIR)
    model = AutoModelForSeq2SeqLM.from_pretrained(MODEL_NAME, cache_dir=CACHE_DIR)
    return tokenizer, model

tokenizer, model = load_model()

# Define font list
DEFAULT_FONTS = [
    "Arial", "Calibri", "Times New Roman", "Verdana", "Tahoma",
    "Georgia", "Comic Sans MS", "Impact", "Courier New", "Lucida Console",
    "DejaVu Sans", "DejaVu Serif", "STIXGeneral"
]

# Function for grammar correction
def correct_grammar(text):
    input_text = f"grammar: {text}"
    input_ids = tokenizer.encode(input_text, return_tensors="pt")
    outputs = model.generate(input_ids, max_length=512)
    return tokenizer.decode(outputs[0], skip_special_tokens=True)

# Function to validate fonts
def validate_fonts(file_path, default_font, output_ppt_path):
    presentation = Presentation(file_path)
    font_issues = []
    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        font_name = getattr(run.font, "name", None)
                        if font_name and font_name != default_font:
                            font_issues.append(
                                {"slide": slide_index, "issue": "Font Issue", "text": run.text, "detail": font_name}
                            )
                            run.font.color.rgb = RGBColor(255, 0, 0)  # Highlight text

    presentation.save(output_ppt_path)
    return font_issues

# Function to check grammar and punctuation in parallel
def check_grammar_punctuation(slides_texts):
    def process_slide_text(slide_text):
        corrected_text = correct_grammar(slide_text)
        if corrected_text != slide_text:
            return {"issue": "Grammar Issue", "original": slide_text, "corrected": corrected_text}
        return None

    with ThreadPoolExecutor() as executor:
        results = list(executor.map(process_slide_text, slides_texts))
    return [result for result in results if result]

# Function to save CSV
def save_csv(data, output_csv_path):
    with open(output_csv_path, mode="w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=["slide", "issue", "text", "detail"])
        writer.writeheader()
        for row in data:
            writer.writerow(row)

# Streamlit UI
def main():
    st.title("PPT Validator")
    st.write("This is a Streamlit app for validating PowerPoint presentations.")

    if "validation_complete" not in st.session_state:
        st.session_state.validation_complete = False
        st.session_state.output_csv_path = None
        st.session_state.highlighted_ppt_path = None

    if "reset" not in st.session_state:
        st.session_state.reset = False

    if st.session_state.reset:
        st.session_state.validation_complete = False
        st.session_state.output_csv_path = None
        st.session_state.highlighted_ppt_path = None
        st.session_state.reset = False
        st.experimental_rerun()

    uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])
    default_font = st.selectbox("Select the default font for validation", DEFAULT_FONTS)

    if uploaded_file and default_font:
        if st.button("Run Validation"):
            with st.spinner("Processing the presentation..."):
                start_time = time.time()

                # Save uploaded file to a temporary location
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_file:
                    temp_file.write(uploaded_file.read())
                    temp_ppt_path = temp_file.name

                # Paths for output files
                highlighted_ppt_path = Path(tempfile.gettempdir()) / "highlighted_presentation.pptx"
                output_csv_path = Path(tempfile.gettempdir()) / "validation_report.csv"

                # Validate fonts
                font_issues = validate_fonts(temp_ppt_path, default_font, highlighted_ppt_path)

                # Extract texts for grammar and punctuation checking
                presentation = Presentation(temp_ppt_path)
                slides_texts = [
                    " ".join(run.text for shape in slide.shapes if shape.has_text_frame
                             for paragraph in shape.text_frame.paragraphs
                             for run in paragraph.runs)
                    for slide in presentation.slides
                ]

                # Run grammar and punctuation check
                grammar_punctuation_issues = check_grammar_punctuation(slides_texts)

                # Combine all issues
                all_issues = font_issues + [
                    {"slide": i + 1, "issue": gp["issue"], "text": gp["original"], "detail": gp["corrected"]}
                    for i, gp in enumerate(grammar_punctuation_issues)
                ]

                # Save to CSV
                save_csv(all_issues, output_csv_path)

                # Update session state
                st.session_state.validation_complete = True
                st.session_state.output_csv_path = output_csv_path
                st.session_state.highlighted_ppt_path = highlighted_ppt_path

                end_time = time.time()
                st.success(f"Validation completed in {round(end_time - start_time, 2)} seconds!")

    if st.session_state.validation_complete:
        st.download_button(
            label="Download validation report (CSV)",
            data=open(st.session_state.output_csv_path, "rb").read(),
            file_name="validation_report.csv",
            mime="text/csv",
        )
        st.download_button(
            label="Download highlighted PowerPoint",
            data=open(st.session_state.highlighted_ppt_path, "rb").read(),
            file_name="highlighted_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        if st.button("Reset"):
            st.session_state.reset = True

if __name__ == "__main__":
    main()
