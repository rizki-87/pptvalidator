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
import csv

# Initialize T5 model
MODEL_NAME = "t5-small"
tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
model = AutoModelForSeq2SeqLM.from_pretrained(MODEL_NAME)

# Function to correct grammar using T5 model
def correct_grammar(text):
    if not text.strip():
        return text  # Return original text if empty
    try:
        input_text = f"grammar: {text}"
        input_ids = tokenizer.encode(input_text, return_tensors="pt", truncation=True)
        outputs = model.generate(input_ids, max_length=512)
        corrected_text = tokenizer.decode(outputs[0], skip_special_tokens=True)
        return corrected_text
    except Exception as e:
        st.error(f"Error in grammar correction: {e}")
        return text

# Function to validate fonts in a presentation
def validate_fonts(input_ppt, default_font):
    presentation = Presentation(input_ppt)
    font_issues = []

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.font.name != default_font:
                            font_issues.append({
                                'slide': slide_index,
                                'issue': 'Font Issue',
                                'text': run.text,
                                'corrected': ""
                            })

    return font_issues

# Function to save issues to CSV
def save_to_csv(issues, output_csv):
    # Filter out slides that have no issues
    issues_with_content = [issue for issue in issues if issue['text'].strip()]
    with open(output_csv, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=['slide', 'issue', 'text', 'corrected'])
        writer.writeheader()
        writer.writerows(issues_with_content)

# Main Streamlit app
def main():
    st.title("PPT Validator")

    # Create session state for uploaded file
    if "uploaded_file" not in st.session_state:
        st.session_state.uploaded_file = None

    uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])

    # Update session state with uploaded file
    if uploaded_file:
        st.session_state.uploaded_file = uploaded_file

    font_options = ["Arial", "Calibri", "Times New Roman", "Verdana", "Helvetica"]
    default_font = st.selectbox("Select the default font for validation", font_options)

    if st.session_state.uploaded_file and st.button("Run Validation"):
        with tempfile.TemporaryDirectory() as tmpdir:
            # Save uploaded file temporarily
            temp_ppt_path = Path(tmpdir) / "uploaded_ppt.pptx"
            with open(temp_ppt_path, "wb") as f:
                f.write(st.session_state.uploaded_file.getbuffer())

            # Output paths
            csv_output_path = Path(tmpdir) / "validation_report.csv"

            # Validate fonts
            font_issues = validate_fonts(temp_ppt_path, default_font)

            # Grammar validation
            grammar_issues = []
            for issue in font_issues:
                corrected_text = correct_grammar(issue['text'])
                if corrected_text != issue['text']:
                    grammar_issues.append({
                        'slide': issue['slide'],
                        'issue': issue['issue'],
                        'text': issue['text'],
                        'corrected': corrected_text
                    })

            # Save to CSV
            save_to_csv(grammar_issues, csv_output_path)

            # Display download link for CSV
            st.success("Validation completed!")
            st.download_button("Download Validation Report (CSV)", csv_output_path.read_bytes(),
                               file_name="validation_report.csv")

    # Reset button functionality
    if st.button("Reset"):
        # Clear uploaded file and reset session state
        st.session_state.uploaded_file = None
        st.session_state.clear()

        # Refresh the app to the initial state
        st.experimental_rerun()

if __name__ == "__main__":
    main()
