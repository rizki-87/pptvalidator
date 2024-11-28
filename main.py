import os
import logging
import csv
import re
from datetime import datetime
from pathlib import Path
from pptx import Presentation
from pptx.dml.color import RGBColor
from transformers import AutoModelForSeq2SeqLM, AutoTokenizer

# Disable symlink warnings in huggingface_hub
import warnings
os.environ["HF_HUB_DISABLE_SYMLINKS_WARNING"] = "1"
warnings.filterwarnings("ignore", category=UserWarning, module="huggingface_hub")

# Configure logging
logging.basicConfig(
    filename="ppt_validator.log",
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

# Initialize the T5 model for grammar correction
MODEL_NAME = "vennify/t5-base-grammar-correction"
CACHE_DIR = "./model_cache"

# Ensure cache directory exists
Path(CACHE_DIR).mkdir(parents=True, exist_ok=True)

# Load or download the model
try:
    tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME, cache_dir=CACHE_DIR)
    model = AutoModelForSeq2SeqLM.from_pretrained(MODEL_NAME, cache_dir=CACHE_DIR)
except Exception as e:
    logging.error(f"Error loading model: {e}")
    raise

def correct_grammar(text):
    """Correct grammar and spelling using the T5 model."""
    try:
        logging.info(f"Original Text: {text}")

        # Split text into chunks if too long for the model
        chunk_size = 200
        chunks = [text[i:i + chunk_size] for i in range(0, len(text), chunk_size)]
        corrected_text = []

        for chunk in chunks:
            input_text = f"grammar correction: {chunk}"
            input_ids = tokenizer.encode(input_text, return_tensors="pt")
            outputs = model.generate(input_ids, max_length=512)
            corrected_chunk = tokenizer.decode(outputs[0], skip_special_tokens=True)
            corrected_text.append(corrected_chunk)

        final_text = " ".join(corrected_text)
        logging.info(f"Corrected Text: {final_text}")
        return final_text

    except Exception as e:
        logging.error(f"Error during grammar correction: {e}")
        return text

def check_inconsistent_fonts(file_path, default_font, output_pptx_path=None):
    """Check for inconsistent fonts in the presentation."""
    presentation = Presentation(file_path)
    inconsistent_fonts = []

    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        font = getattr(run.font, 'name', None)
                        text = run.text.strip()
                        if font and font != default_font and text:
                            inconsistent_fonts.append({
                                'slide_number': presentation.slides.index(slide) + 1,
                                'text': text,
                                'font': font
                            })
                            if run.font and hasattr(run.font, 'color'):
                                run.font.color.rgb = RGBColor(255, 0, 0)

    if output_pptx_path:
        presentation.save(output_pptx_path)

    return inconsistent_fonts

def check_grammar_and_spelling(file_path):
    """Detect grammar and spelling issues in the presentation."""
    presentation = Presentation(file_path)
    issues = []

    for slide in presentation.slides:
        slide_text = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                slide_text.extend(
                    run.text.strip() for paragraph in shape.text_frame.paragraphs for run in paragraph.runs
                )

        if slide_text:
            combined_text = " ".join(slide_text)
            corrected_text = correct_grammar(combined_text)
            if combined_text.lower().strip() != corrected_text.lower().strip():
                issues.append({
                    'slide_number': presentation.slides.index(slide) + 1,
                    'issue': "Grammar & Spelling Issue",
                    'original': combined_text,
                    'corrected': corrected_text
                })

    return issues

def check_punctuation_issues(file_path):
    """Detect excessive punctuation and repeated words in the presentation."""
    presentation = Presentation(file_path)
    punctuation_issues = []

    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        text = run.text.strip()
                        if text:
                            # Regex for punctuation and repeated words
                            excessive_punctuation_pattern = r"(?:[!?.:,;]{2,})"
                            repeated_word_pattern = r"\b(\w+)\s+\1\b"
                            if re.findall(excessive_punctuation_pattern, text) or re.findall(repeated_word_pattern, text):
                                punctuation_issues.append({
                                    'slide_number': presentation.slides.index(slide) + 1,
                                    'text': text,
                                    'issue': "Punctuation Issue"
                                })

    return punctuation_issues

def save_issues_to_csv(font_issues, grammar_and_spelling_issues, punctuation_issues, output_name):
    """Save detected issues to a CSV file."""
    csv_file_name = f"{output_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

    try:
        with open(csv_file_name, mode='w', newline='', encoding='utf-8-sig') as file:
            writer = csv.writer(file)
            writer.writerow(["Slide Number", "Issue Type", "Original Text", "Corrected Text or Details"])
            for issue in font_issues:
                writer.writerow([issue['slide_number'], "Inconsistent Font", issue['text'], f"Font: {issue['font']}"])
            for issue in grammar_and_spelling_issues:
                writer.writerow([issue['slide_number'], issue['issue'], issue['original'], issue['corrected']])
            for issue in punctuation_issues:
                writer.writerow([issue['slide_number'], "Punctuation Issue", issue['text'], issue['issue']])
        return csv_file_name
    except Exception as e:
        logging.error(f"Error saving CSV: {e}")
        return None
