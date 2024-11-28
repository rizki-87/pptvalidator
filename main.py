import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pptx import Presentation
from pptx.dml.color import RGBColor
from datetime import datetime
import os
import csv
import re
from pathlib import Path
from matplotlib import font_manager
from transformers import AutoModelForSeq2SeqLM, AutoTokenizer
import logging
import warnings

# Configure logging for debugging
logging.basicConfig(
    filename="ppt_validator.log",
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

# Disable symlink warnings in huggingface_hub
os.environ["HF_HUB_DISABLE_SYMLINKS_WARNING"] = "1"
warnings.filterwarnings("ignore", category=UserWarning, module="huggingface_hub")

# Define model name and cache directory
MODEL_NAME = "vennify/t5-base-grammar-correction"
CACHE_DIR = "./model_cache"

# Ensure cache directory exists
Path(CACHE_DIR).mkdir(parents=True, exist_ok=True)

# Load T5 model for grammar correction
try:
    tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME, cache_dir=CACHE_DIR)
    model = AutoModelForSeq2SeqLM.from_pretrained(MODEL_NAME, cache_dir=CACHE_DIR)
except Exception as e:
    logging.error(f"Error loading model: {e}")
    raise e  # Stop execution if model fails to load

# Grammar correction function
def correct_grammar(text):
    """
    Correct grammar and spelling using T5 model.
    :param text: Input text
    :return: Corrected text
    """
    try:
        # Split text into smaller chunks for processing
        chunks = [text[i:i + 200] for i in range(0, len(text), 200)]
        corrected_text = []

        # Process each chunk
        for chunk in chunks:
            input_text = f"grammar correction: {chunk}"
            input_ids = tokenizer.encode(input_text, return_tensors="pt")
            outputs = model.generate(input_ids, max_length=512)
            corrected_chunk = tokenizer.decode(outputs[0], skip_special_tokens=True)
            corrected_text.append(corrected_chunk)

        # Join corrected chunks
        return " ".join(corrected_text)
    except Exception as e:
        logging.error(f"Error during grammar correction: {e}")
        return text  # Return original text if correction fails

# Function to check inconsistent fonts in a PowerPoint file
def check_inconsistent_fonts(file_path, default_font, output_pptx_path=None):
    """
    Identify inconsistent fonts in a PowerPoint file and highlight them.
    :param file_path: Path to the PowerPoint file
    :param default_font: Expected font
    :param output_pptx_path: Path to save the updated PowerPoint file
    :return: List of issues with fonts
    """
    presentation = Presentation(file_path)
    inconsistent_fonts = []

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        font = getattr(run.font, 'name', None)
                        if font and font != default_font:
                            # Record inconsistent font issues
                            inconsistent_fonts.append({
                                'slide_number': slide_index,
                                'text': run.text.strip(),
                                'font': font
                            })
                            # Highlight inconsistent text in red
                            if hasattr(run.font, 'color'):
                                run.font.color.rgb = RGBColor(255, 0, 0)

    # Save updated PowerPoint file
    if output_pptx_path:
        presentation.save(output_pptx_path)
    return inconsistent_fonts

# Function to save issues to a CSV file
def save_issues_to_csv(issues, output_name):
    """
    Save issues to a CSV file.
    :param issues: List of detected issues
    :param output_name: Name of the output CSV file
    :return: Path to the saved CSV file
    """
    output_path = Path.home() / "Downloads" / f"{output_name}.csv"
    try:
        with open(output_path, mode='w', newline='', encoding='utf-8-sig') as file:
            writer = csv.writer(file)
            writer.writerow(["Slide Number", "Issue Type", "Text", "Details"])
            for issue in issues:
                writer.writerow([issue['slide_number'], issue.get('issue', ''), issue['text'], issue.get('font', '')])
        return output_path
    except Exception as e:
        logging.error(f"Error saving CSV: {e}")
        return None

# Main GUI function using Tkinter
def run_tool():
    """
    Function to execute the tool with user inputs from the GUI.
    """
    file_path = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
    default_font = default_font_var.get()

    # Validate inputs
    if not file_path or not file_path.endswith('.pptx'):
        messagebox.showerror("Error", "Invalid PowerPoint file.")
        return
    if not default_font:
        messagebox.showerror("Error", "Select a font.")
        return

    try:
        # Perform validation
        inconsistent_fonts = check_inconsistent_fonts(file_path, default_font)
        if inconsistent_fonts:
            csv_file = save_issues_to_csv(inconsistent_fonts, "Validation_Report")
            messagebox.showinfo("Success", f"Validation complete. Report saved to {csv_file}")
        else:
            messagebox.showinfo("Success", "No issues found.")
    except Exception as e:
        logging.error(f"Error: {e}")
        messagebox.showerror("Error", f"Unexpected error: {e}")

# Initialize Tkinter GUI
root = tk.Tk()
root.title("PPT Validator")
root.geometry("600x400")

# Dropdown for selecting default font
tk.Label(root, text="Select Default Font:").pack()
default_font_var = tk.StringVar()
ttk.Combobox(root, textvariable=default_font_var, values=[f.name for f in font_manager.fontManager.ttflist]).pack()

# Button to run validation
tk.Button(root, text="Run Validation", command=run_tool).pack()
root.mainloop()
