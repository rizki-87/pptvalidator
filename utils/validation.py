# utils/validation.py

from concurrent.futures import ThreadPoolExecutor
import tempfile
from pathlib import Path
from pptx import Presentation
from pptx.dml.color import RGBColor  
import csv
import logging

def highlight_ppt(input_ppt, output_ppt, issues):
    presentation = Presentation(input_ppt)
    for issue in issues:
        if isinstance(issue, dict):
            slide_index = issue['slide'] - 1
            slide = presentation.slides[slide_index]
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if issue['text'] in run.text:
                                run.font.color.rgb = RGBColor(255, 255, 0)
    presentation.save(output_ppt)

def save_to_csv(issues, output_csv):
    with open(output_csv, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=['slide', 'issue', 'text', 'corrected', 'details'])
        writer.writeheader()
        for issue in issues:
            if isinstance(issue, dict):
                writer.writerow({
                    'slide': issue['slide'],
                    'issue': issue['issue'],
                    'text': issue['text'],
                    'corrected': issue.get('corrected', ''),
                    'details': issue.get('details', '')
                })
