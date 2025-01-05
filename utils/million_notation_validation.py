# utils/million_notation_validation.py

import re
import logging  

def validate_million_notations(slide, slide_index):
    issues = []
    million_patterns = {
        r'\b\d+M\b': 'M', r'\b\d+\s?Million\b': 'Million', r'\b\d+mn\b': 'mn', r'\b\d+\sm\b': 'm',
        r'\b\d+MM\b': 'MM', r'\b\d+\s?Millions\b': 'Millions', r'\b\d+\s?Juta\b': 'Juta'
    }
    notation_set = set()
    all_matches = []
    logging.debug(f"Slide {slide_index}: Checking shapes for million notations")  
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                for pattern, notation in million_patterns.items():
                    matches = re.findall(pattern, run.text, re.IGNORECASE)
                    all_matches.extend(matches)
                    for match in matches:
                        notation_set.add(notation)
    if len(notation_set) > 1:
        for match in all_matches:
            issues.append({
                'slide': slide_index,
                'issue': 'Inconsistent Million Notations',
                'text': match,
                'details': f'Found inconsistent million notations: [using {", ".join(notation_set)}]'
            })
    return issues
