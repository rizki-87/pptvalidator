# utils/million_notation_validation.py

import re
import logging

def validate_million_notations(slide, slide_index, notation='m', decimal_places=0):
    issues = []
    
    # Determines the regex pattern based on the selected notation.
    if notation.lower() == 'm':
        pattern = r'\b\d+\s?m\b'
    elif notation.lower() == 'mn':
        pattern = r'\b\d+\s?Mn\b'
    else:
        pattern = r'\b\d+\s?M\b'  

    notation_set = set()
    all_matches = []
    logging.debug(f"Slide {slide_index}: Checking shapes for million notations")

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                matches = re.findall(pattern, run.text, re.IGNORECASE)
                all_matches.extend(matches)
                for match in matches:
                    notation_set.add(match.strip())
    
    if len(notation_set) > 1:
        for match in all_matches:
            issues.append({
                'slide': slide_index,
                'issue': 'Inconsistent Million Notations',
                'text': match,
                'details': f'Found inconsistent million notations: [using {", ".join(notation_set)}]'
            })

    return issues
