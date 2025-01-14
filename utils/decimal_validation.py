import re
import logging

# Save regex pattern in a variable
decimal_pattern = re.compile(r'\b\d+[\.,]\d+\b')

def validate_decimal_consistency(slide, slide_index, decimal_places):
    issues = []
    
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    text = run.text
                    # Find all decimal numbers with either a dot or comma as the decimal separator
                    matches = decimal_pattern.findall(text)
                    logging.debug(f"Slide {slide_index}: Found matches: {matches}")
                    for match in matches:
                        # Replace comma with dot for consistency
                        match = match.replace(',', '.')
                        # Count the number of digits after the dot
                        decimal_part = match.split('.')[-1]
                        if len(decimal_part) != decimal_places:
                            issues.append({
                                'slide': slide_index,
                                'issue': 'Inconsistent Decimal Points',
                                'text': match,
                                'details': f'Expected {decimal_places} decimal place(s), found {len(decimal_part)} in "{match}".'
                            })
                            logging.debug(f"Slide {slide_index}: Inconsistent decimal points found in \"{match}\". Expected {decimal_places}, found {len(decimal_part)}.")
    
    return issues
