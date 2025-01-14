import re  
import logging  # Pastikan ini ada  
  
def validate_million_notations(slide, slide_index):  
    issues = []  
    million_patterns = {  
        r'\b[€$£]?\s*\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*[Mm]\b': 'M',  # M atau m  
        r'\b[€$£]?\s*\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*Million\b': 'Million',   
        r'\b[€$£]?\s*\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*mn\b': 'mn',   
        r'\b[€$£]?\s*\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*m\b': 'm',  
        r'\b[€$£]?\s*\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*MM\b': 'MM',   
        r'\b[€$£]?\s*\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*Millions\b': 'Millions',   
        r'\b[€$£]?\s*\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*Juta\b': 'Juta'  
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
  
    # Cek konsistensi notasi  
    if len(notation_set) > 1:  
        # Hanya catat masalah unik  
        unique_matches = set(all_matches)  # Menggunakan set untuk menghindari duplikasi  
        for match in unique_matches:  
            issues.append({  
                'slide': slide_index,  
                'issue': 'Inconsistent Million Notations',  
                'text': match,  
                'details': f'Found inconsistent million notations: [using {", ".join(notation_set)}]'  
            })  
    return issues  
