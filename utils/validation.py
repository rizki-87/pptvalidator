import re          
import logging          
import string          
import pandas as pd  # Ensure pandas is imported      
from utils.spelling_validation import validate_spelling_slide, validate_spelling_in_text          
from utils.million_notation_validation import validate_million_notations  # Ensure this is present    
    
def validate_tables(slide, slide_index):      
    issues = []          
    for shape in slide.shapes:          
        if shape.has_table:          
            table = shape.table          
            for row in table.rows:          
                for cell in row.cells:          
                    # Validate text within the cell          
                    text = cell.text.strip()          
                    if text:  # If there is text          
                        issues.extend(validate_spelling_in_text(text, slide_index))          
              
    # Validate million notations using the slide object        
    issues.extend(validate_million_notations(slide, slide_index))  # Call the new function with the slide object          
              
    return issues          
    
def validate_charts(slide, slide_index):      
    issues = []          
    for shape in slide.shapes:          
        if shape.has_chart:          
            chart = shape.chart          
            # Validate data within the chart          
            for series in chart.series:          
                for point in series.points:          
                    label = point.data_label.text.strip()          
                    if label:          
                        issues.extend(validate_spelling_in_text(label, slide_index))          
            # If the chart has data displayed in a table, validate it as well          
            if chart.has_data_table:          
                for row in chart.data_table.rows:          
                    for cell in row.cells:          
                        text = cell.text.strip()          
                        if text:          
                            issues.extend(validate_spelling_in_text(text, slide_index))          
              
    # Validate million notations using the slide object        
    issues.extend(validate_million_notations(slide, slide_index))  # Call the new function with the slide object          
              
    return issues  
