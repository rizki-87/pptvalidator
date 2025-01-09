import re
import logging
import string
from utils.spelling_validation import validate_spelling_slide, validate_spelling_in_text  # Pastikan ini diimpor

def validate_tables(slide, slide_index):
    issues = []
    for shape in slide.shapes:
        if shape.has_table:
            table = shape.table
            for row in table.rows:
                for cell in row.cells:
                    # Validasi teks di dalam sel
                    text = cell.text.strip()
                    if text:  # Jika ada teks
                        issues.extend(validate_spelling_in_text(text, slide_index))
    return issues

def validate_charts(slide, slide_index):
    issues = []
    for shape in slide.shapes:
        if shape.has_chart:
            chart = shape.chart
            # Validasi data di dalam chart
            for series in chart.series:
                for point in series.points:
                    label = point.data_label.text.strip()
                    if label:
                        issues.extend(validate_spelling_in_text(label, slide_index))
            # Jika chart memiliki data yang ditampilkan dalam tabel, validasi juga
            if chart.has_data_table:
                for row in chart.data_table.rows:
                    for cell in row.cells:
                        text = cell.text.strip()
                        if text:
                            issues.extend(validate_spelling_in_text(text, slide_index))
    return issues
