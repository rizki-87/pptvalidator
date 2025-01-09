# utils/spelling_validation.py

import re
import string
from spellchecker import SpellChecker
from config import TECHNICAL_TERMS  

spell = SpellChecker()
spell.word_frequency.load_words(TECHNICAL_TERMS)

def is_exempted(word, TECHNICAL_TERMS):
    return word in TECHNICAL_TERMS or re.match(r"^\d+\+?$", word)

def validate_spelling_slide(slide, slide_index):
    issues = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    words = re.findall(r"\b[\w+]+\b", run.text)
                    for word in words:
                        clean_word = word.strip(string.punctuation)
                        if is_exempted(clean_word, TECHNICAL_TERMS):
                            continue
                        if clean_word.lower() not in spell:
                            correction = spell.correction(clean_word)
                            if correction and correction != clean_word:
                                issues.append({
                                    'slide': slide_index,
                                    'issue': 'Misspelling',
                                    'text': word,
                                    'corrected': correction
                                })
    return issues

def validate_spelling_in_text(text, slide_index):
    issues = []
    words = re.findall(r"\b[\w+]+\b", text)
    for word in words:
        clean_word = word.strip(string.punctuation)
        if clean_word.lower() not in spell:
            correction = spell.correction(clean_word)
            if correction and correction != clean_word:
                issues.append({
                    'slide': slide_index,
                    'issue': 'Misspelling',
                    'text': word,
                    'corrected': correction
                })
    return issues
