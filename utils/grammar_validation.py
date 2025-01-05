# utils/grammar_validation.py

import language_tool_python

def initialize_language_tool():
    try:
        return language_tool_python.LanguageToolPublicAPI('en-US')
    except Exception as e:
        st.error(f"LanguageTool initialization failed: {e}")
        return None

def validate_grammar_slide(slide, slide_index, grammar_tool):
    issues = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    text = run.text.strip()
                    if text and grammar_tool:
                        matches = grammar_tool.check(text)
                        for match in matches:
                            issues.append({
                                'slide': slide_index,
                                'issue': 'Grammar Error',
                                'text': text,
                                'corrected': match.replacements
                            })
    return issues
