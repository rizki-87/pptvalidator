# utils/highlight.py

from pptx.dml.color import RGBColor

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
