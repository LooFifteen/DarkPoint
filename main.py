import os

from pptx import Presentation
from pptx.dml.color import RGBColor

BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)

file_name = os.getenv("FILE_NAME")

if file_name is None: file_name = input("Enter file name to darken >> ")
presentation = Presentation(file_name)

for slide in presentation.slides:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = BLACK
    for shape in slide.shapes:
        # TODO: Change outline of shape to #WHITE
        if not shape.has_text_frame: continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                font = run.font
                font.color.rgb = WHITE

new_file_name = f"dark.{file_name}"
presentation.save(new_file_name)
print(f"Successfully saved as: {new_file_name}")
