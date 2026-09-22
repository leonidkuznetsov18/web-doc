"""Regenerate oversized-inline-image.docx with python-docx 1.2.0 and Pillow 12.3.0.

Run `python tests/fixtures/docx/generate.py` from the repository root. These
dependencies are needed only to regenerate the fixture, not to run the browser
test. The document and image are synthetic.
The 21-inch image must fit the 6.5-inch content width without losing its 4:1
aspect ratio. An engine upgrade must preserve the viewer's image-fitting step.
"""

from io import BytesIO
from pathlib import Path

from docx import Document
from docx.shared import Inches
from PIL import Image


def main() -> None:
    image = BytesIO()
    Image.new("RGB", (1600, 400), "red").save(image, format="PNG")
    image.seek(0)
    document = Document()
    section = document.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.left_margin = section.right_margin = Inches(1)
    section.top_margin = section.bottom_margin = Inches(1)
    document.add_picture(image, width=Inches(21))
    document.core_properties.author = "web-doc regression tests"
    document.core_properties.title = "Oversized inline image regression"
    document.save(str(Path(__file__).with_name("oversized-inline-image.docx")))


if __name__ == "__main__":
    main()
