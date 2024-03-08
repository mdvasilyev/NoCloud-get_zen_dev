from docx import Document
from python_docx_replace import docx_replace
import datetime


def main():
    template_path: str = "Act-template.docx"
    properties: dict = {
        "TODAY-RU": datetime.date.today().strftime("<<%d>> %b %Y г."),
        "TODAY-EN": datetime.date.today().strftime("``%d'' %B %Y"),
        "PERIOD-START": "01.02.2024",
        "PERIOD-END": "28.02.2024"
    }
    doc = Document(template_path)
    docx_replace(doc, **properties)
    doc.save("result.docx")


if __name__ == "__main__":
    main()
