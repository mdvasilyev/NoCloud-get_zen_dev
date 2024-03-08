import datetime
from docx import Document
from docx2pdf import convert
from python_docx_replace import docx_replace


def main():
    template_path: str = "Act-template.docx"
    monthsConvert = {
        "January": "января",
        "February": "февраля",
        "March": "марта",
        "April": "апреля",
        "May": "мая",
        "June": "июня",
        "July": "июля",
        "August": "августа",
        "September": "сентября",
        "October": "октября",
        "November": "ноября",
        "December": "декабря"
    }
    today = datetime.date.today()
    monthEn = today.strftime("%B")
    monthRu = monthsConvert.get(monthEn)
    properties: dict = {
        "TODAY-RU": today.strftime(f"«%d» {monthRu} %Y г."),
        "TODAY-EN": today.strftime("“%d” %B %Y"),
        "PERIOD-START": "01.02.2024",
        "PERIOD-END": "28.02.2024"
    }
    doc = Document(template_path)
    docx_replace(doc, **properties)
    doc.save("result.docx")
    convert("result.docx", "result.pdf")


if __name__ == "__main__":
    main()
