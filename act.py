import datetime
from docx import Document
from docx2pdf import convert
import sys
from python_docx_replace import docx_replace


def main():
    template_path: str = "Act-template.docx"
    monthsConvert: dict = {
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
    if len(sys.argv) == 1:
        periodEnd = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
        periodStart = periodEnd.replace(day=1)
        periodEnd = periodEnd.strftime("%d.%m.%Y")
        periodStart = periodStart.strftime("%d.%m.%Y")
    else:
        periodStart = sys.argv[1]
        periodEnd = sys.argv[2]
    properties: dict = {
        "TODAY-RU": today.strftime(f"«%d» {monthRu} %Y г."),
        "TODAY-EN": today.strftime("“%d” %B %Y"),
        "PERIOD-START": f"{periodStart}",
        "PERIOD-END": f"{periodEnd}"
    }
    doc = Document(template_path)
    docx_replace(doc, **properties)
    doc.save("result.docx")
    convert("result.docx", "result.pdf")


if __name__ == "__main__":
    main()
