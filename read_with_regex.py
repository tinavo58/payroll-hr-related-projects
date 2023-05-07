import re
from pypdf import PdfReader


wPDF = r"test/ATO_tax_tables/Weekly-tax-table-from-13-October-2020.pdf"


def data_extraction(filePath, pattern=r".+00"):
    reader = PdfReader(filePath)
    for numPage in range(len(reader.pages)):
        page = reader.pages[numPage].extract_text()
        data_required = re.findall(pattern, page)
        for each in data_required:
            if len(each.split()) == 3:
                yield tuple(float(re.sub(',', '', x)) for x in each.split())
