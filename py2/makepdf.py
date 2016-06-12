import trml2pdf
import os
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

pdfmetrics.registerFont(TTFont('HelveticaNeue-UltraLight', 'helveticaneuelight-webfont.ttf'))

path = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) + "/templates/daily_report.rml"
output_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) + "/output/output.pdf"

content = trml2pdf.parseString(open(path).read())

with open(output_path, "a") as myfile:
    myfile.write(content)