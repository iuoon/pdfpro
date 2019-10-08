import pdfkit
import os

if __name__ == '__main__':
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    path_wk = BASE_DIR + '\\w\\wkhtmltopdf.exe'
    config = pdfkit.configuration(wkhtmltopdf=path_wk)
    with open("test.html", 'r', encoding='utf-8') as f:
        content = f.read()
    content = content.replace("{替换标签}", str("替换结果"))
    options = {
        'encoding': "utf-8",
        'page-height': '80mm',  # 设置宽
        'page-width': '100mm',  # 设置高
        'margin-top': '0mm',
        'margin-right': '0mm',
        'margin-bottom': '0mm',
        'margin-left': '0mm'
    }
    pdfkit.from_string(content, "test.pdf", options=options, configuration=config)
