import docx
from docx.shared import RGBColor
from PyQt5 import QtWidgets
from converterui import Ui_MainWindow


def clarify(acc_string):
    return acc_string.replace('́', '').replace('\n', '<br>').replace('\xa0', '').replace('.00', ':00')


class ConverterApp(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.html_content = "" # TODO: https://www.youtube.com/watch?v=0hN6vSSHT0I

    def docx_to_html(self, doc: docx.Document):
        for counter, paragraph in enumerate(doc.paragraphs):
            self.html_content = self.html_content + f"<p>{paragraph.text}</p>\n"
        table = doc.tables[0]
        self.html_content = self.html_content + "<table>\n"
        for counter_rows, row in enumerate(table.rows):
            if counter_rows > 0:
                self.html_content = self.html_content + "<tr>\n"
                for cell in row.cells:
                    color_rgb = None
                    is_bold = None
                    if cell.paragraphs:
                        color_rgb = cell.paragraphs[0].runs[0].font.color.rgb
                        if cell.paragraphs[0].runs[0].font.bold:
                            is_bold = True
                    fixed_text = clarify(cell.text)
                    if color_rgb == docx.shared.RGBColor(0xff, 0x00, 0x00):
                        fixed_text = f'<span style="color:#ff0000;">{fixed_text}</span>'
                    if is_bold:
                        fixed_text = f'<strong>{fixed_text}</strong>'
                    fixed_text = f'<td>{fixed_text}</td>\n'
                    self.html_content = self.html_content + fixed_text
                self.html_content = self.html_content + "</tr>\n"
        self.html_content = self.html_content + "</table>\n"
        return self.html_content


app = QtWidgets.QApplication([])
window = ConverterApp()
window.show()
app.exec_()
