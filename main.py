import docx
from docx.shared import RGBColor


def clarify(acc_string):
    return acc_string.replace('́', '').replace('\n', '<br>').replace('\xa0', '').replace('.00', ':00')


def read_docx(doc: docx.Document):
    print(f"{len(doc.paragraphs)=}")
    for counter, paragraph in enumerate(doc.paragraphs):
        print(f"paragraph[{counter}] = {paragraph.text}")
        # print(paragraph.runs[0].font.color.rgb)
    table = doc.tables[0]
    for row in table.rows:
        for cell in row.cells:
            color_rgb = None
            if cell.paragraphs:
                color_rgb = cell.paragraphs[0].runs[0].font.color.rgb

            fixed_text = clarify(cell.text)
            if color_rgb == docx.shared.RGBColor(0xff, 0x00, 0x00):
                print(f' RED {fixed_text}')
            # else:
            #     print(f'{fixed_text}')


if __name__ == '__main__':
    document_to_parse = docx.Document("test.docx")
    read_docx(document_to_parse)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
