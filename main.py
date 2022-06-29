# -*- coding: utf-8 -*-

from docx.document import Document
try:
    document = Document()
except TypeError:
    from docx import Document
    document = Document()

# Le but va etre de mettre en forme le rapport des essais SSC de la corrosion


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    doc = Document("test.docx")
    all_paras = doc.paragraphs
    print(len(all_paras))
    tb = doc.tables[0]

    data =[]
    for i, row in enumerate(tb.rows):
        print(i)
        text = (cell.text for cell in row.cells)

        # Establish the mapping based on the first row
        # headers; these will become the keys of our dictionary
        if i == 0:
            keys = tuple(text)
            print(keys)
            continue

        # Construct a dictionary for this row, mapping
        # keys to values for this row
        row_data = dict(zip(keys, text))
        print(row_data)
        data.append(row_data)

    print(data)



