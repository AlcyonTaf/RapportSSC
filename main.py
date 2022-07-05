# -*- coding: utf-8 -*-
import docx
from lxml import etree as et
from docx.document import Document

try:
    document = Document()
except TypeError:
    from docx import Document

    document = Document()

from docx.oxml.shared import qn
from docx.oxml.text.paragraph import CT_P


# Le but va etre de mettre en forme le rapport des essais SSC de la corrosion
# TOdo : Premiere chose a faire : Parcourir toutes les tables provisoire pour extraire les données et les stocker, ensuite on les supprimes



# Trouver un bookmark : https://stackoverflow.com/questions/24965042/python-docx-insertion-point
def get_bookmark_par_element(document2, bookmark_name):
    """
    Return the named bookmark parent paragraph element. If no matching
    bookmark is found, the result is '1'. If an error is encountered, '2'
    is returned.
    """
    doc_element = document2.part._element
    bookmarks_list = doc_element.findall('.//' + qn('w:bookmarkStart'))
    for bookmark in bookmarks_list:
        name = bookmark.get(qn('w:name'))
        if name == bookmark_name:
            par = bookmark.getparent()
            if not isinstance(par, docx.oxml.CT_P):
                return 2
            else:
                return par
    return 1


def bookmark_text(
        doc,
        bookmark_name,
        text,
        underline=False,
        italic=False,
        bold=False,
        style=None,
        header=False, ):
    doc_element = (
        doc.sections[0].header.part._element
        if header is True
        else doc._part._element
    )
    #TODO : Voir pour chercher également bookmarkend. s'en servir pour supprimer le contenue entre start et end
    bookmarks_list = doc_element.findall(".//" + qn("w:bookmarkStart"))
    for bookmark in bookmarks_list:
        name = bookmark.get(qn("w:name"))
        if name == bookmark_name:
            par = bookmark.getparent()
            if not isinstance(par, CT_P):
                return False
            else:
                # for elem in par.iter():
                #     print("%s - %s" % (elem.tag, elem.text))
                # print(et.tostring(par, pretty_print=True))
                i = par.index(bookmark) - 1
                p = doc.add_paragraph()
                run = p.add_run(text, style)
                run.underline = underline
                run.italic = italic
                run.bold = bold
                par.insert(i, run._element)
                p = p._element
                p.getparent().remove(p)
                p._p = p._element = None
                # Essai pour suppresion bookmarks + text
                for i in range(2, -1, -1):
                    # print(i)
                    asup = par[par.index(bookmark) + i]
                    #
                    #print(et.tostring(asup, pretty_print=True))
                    par.remove(asup)

                # par.remove(bookmark)
                # print(et.tostring(par, pretty_print=True))
                return True
    return False


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    doc = Document(".\ExtractionLims\Test Rapport Essai corrosion SSC - A22_00202 - 2022-07-05.docx")
    all_paras = doc.paragraphs
    print(len(all_paras))

    test = get_bookmark_par_element(doc, "BK_Condition_Solution")
    # print(test.r)

    test2 = bookmark_text(doc, "BK_Condition_Solution", " essai remplacement")

    # Exemple pour recupérer les données d'un tableau
    tb = doc.tables[0]
    # data =[]
    # for i, row in enumerate(tb.rows):
    #     print(i)
    #     text = (cell.text for cell in row.cells)
    #
    #     # Establish the mapping based on the first row
    #     # headers; these will become the keys of our dictionary
    #     if i == 0:
    #         keys = tuple(text)
    #         print(keys)
    #         continue
    #
    #     # Construct a dictionary for this row, mapping
    #     # keys to values for this row
    #     row_data = dict(zip(keys, text))
    #     print(row_data)
    #     data.append(row_data)
    #
    # print(data)

    doc.save("result.docx")
