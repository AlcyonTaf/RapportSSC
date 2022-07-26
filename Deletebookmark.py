# -*- coding: utf-8 -*-
import itertools

import docx
from lxml import etree as et
from docx.document import Document

try:
    document = Document()
except TypeError:
    from docx import Document

    document = Document()

from docx.table import _Cell, Table
from docx.document import Document as _Document
from docx.oxml.table import CT_Tbl
from docx.text.paragraph import Paragraph
from docx.oxml.shared import qn
from docx.oxml.text.paragraph import CT_P
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def is_root(element):
    if element.getparent() is None:
        return True
    else:
        return False


def delete_bookmark(
        doc,
        bookmark_name,
        header=False):
    doc_element = (
        doc.sections[0].header.part._element
        if header is True
        else doc._part._element
    )

    find = False
    element_to_remove = []
    for element in doc_element.iter():
        if element.tag == qn("w:bookmarkStart") and element.get(qn("w:name")) == bookmark_name:
            #print("%s - %s - %s" % (element.tag, element.text, element.get(qn("w:name"))))
            #print(element.get(qn("w:id")))
            Bk_id = element.get(qn("w:id"))
            find = True

        if find:
            elem_loop = element
            # print(element.tag + " - " + qn("w:body>"))
            # On considere que tous les éléments a supprimer sont des enfants de w:body
            # On va donc remonter depuis l'element en cours pour trouvé le parent qui est enfant de w:body
            while True:
                # print(elem_loop.tag + " - " + qn("w:body>"))
                if elem_loop.getparent().tag == qn("w:body"):
                    # print('trouvé')
                    element_to_remove.append(elem_loop)
                    break
                else:
                    # On remonte au parent
                    elem_loop = elem_loop.getparent()
                    # print(elem_loop)

            if element.tag == qn("w:bookmarkEnd") and element.get(qn("w:id")) == Bk_id:
                # print('Fin de notre BK')
                # print(is_root(element.getparent()))
                # print(is_root(element.getparent().getparent()))
                find = False

    # On supprime les doublons
    element_to_remove = list(dict.fromkeys(element_to_remove))

    body = doc_element.getchildren()[0]
    # On boucle sur les éléments a supprimer
    for elem_supp in element_to_remove:
        body.remove(elem_supp)


doc = Document(".\ExtractionLims\Test Rapport Essai corrosion - modif.docx")

delete_bookmark(doc, 'BK_Delete_1')


doc.save("resultdelete.docx")