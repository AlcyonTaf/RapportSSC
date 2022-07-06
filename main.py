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

from docx.oxml.shared import qn
from docx.oxml.text.paragraph import CT_P


# Le but va etre de mettre en forme le rapport des essais SSC de la corrosion
# TOdo : Premiere chose a faire : Parcourir toutes les tables provisoire pour extraire les données et les stocker, ensuite on les supprimes


def get_data_from_table():
    """
    Le but est de récupérer le contenue des tables provisoires et ensuite les supprimer
    Position des tables :
    :return:
    """
    # Nom et position des tableaux à récupérer
    table_list = {'items_essai': 1, 'items_parent': 2, 'conditions_essais': 3, 'resultats_essais': 5}
    table_valeur = {}

    for nom_table, position in table_list.items():
        # Exemple pour recupérer les données d'un tableau
        get_table = doc.tables[position]

        # pour mettre dans une liste :
        data = [['' for i in range(len(get_table.columns))] for j in range(len(get_table.rows))]
        for y, row in enumerate(get_table.rows):
            for j, cell in enumerate(row.cells):
                if cell.text:
                    data[y][j] = cell.text

        table_valeur[nom_table] = data

    return table_valeur


def clean_table_data():
    """
    On va faire le trie/nettoyage dans certaines tables ici
    :param dict_tables:
    :return: dict nettoyer
    """
    dict_tables = get_data_from_table()
    # items_parent :
    # On va supprimer les doublons puis vérifier que la table contient bien uniquement 2 lignes, sinon c'est qu'il y a plusieurs items parents!
    #Todo : Voir pour vérifier également que cette table contient bien des données.
    items_parent = dict_tables['items_parent']
    items_parent = list(items_parent for items_parent, _ in itertools.groupby(items_parent))
    try:
        if len(items_parent) == 2:
            dict_tables['items_parent'] = items_parent
            print(dict_tables['items_parent'])
        else:
            # Plus ou moins de 2 ligne = probleme!!
            raise ValueError
    except ValueError:
        print("Les items n'ont pas tous le même parents!")

    # items_essai :
    # On va lui donner la forme du tableau de destination :
        # Concatenation des dimensions => Position 2
        # Concaténation des positons => Position 0
        # Ref client => Position 1




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
    # TODO : Voir pour chercher également bookmarkend. s'en servir pour supprimer le contenue entre start et end
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
                    # print(et.tostring(asup, pretty_print=True))
                    par.remove(asup)

                # par.remove(bookmark)
                # print(et.tostring(par, pretty_print=True))
                return True
    return False


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    doc = Document(".\ExtractionLims\Test Rapport Essai corrosion SSC - A22_00202 - 2022-07-05.docx")
    all_paras = doc.paragraphs
    #print(len(all_paras))

    #print(get_data_from_table())
    clean_table_data()

    # test = get_bookmark_par_element(doc, "BK_Condition_Solution")
    # print(test.r)

    # test2 = bookmark_text(doc, "BK_Condition_Solution", " essai remplacement")

    # Exemple pour recupérer les données d'un tableau
    # tb = doc.tables[1]

    # pour mettre dans une liste :
    # list = [['' for i in range(len(tb.columns))] for j in range(len(tb.rows))]
    # for y, row in enumerate(tb.rows):
    #     for j, cell in enumerate(row.cells):
    #         if cell.text:
    #             list[y][j] = cell.text
    #
    # print(list)
    # print(type(list))

    # Pour recuper le contenue dans un dict avec nom de colonne comme clef
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
