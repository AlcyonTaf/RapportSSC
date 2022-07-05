# Source : https://github.com/python-openxml/python-docx/pull/341

def bookmark_text(
        doc,
        bookmark_name,
        text,
        underline=False,
        italic=False,
        bold=False,
        style=None,
        header=False,
):
    doc_element = (
        doc.sections[0].header.part._element
        if header is True
        else doc._part._element
    )

    bookmarks_list = doc_element.findall(".//" + qn("w:bookmarkStart"))
    for bookmark in bookmarks_list:
        name = bookmark.get(qn("w:name"))
        if name == bookmark_name:
            par = bookmark.getparent()
            if not isinstance(par, CT_P):
                return False
            else:
                i = par.index(bookmark) + 1
                p = doc.add_paragraph()
                run = p.add_run(text, style)
                run.underline = underline
                run.italic = italic
                run.bold = bold
                par.insert(i, run._element)
                p = p._element
                p.getparent().remove(p)
                p._p = p._element = None
                return True
    return False





def bookmark_table(self, bookmark_name, rows, cols, style=None):
    tb = self.add_table(rows=rows, cols=cols, style=style)
    doc_element = self._part._element
    bookmarks_list = doc_element.findall('.//' + qn('w:bookmarkStart'))
    for bookmark in bookmarks_list:
        name = bookmark.get(qn('w:name'))
        if name == bookmark_name:
            par = bookmark.getparent()
            if not isinstance(par, CT_P):
                return False
            else:
                i = par.index(bookmark) + 1
                par.addnext(tb._element)
                return tb
    return tb


def bookmark_picture(self, bookmark_name, picture):
    doc_element = self._part._element
    bookmarks_list = doc_element.findall('.//' + qn('w:bookmarkStart'))
    for bookmark in bookmarks_list:
        name = bookmark.get(qn('w:name'))
        if name == bookmark_name:
            par = bookmark.getparent()
            if not isinstance(par, CT_P):
                return False
            else:
                i = par.index(bookmark) + 1
                p = self.add_paragraph()
                run = p.add_run()
                run.add_picture(picture)
                par.insert(i, run._element)
                p = p._element
                p.getparent().remove(p)
                p._p = p._element = None
                return True