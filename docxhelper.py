import docx

#https://github.com/virajkanwade/python-docx/commit/7cb692a14b8010a04bc0945fc6b4a149a045f776

def numpr_set_ilvl(self, val):
    ilvl = self.get_or_add_ilvl()
    ilvl.val = val

def numpr_set_numId(self, val):
    numId = self.get_or_add_numId()
    numId.val = val

def ctpr_set_numbering(self, ilvl=0, numId=10):
    numPr = self.get_or_add_numPr()
    numPr.numpr_set_ilvl(ilvl)
    numPr.numpr_set_numId(numId)

def para_set_numbering(self, ilvl=0, numId=10):
    pPr = self.get_or_add_pPr()
    pPr.ctpr_set_numbering(ilvl, numId)

# TODO make it colored
def add_hyperlink(paragraph, url, text):
    """
    A function that places a hyperlink within a paragraph object.

    :param paragraph: The paragraph we are adding the hyperlink to.
    :param url: A string containing the required url
    :param text: The text displayed for the url
    :return: The hyperlink object
    """

    # This gets access to the document.xml.rels file and gets a new relation id value
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    # Create the w:hyperlink tag and add needed values
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id, )

    # Create a w:r element
    new_run = docx.oxml.shared.OxmlElement('w:r')

    # Create a new w:rPr element
    rPr = docx.oxml.shared.OxmlElement('w:rPr')

    # Join all the xml elements together add add the required text to the w:r element
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)

    r = paragraph.add_run()
    r._r.append(hyperlink)
    r.font.color.theme_color = docx.enum.dml.MSO_THEME_COLOR_INDEX.HYPERLINK
    r.font.underline = True

    #paragraph._p.append(hyperlink)

    return hyperlink