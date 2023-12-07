import json

from docx import Document
from docx.shared import Mm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import config

@staticmethod
def setup_page(document: Document, page_size: str = 'a4'):
    page_param = json.load(open(f"{config.get_config('page_presets')}/{str(page_size)}.json"))
    
    section = document.sections[0]
    section.page_height = Mm(page_param['page_height'])
    section.page_width = Mm(page_param['page_width'])
    section.left_margin = Mm(page_param['margin_left'])
    section.right_margin = Mm(page_param['margin_right'])
    section.top_margin = Mm(page_param['margin_top'])
    section.bottom_margin = Mm(page_param['margin_bottom'])

    return document

@staticmethod
def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # check for tag existnace, if none found, then create one
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    # list over all available tags
    for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)

            # check for tag existnace, if none found, then create one
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)

            # looks like order of attributes is important
            for key in ["sz", "val", "color", "space"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))