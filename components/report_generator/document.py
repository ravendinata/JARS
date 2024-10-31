import json

from docx import Document
from docx.shared import Mm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import config

@staticmethod
def setup_page(document: Document, page_size: str = 'a4'):
    """
    Sets up the page size of the document.
    
    Args:
        document (Document): The document to setup.
        page_size (str, optional): The page size to setup. Defaults to 'a4'.

    Returns:
        Document: The document with the page size setup.

    Raises:
        FileNotFoundError: When the page size preset file is not found.
    """
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
    """
    Set a table cell's border properties. Usage example:
    >>> set_cell_border(
    >>>     cell,
    >>>     top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
    >>>     bottom={"sz": 12, "color": "#00FF00", "val": "single"},
    >>>     start={"sz": 24, "val": "dashed", "shadow": "true"},
    >>>     end={"sz": 12, "val": "dashed"},
    >>> )

    Args:
        cell (docx.table._Cell): The cell to set the border of.
        **kwargs: The border properties to set.

    Raises:
        ValueError: When an invalid border property is passed.
    """
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