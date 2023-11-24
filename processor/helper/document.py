import json

from docx import Document
from docx.shared import Mm

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