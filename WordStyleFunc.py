#!/usr/bin/python
# -*- coding: utf-8 -*-


from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
from docx.text.font import Font
from docx.text.parfmt import ParagraphFormat
from docx.table import Table


def set_cell_border(cell, **kwargs):
    """
    Set cell`s border
    Usage:
    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
        bottom={"sz": 12, "color": "#00FF00", "val": "single"},
        left={"sz": 24, "val": "dashed", "shadow": "true"},
        right={"sz": 12, "val": "dashed"},
    )
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # check for tag existnace, if none found, then create one
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    # list over all available tags
    for edge in ('left', 'top', 'right', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)

            # check for tag existnace, if none found, then create one
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)

            # looks like order of attributes is important
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))


def set_font_format(font: Font, **kwargs):
    """
    set font parameter
    Usage:
    set_font(font, font_param={'size': 22, 'name': '华文中宋'})
    设置字体: 华文中宋, 字体大小:22
    """
    if isinstance(font, Font):
        # list over all available tags
        for edge in ('font_format',):
            edge_data = kwargs.get(edge)
            if edge_data:
                # looks like order of attributes is important
                for key in ["size", "name", "bold"]:
                    if key in edge_data:
                        if key == 'size':
                            font.size = Pt(edge_data[key])
                        elif key == 'name':
                            font.name = edge_data[key]
                            font.element.rPr.rFonts.set(qn('w:eastAsia'), edge_data[key])  # 东亚地区的字符设置
                        elif key == 'bold':
                            font.bold = edge_data[key]


def set_paragraph_format(pformat: ParagraphFormat, **kwargs):
    if isinstance(pformat, ParagraphFormat):
        # list over all available tags
        for edge in ('paragraph_format',):
            edge_data = kwargs.get(edge)
            if edge_data:
                # looks like order of attributes is important
                for key in ["size", "name", "bold"]:
                    if key in edge_data:
                        if key == 'space_before':
                            pformat.space_before = Pt(edge_data[key])
                        elif key == 'space_after':
                            pformat.space_after = Pt(edge_data[key])
                        elif key == 'line_spacing_rule':
                            pformat.line_spacing_rule = edge_data[key]
                        elif key == 'line_spacing':
                            pformat.line_spacing = Pt(edge_data[key])


def merge(table: Table, row1, col1, row2, col2):
    if isinstance(table, Table):
        table.rows[row1].cells[col1].merge(table.rows[row2].cells[col2])
