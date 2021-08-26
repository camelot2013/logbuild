#!/usr/bin/python
# -*- coding: utf-8 -*-


from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
from docx.text.font import Font
from docx.text.parfmt import ParagraphFormat
import xlwt


def thin_border():
    borders = xlwt.Borders()  # Create Borders
    # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1

    return borders
