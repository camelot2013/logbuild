#!/usr/bin/python
# -*- coding: utf-8 -*-


import random
from WorkLogXls import *
from ExcelStyleFunc import *
import xlwt


def head1_style():
    style = xlwt.XFStyle()  # 初始化样式
    borders = thin_border()
    font = xlwt.Font()  # 为样式创建字体
    font.name = '宋体'
    font.bold = True  # 黑体
    font.colour_index = 2
    # 字体大小，11为字号，20为衡量单位
    font.height = 20 * 16
    style.borders = borders
    style.font = font
    alignment = xlwt.Alignment()  # Create Alignment
    alignment.horz = xlwt.Alignment.HORZ_CENTER  # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER  # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    style.alignment = alignment  # Add Alignment to Style
    # 设置背景颜色
    pattern = xlwt.Pattern()
    # 设置背景颜色的模式
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    # 背景颜色
    pattern.pattern_fore_colour = 40  # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
    style.pattern = pattern

    return style


def head2_style():
    style = xlwt.XFStyle()  # 初始化样式
    borders = thin_border()
    font = xlwt.Font()  # 为样式创建字体
    font.name = '宋体'
    font.bold = True  # 黑体
    # 字体大小，11为字号，20为衡量单位
    font.height = 20 * 10
    style.borders = borders
    style.font = font
    alignment = xlwt.Alignment()  # Create Alignment
    alignment.horz = xlwt.Alignment.HORZ_CENTER  # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER  # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    # 设置自动换行
    alignment.wrap = 1
    style.alignment = alignment  # Add Alignment to Style
    # 设置背景颜色
    pattern = xlwt.Pattern()
    # 设置背景颜色的模式
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    # 背景颜色
    pattern.pattern_fore_colour = 5     # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
    style.pattern = pattern

    return style


def body_style():
    style = xlwt.XFStyle()  # 初始化样式
    borders = thin_border()
    font = xlwt.Font()  # 为样式创建字体
    font.name = '宋体'
    font.bold = False  # 黑体
    # 字体大小，11为字号，20为衡量单位
    font.height = 20 * 10
    style.borders = borders
    style.font = font
    alignment = xlwt.Alignment()  # Create Alignment
    alignment.horz = xlwt.Alignment.HORZ_CENTER  # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER  # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    # 设置自动换行
    alignment.wrap = 1
    style.alignment = alignment  # Add Alignment to Style

    return style


def get_code_lines(workload: int):
    read_code_lines = workload * random.randint(10, 12) + random.randint(0, 20)
    write_code_lines = workload * random.randint(7, 11) + random.randint(3, 15)

    return read_code_lines, write_code_lines


def get_work_content(demand_no: str, logs: list):
    work_content = ''
    demand_name = ''
    workload = 0
    tmp_content = []
    for log in logs:
        if demand_no == log['demand_no']:
            demand_name = log['demand_name']
            workload = workload + log['workload']
            if log['work_content'] not in tmp_content:
                tmp_content.append(log['work_content'])
    work_content = demand_name
    for content in tmp_content:
        work_content = work_content + '\r\n  '+content
    return work_content, workload


def __sort_key_demand(elem):
    return elem['demand_no']


def creat_month_log(month_log: dict):
    month_log['logs'].sort(key=__sort_key_demand)
    tmp_demand_no = ''
    demand_count = 0

    year_str = month_log['month'][:4]
    mon_str = month_log['month'][4:]
    person_name = month_log['person_name']
    file_name = '人月外包人员月报-{}年{}月-{}.xls'.format(year_str, mon_str, person_name)
    person_cfg = read_cfg()
    # 创建工作簿
    wb = xlwt.Workbook()
    # 创建表单
    sh = wb.add_sheet('开发人员工作量统计')

    sh.write_merge(0, 0, 0, 13, '人月开发人员工作量统计表-{}年{}月'.format(year_str, mon_str), head1_style())
    sh.write_merge(1, 2, 0, 0, '序号', head2_style())
    style1 = head2_style()
    style1.font.colour_index =2
    sh.write_merge(1, 2, 1, 1, '年月', style1)
    sh.write_merge(1, 2, 2, 2, '公司', head2_style())
    sh.write_merge(1, 2, 3, 3, '姓名', head2_style())
    sh.write_merge(1, 2, 4, 4, '岗位', head2_style())
    sh.write_merge(1, 2, 5, 5, '级别', head2_style())
    sh.write_merge(1, 2, 6, 6, '月单价', head2_style())
    sh.write_merge(1, 2, 7, 7, '需求编号/生产问题', head2_style())
    sh.write_merge(1, 2, 8, 8, '需求/问题及实现概述', head2_style())
    sh.write_merge(1, 1, 9, 10, '代码量（行）', head2_style())
    sh.write(2, 9, '阅读', head2_style())
    sh.write(2, 10, '编写', head2_style())
    sh.write_merge(1, 2, 11, 11, '工作量（人天）', head2_style())
    sh.write_merge(1, 2, 12, 12, '进度', head2_style())
    sh.write_merge(1, 2, 13, 13, '合计工作量（人天）', head2_style())
    sh.row(0).height = Pt(27)
    tot_workload = 0
    base_row_idx = 3
    for log in month_log['logs']:
        work_content = log['work_content']
        tot_workload = tot_workload + log['workload']
        if tmp_demand_no != log['demand_no']:
            tmp_demand_no = log['demand_no']
            demand_count = demand_count + 1
            sh.write(base_row_idx, 7, tmp_demand_no, body_style())
            style2 = body_style()
            style2.alignment.horz = xlwt.Alignment.HORZ_LEFT
            work_content, workload = get_work_content(tmp_demand_no, month_log['logs'])
            sh.write(base_row_idx, 8, work_content, style2)
            sh.write(base_row_idx, 11, workload / 8, body_style())
            read_code_lines, write_code_lines = get_code_lines(workload)
            sh.write(base_row_idx, 9, read_code_lines, body_style())
            sh.write(base_row_idx, 10, write_code_lines, body_style())
            sh.write(base_row_idx, 12, ' ', style2)
            base_row_idx = base_row_idx + 1

    sh.write_merge(3, 3 + demand_count - 1, 0, 0, '1', body_style())
    sh.write_merge(3, 3 + demand_count - 1, 1, 1, month_log['month'], body_style())

    sh.write_merge(3, 3 + demand_count - 1, 2, 2, person_cfg['person_info'].get(person_name).get('company'), body_style())
    sh.write_merge(3, 3 + demand_count - 1, 3, 3, person_name, body_style())
    sh.write_merge(3, 3 + demand_count - 1, 4, 4, person_cfg['person_info'].get(person_name).get('station'), body_style())
    sh.write_merge(3, 3 + demand_count - 1, 5, 5, person_cfg['person_info'].get(person_name).get('level'), body_style())
    sh.write_merge(3, 3 + demand_count - 1, 6, 6, person_cfg['person_info'].get(person_name).get('MonthlyUnitPrice'), body_style())
    sh.write_merge(3, 3 + demand_count - 1, 13, 13, tot_workload / 8, body_style())

    # 设置列宽，一个中文等于两个英文等于两个字符，11为字符数，256为衡量单位
    sh.col(0).width = 5 * 256
    sh.col(1).width = 8 * 256
    sh.col(2).width = 16 * 256
    sh.col(3).width = 7 * 256
    sh.col(4).width = 5 * 256
    sh.col(5).width = 10 * 256
    sh.col(6).width = 7 * 256
    sh.col(7).width = 18 * 256
    sh.col(8).width = 62 * 256
    sh.col(9).width = 9 * 256
    sh.col(10).width = 8 * 256
    sh.col(11).width = 9 * 256
    sh.col(12).width = 10 * 256
    sh.col(13).width = 12 * 256

    log_out_dir = get_out_dir()
    wb.save(os.path.join(log_out_dir, file_name))
