#!/usr/bin/python
# -*- coding: utf-8 -*-


from WordStyleFunc import *
from docx.shared import Cm  # 导入单位转换函数
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from SeasonLog import get_season_work_content
from WorkLogXls import *


def get_season_log_name(person_name: str, year_and_season: str):
    return '公司季报-思瑞奇-{}({}年{}季度).docx'.format(person_name, year_and_season[:4], year_and_season[4:])


def __set_cell_font(cell, **kwargs):
    paragraphs = cell.paragraphs
    for paragraph in paragraphs:
        for run in paragraph.runs:
            font = run.font
            set_font_format(font, **kwargs)


def create_corp_season_log(season_log: dict, person_cfg: json):
    year_and_season = season_log['year_and_season']
    person_name = season_log['person_name']
    company_name = person_cfg['person_info'].get(person_name).get('company')
    doc1 = init_season_doc(person_name, year_and_season[:4], year_and_season, company_name)

    table = doc1.add_table(rows=1, cols=5, style='Table Grid')
    table.style = 'Table Grid'
    table.width = Cm(15.6)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER  # 表格整体居中

    table.rows[0].cells[0].text = '服务人员\r岗位级别\r（工作地）'
    table.rows[0].cells[1].text = '项目名称/涉及系统'
    table.rows[0].cells[2].text = '需求名称'
    table.rows[0].cells[3].text = '工作时间\r（需列举出每项工作的人天天数）'
    table.rows[0].cells[4].text = '科技项目经理确认'

    table.rows[0].height = Cm(1.17)
    table.rows[0].cells[0].width = Cm(2.31)
    table.rows[0].cells[1].width = Cm(2.30)
    table.rows[0].cells[2].width = Cm(4.46)
    table.rows[0].cells[3].width = Cm(4.85)
    table.rows[0].cells[4].width = Cm(1.68)

    tmp_demand_no = ''
    demand_idx = 1
    tot_workload = 0
    for log in season_log['logs']:
        tot_workload = tot_workload + log['workload']
        if tmp_demand_no != log['demand_no']:
            tmp_demand_no = log['demand_no']
            table.add_row()
            table.rows[demand_idx].cells[2].text = '[' + log['month_request_no'] + ']' + log['demand_name']
            _, workload = get_season_work_content(log['demand_no'], season_log['logs'])
            table.rows[demand_idx].cells[3].text = '{}人天；'.format(workload // 8)
            demand_idx = demand_idx + 1

    table.add_row()
    table.rows[demand_idx].cells[1].text = '合计'
    if tot_workload % 8 == 0:
        table.rows[demand_idx].cells[3].text = '{}人天；'.format(tot_workload // 8)
    else:
        table.rows[demand_idx].cells[3].text = '{}人天；'.format(tot_workload / 8)

    table.rows[1].cells[0].text = '{}:{}岗{}（成都）'.format(person_name,
                                                        person_cfg['person_info'].get(person_name).get('station'),
                                                        person_cfg['person_info'].get(person_name).get('level'))
    table.rows[1].cells[1].text = '综合前端系统（{}人天）'.format(tot_workload // 8)

    for row_idx in range(table.rows.__len__()):
        cells = table.row_cells(row_idx)
        for col_idx in range(cells.__len__()):
            cell = table.cell(row_idx, col_idx)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if col_idx == 0 or col_idx == 1:
                cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            else:
                cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            __set_cell_font(cell, font_format={'size': 12, 'name': '仿宋'})

    merge(table, 1, 0, demand_idx, 0)
    merge(table, 1, 1, demand_idx - 1, 1)
    merge(table, demand_idx, 1, demand_idx, 2)

    log_out_dir = get_out_dir()
    doc1.save(os.path.join(log_out_dir, get_season_log_name(person_name, year_and_season)))
