#!/usr/bin/python
# -*- coding: utf-8 -*-


from SeasonLog import *


class CorpSeasonLog(SeasonLog):
    def __init__(self, season_log: dict):
        super(CorpSeasonLog, self).__init__(season_log)
        self.season_log['logs'].sort(key=lambda elem: elem['demand_no'])
        self.__person_name = self.season_log['person_name']
        year_and_season = self.season_log['year_and_season']
        self.__year_num = year_and_season[:4]
        self.__season_num = year_and_season[4:]

    def create_corp_season_log(self):
        table = self.document.add_table(rows=1, cols=5, style='Table Grid')
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
        for log in self.season_log['logs']:
            tot_workload = tot_workload + log['workload']
            if tmp_demand_no != log['demand_no']:
                tmp_demand_no = log['demand_no']
                table.add_row()
                table.rows[demand_idx].cells[2].text = '[' + log['month_request_no'] + ']' + log['demand_name']
                _, workload = self.get_season_work_content(log['demand_no'])
                table.rows[demand_idx].cells[3].text = f'{workload / 8 if workload % 8 else workload // 8}人天；'
                demand_idx = demand_idx + 1

        table.add_row()
        table.rows[demand_idx].cells[1].text = '合计'
        table.rows[demand_idx].cells[3].text = f'{tot_workload / 8 if tot_workload % 8 else tot_workload // 8}人天；'
        table.rows[1].cells[0].text = '{}:{}岗{}（成都）'.format(self.__person_name,
                                                            self.cfg['person_info'].get(self.__person_name).get('station'),
                                                            self.cfg['person_info'].get(self.__person_name).get('level'))
        table.rows[1].cells[1].text = f'综合前端系统（{tot_workload / 8 if tot_workload % 8 else tot_workload // 8}人天）'

        for row_idx in range(table.rows.__len__()):
            cells = table.row_cells(row_idx)
            for col_idx in range(cells.__len__()):
                cell = table.cell(row_idx, col_idx)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                if col_idx == 0 or col_idx == 1:
                    cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                else:
                    cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                set_cell_font(cell, font_format={'size': 12, 'name': '仿宋'})

        merge(table, 1, 0, demand_idx, 0)
        merge(table, 1, 1, demand_idx - 1, 1)
        merge(table, demand_idx, 1, demand_idx, 2)

        log_out_dir = get_out_dir()
        self.document.save(os.path.join(log_out_dir, f'公司季报-思瑞奇-({self.__year_num}年{self.__season_num}季度).docx'))
