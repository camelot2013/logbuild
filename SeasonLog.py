#!/usr/bin/python
# -*- coding: utf-8 -*-


from WordStyleFunc import *
from EpibolyWorkTotalOnSystem import *
from WorkLogXls import *


class SeasonLog(object):
    @property
    def cfg(self):
        return self.__cfg

    @property
    def document(self):
        return self.__document

    @property
    def season_log(self):
        return self.__season_log

    def __init__(self, season_log: dict):
        self.__cfg = read_cfg()
        self.__document = Document()
        self.__season_log = season_log

    def get_season_work_content(self, demand_no: str):
        workload = 0
        tmp_content = []
        for log in self.__season_log['logs']:
            if demand_no == log['demand_no']:
                workload = workload + log['workload']
                if log['work_content'] not in tmp_content:
                    tmp_content.append(log['work_content'])
        return tmp_content, workload
        
    def init_document_head(self):
        # 设置中文内容的字体的方法
        head1_style = self.__document.styles.add_style('Head1', WD_STYLE_TYPE.CHARACTER)
        head1_style.font.name = '黑体'
        head1_style.font.size = Pt(16)
        head1_style.font.color.rgb = RGBColor(0, 0, 0)
        head1_style.font.bold = False
        head1_style.font.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')

        head2_style = self.__document.styles.add_style('Head2', WD_STYLE_TYPE.CHARACTER)
        head2_style.font.name = '华文中宋'
        head2_style.font.size = Pt(18)
        head2_style.font.color.rgb = RGBColor(0, 0, 0)
        head2_style.font.bold = False
        head2_style.font.element.rPr.rFonts.set(qn('w:eastAsia'), '华文中宋')

        head3_style = self.__document.styles.add_style('Head3', WD_STYLE_TYPE.CHARACTER)
        head3_style.font.name = '仿宋'
        head3_style.font.size = Pt(14)
        head3_style.font.color.rgb = RGBColor(0, 0, 0)
        head3_style.font.bold = False
        head3_style.font.element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')

        table_content_style = self.__document.styles.add_style('table_content_style', WD_STYLE_TYPE.CHARACTER)
        table_content_style.font.name = '仿宋'
        table_content_style.font.size = Pt(14)
        table_content_style.font.color.rgb = RGBColor(0, 0, 0)
        table_content_style.font.bold = False
        table_content_style.font.element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')

        p = self.__document.add_heading(level=2)
        p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        run = p.add_run('附录5', style='Head1')
        f = run.font
        f.bold = False

        p = self.__document.add_heading(level=2)
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = p.add_run('人月外包服务人员季度工作情况报告', style='Head2')
        f = run.font
        f.bold = False        


class PersonSeasonLog(SeasonLog):
    def __init__(self, person_season_log: dict, epiboly_work: EpibolyWorkTotalOnSystem):
        super(PersonSeasonLog, self).__init__(person_season_log)
        self.season_log['logs'].sort(key=lambda elem: elem['demand_no'])
        self.__person_name = self.season_log['person_name']
        self.__epiboly_work = epiboly_work
        self.__company_name = self.cfg['person_info'].get(self.__person_name).get('company')
        year_and_season = self.season_log['year_and_season']
        self.__year_num = year_and_season[:4]
        self.__season_num = year_and_season[4:]

    def init_document_head(self):
        super(PersonSeasonLog, self).init_document_head()
        p = self.document.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        text1 = "({}年{}季度)".format(self.__year_num, self.__season_num)
        p.add_run(text1, style='Head3')

        p = self.document.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        text1 = "姓名：{}        公司名称：{}".format(self.__person_name, self.__company_name)
        p.add_run(text1, style='Head3')
        pformat = p.paragraph_format
        paragraph_format = {'space_before': 0, 'space_after': 0, 'line_spacing_rule': WD_LINE_SPACING.SINGLE}
        set_paragraph_format(pformat, paragraph_format=paragraph_format)

    def create_season_log(self):
        table = self.document.add_table(rows=1, cols=4, style='Table Grid')
        table.style = 'Table Grid'
        table.width = Cm(15.77)
        table.autofit = False
        table.alignment = WD_TABLE_ALIGNMENT.CENTER  # 表格整体居中

        table.rows[0].cells[0].text = '项目名称\r/涉及系统'
        table.rows[0].cells[1].text = '目前\r进度'
        table.rows[0].cells[2].text = '工作内容及工作时间\r（需列举出每项工作的人天天数）'
        table.rows[0].cells[3].text = '科技项目\r经理确认'

        table.rows[0].height = Cm(1.03)
        table.rows[0].cells[0].width = Cm(2.89)
        table.rows[0].cells[1].width = Cm(1.57)
        table.rows[0].cells[2].width = Cm(8.86)
        table.rows[0].cells[3].width = Cm(2.45)
        tot_workload = 0
        tmp_demand_no = ''
        demand_idx = 1
        for log in self.season_log['logs']:
            tot_workload = tot_workload + log['workload']
            if tmp_demand_no != log['demand_no']:
                tmp_demand_no = log['demand_no']
                table.add_row()
                table.rows[demand_idx].cells[0].text = log['demand_name'] + '/' + log['system_name']
                table.rows[demand_idx].cells[1].text = self.__epiboly_work.demands[log['demand_no']]
                table.rows[demand_idx].cells[3].text = log['manager_name']
                log_contents, workload = self.get_season_work_content(log['demand_no'])
                for idx, log_content in enumerate(log_contents):
                    p = None
                    if idx == 0:
                        p = table.rows[demand_idx].cells[2].paragraphs[0]
                        p.style = 'List Bullet'
                    else:
                        p = table.rows[demand_idx].cells[2].add_paragraph(style='List Bullet')

                    run = p.add_run(log_content)

                    f = run.font
                    set_font_format(f, font_format={'size': 10.5, 'name': '仿宋'})
                p = table.rows[demand_idx].cells[2].add_paragraph(style='List Continue')
                run = p.add_run(f'以上工作累计花了{workload / 8 if workload % 8 else workload // 8}个工作日。')
                f = run.font
                set_font_format(f, font_format={'size': 10.5, 'name': '仿宋'})

                demand_idx = demand_idx + 1

        for row_idx in range(table.rows.__len__()):
            cells = table.row_cells(row_idx)
            for col_idx in range(cells.__len__()):
                cell = table.cell(row_idx, col_idx)
                # for cell in table.row_cells(row):
                if row_idx == 0:
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    set_cell_font(cell, font_format={'size': 14, 'name': '仿宋'})
                else:
                    if col_idx == 0:
                        set_cell_font(cell, font_format={'size': 10.5, 'name': '仿宋'})
                        table.rows[row_idx].cells[0].width = Cm(2.89)
                        cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                    elif col_idx == 1:
                        set_cell_font(cell, font_format={'size': 10.5, 'name': '仿宋'})
                        table.rows[row_idx].cells[1].width = Cm(1.57)
                        cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    elif col_idx == 2:
                        set_cell_font(cell, font_format={'size': 10.5, 'name': '仿宋'})
                        table.rows[row_idx].cells[2].width = Cm(8.86)
                        cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                    elif col_idx == 3:
                        set_cell_font(cell, font_format={'size': 10.5, 'name': '仿宋'})
                        table.rows[row_idx].cells[3].width = Cm(2.45)
                        cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        end_text = '文件命名规则为“人月外包服务人员季度工作情况报告-所属公司-姓名(xx年xx季度)'
        p = self.document.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        run = p.add_run(end_text)
        f = run.font
        set_font_format(f, font_format={'size': 14, 'name': '仿宋'})
        pformat = p.paragraph_format
        paragraph_format2 = {'space_before': 0, 'space_after': 0, 'line_spacing': Pt(20)}
        set_paragraph_format(pformat, paragraph_format=paragraph_format2)

        log_out_dir = get_out_dir()
        self.document.save(os.path.join(log_out_dir, f'附录5人月外包服务人员季度工作情况报告-成都思瑞奇信息产业有限公司-{self.__person_name}({self.__year_num}年{self.__season_num}季度).docx'))
