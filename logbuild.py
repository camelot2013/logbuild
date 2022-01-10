#!/usr/bin/python
# -*- coding: utf-8 -*-


from WorkLogXls import *
from WeekLog import create_weeklog
from MonthLog import creat_month_log
from SeasonLog import create_season_log
from CorpSeasonLog import create_corp_season_log
import traceback

if __name__ == '__main__':
    try:
        work_log_xls = WorkLogXls()
        xls_file_path = input('请输入工作日志信息导出文件（全路径文件名）：')
        if not os.path.exists(xls_file_path):
            print('工作日志信息导出文件不存在，请确认!')
            exit(0)

        work_log_xls.read_xls(xls_file_path)
        work_log_xls.split_weeks()
        for week in work_log_xls.weeks:
            create_weeklog(week)

        work_log_xls.split_months()
        for month_log in work_log_xls.months:
            creat_month_log(month_log)

        work_log_xls.split_season()
        for season in work_log_xls.seasons:
            create_season_log(season)
            create_corp_season_log(season)
    except Exception as er:
        traceback.print_exc()
        input("回车键退出！")

