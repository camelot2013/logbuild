#!/usr/bin/python
# -*- coding: utf-8 -*-


import xlrd
import os
import datetime
import json


def read_cfg():
    file_name = 'person_list.json'
    # current_path = os.path.split(os.path.realpath(__file__))[0]
    current_path = os.getcwd()
    file = os.path.join(current_path, file_name)
    with open(file, encoding='UTF-8') as fp:
        content = fp.read()
        return json.loads(content)


def get_out_dir():
    person_cfg = read_cfg()
    log_out_dir = person_cfg.get('log_out_dir')
    if not log_out_dir:
        log_out_dir = 'D:/'
    if not log_out_dir.endswith('/') and not log_out_dir.endswith('\\'):
        log_out_dir = log_out_dir + '/'

    return log_out_dir


class WorkLogXls(object):

    @property
    def work_logs(self):
        return self.__work_logs

    @property
    def weeks(self):
        return self.__weeks

    @property
    def months(self):
        return self.__months

    @property
    def seasons(self):
        return self.__seasons

    def __init__(self):
        self.__xls_table = None
        self.__work_logs = []
        self.__weeks = []
        self.__months = []
        self.__seasons = []

    @staticmethod
    def __sort_key_by_date(elem):
        return elem['work_date']

    @staticmethod
    def __sort_key_by_demand(elem):
        return elem['demand_name']

    @staticmethod
    def __sort_key_by_date_and_demand(elem):
        return elem['work_date'] + elem['demand_no']

    @staticmethod
    def __sort_key_by_person_and_date(elem):
        return elem['person_name'] + elem['work_date']

    @staticmethod
    def __sort_key_by_person_and_date_and_demand(elem):
        return elem['person_name'] + elem['work_date'] + elem['demand_no']

    def work_log_sort_by_date(self):
        self.__work_logs.sort(key=self.__sort_key_by_date)

    def work_log_sort_by_person_and_date(self):
        self.__work_logs.sort(key=self.__sort_key_by_person_and_date)

    def work_log_sort_by_demand(self):
        self.__work_logs.sort(key=self.__sort_key_by_demand)

    def work_log_sort_by_date_and_demand(self):
        self.__work_logs.sort(key=self.__sort_key_by_date_and_demand)

    def work_log_sort_by_person_and_date_and_demand(self):
        self.__work_logs.sort(key=self.__sort_key_by_person_and_date_and_demand)

    def read_xls(self, file_path):
        if os.path.exists(file_path):
            data = xlrd.open_workbook(file_path)
            try:
                self.__xls_table = data.sheet_by_index(0)
            except xlrd.biffh.XLRDError:
                exit(0)
            for row_idx in range(1, self.__xls_table.nrows):
                table_row = self.__xls_table.row_values(row_idx)

                person_name = table_row[0].strip()
                work_date = table_row[1].strip()
                demand_name = table_row[5].strip()
                demand_no = demand_name[:15]
                manager_name = table_row[6].strip()
                month_request_no = table_row[7].strip()
                workload = int(table_row[9].strip())
                work_content = table_row[10].strip()
                work_time = datetime.datetime.strptime(work_date, '%Y-%m-%d').date()
                if workload > 0:
                    work_log = {'work_date': work_time.strftime('%Y%m%d'),
                                'week': work_time.weekday(),
                                'year_week': work_time.strftime('%Y%W'),
                                'person_name': person_name,
                                'demand_name': demand_name,
                                'demand_no': demand_no,
                                'manager_name': manager_name,
                                'month_request_no': month_request_no,
                                'workload': workload,
                                'system_name': table_row[4].strip(),
                                'work_content': work_content}
                    self.__work_logs.append(work_log)

            if self.__work_logs:
                self.work_log_sort_by_date()

    def split_weeks(self):
        self.work_log_sort_by_date()
        tmp_week = []
        year_week = ''
        self.__weeks.clear()
        for work_log in self.work_logs:
            if work_log['year_week'] != year_week:
                tmp_week = list()
                self.__weeks.append(tmp_week)
            tmp_week.append(work_log)
            year_week = work_log['year_week']

    def split_months(self):
        self.work_log_sort_by_date()
        self.__months.clear()
        work_month = ''
        month_log = {}
        for work_log in self.work_logs:
            if work_month != work_log['work_date'][:6]:
                work_month = work_log['work_date'][:6]
                month_log = dict()
                month_log['month'] = work_month
                month_log['person_name'] = work_log['person_name']
                month_log['logs'] = [work_log, ]
                self.months.append(month_log)
            else:
                month_log['logs'].append(work_log)

    def split_season(self):
        self.work_log_sort_by_date()
        self.__seasons.clear()
        tmp_season = ''
        season_log = {}
        for date_log in self.work_logs:
            work_time = datetime.datetime.strptime(date_log['work_date'], '%Y%m%d').date()
            year_and_season = '{}{}'.format(date_log['work_date'][:4], (work_time.month - 1) // 3 + 1)
            if tmp_season != year_and_season:
                season_log = dict()
                tmp_season = year_and_season
                season_log['year_and_season'] = year_and_season
                season_log['person_name'] = date_log['person_name']
                season_log['logs'] = [date_log, ]
                self.seasons.append(season_log)
            else:
                season_log['logs'].append(date_log)
