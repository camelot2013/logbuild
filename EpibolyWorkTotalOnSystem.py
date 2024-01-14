#!/usr/bin/python
# -*- coding: utf-8 -*-


import xlrd
import os
import sys
import re


class EpibolyWorkTotalOnSystem(object):
    @property
    def demands(self):
        return self.__demands

    @property
    def season_work(self):
        return self.__season_work

    @property
    def year_num(self):
        return self.__year_num

    @property
    def season_num(self):
        return self.__season_num

    @property
    def corp_name(self):
        return self.__corp_name

    def __init__(self):
        self.__xls_table = None
        self.__demands = dict()
        self.__season_work = []
        self.__season_num = ''
        self.__year_num = ''
        self.__corp_name = ''

    def read_xls(self, file_path, __person_name: str):
        if not os.path.exists(file_path):
            return
        data = xlrd.open_workbook(file_path)
        try:
            self.__xls_table = data.sheet_by_index(0)
        except xlrd.biffh.XLRDError:
            sys.exit(0)
        table_row = self.__xls_table.row_values(0)
        self.__year_num = re.findall(r"\d*", table_row[0].strip())[12]
        self.__season_num = re.findall(r"\d*", table_row[0].strip())[15]
        __last_system_name = ''
        __last_demand_no = ''
        __last_demand_name = ''
        __last_month_request_no = ''

        def __get_value_by_default(v1: str, __default_value: str):
            return v1 if v1 else __default_value

        for row_idx in range(3, self.__xls_table.nrows):
            table_row = self.__xls_table.row_values(row_idx)

            system_name = __get_value_by_default(table_row[0].strip(), __last_system_name)
            person_name = table_row[9].strip()
            person_level = table_row[11].strip()
            demand_no = __get_value_by_default(table_row[1].strip(), __last_demand_no)
            demand_name = __get_value_by_default(table_row[2].strip(), __last_system_name)
            month_request_no = __get_value_by_default(table_row[5].strip(), __last_month_request_no)
            workload = float(table_row[12].strip())
            corp_name = table_row[10].strip()
            self.__corp_name = corp_name
            if __person_name == person_name:
                self.__demands[table_row[1].strip()] = table_row[8].strip()

            __last_system_name = system_name
            __last_demand_no = demand_no
            __last_demand_name = demand_name
            __last_month_request_no = month_request_no
            person_season = self.get_person_season(person_name)
            if not person_season:
                person_season = {'person_name': person_name, 'person_level': person_level, 'system': [{'system_name': system_name, 'works': [{'demand_no': demand_no, 'demand_name': demand_name, 'demand': f'{demand_no}{demand_name}', 'month_request_no': month_request_no, 'workload': workload}]}]}
                self.__season_work.append(person_season)
                continue
            system_works = self.get_works_at_person_by_system_name(person_season, system_name)
            if not system_works:
                person_season.get('system').append({'system_name': system_name, 'works': [{'demand_no': demand_no, 'demand_name': demand_name, 'demand': f'{demand_no}{demand_name}', 'month_request_no': month_request_no, 'workload': workload}]})
                continue
            if system_works:
                system_works.append({'demand_no': demand_no, 'demand_name': demand_name, 'demand': f'{demand_no}{demand_name}', 'month_request_no': month_request_no, 'workload': workload})

    def get_person_season(self, person_name: str):
        for person_season in self.__season_work:
            if person_name == person_season.get('person_name'):
                return person_season

    @staticmethod
    def get_works_at_person_by_system_name(__person_season: dict, __system_name: str):
        for system_log in __person_season.get('system'):
            if __system_name == system_log.get('system_name'):
                return system_log.get('works')


if __name__ == '__main__':
    from CorpSeasonLog import CorpSeasonLog

    a1 = EpibolyWorkTotalOnSystem()
    a1.read_xls('D:/tmp/2023Q4/外包人月工作量统计报表-2023年第4季度-成都思瑞奇信息产业有限公司.xls', 'XXX')
    corp_season_log = CorpSeasonLog(a1.year_num, a1.season_num, a1.corp_name)
    corp_season_log.create_corp_season_log(a1.season_work)
