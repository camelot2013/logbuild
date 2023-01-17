#!/usr/bin/python
# -*- coding: utf-8 -*-


import xlrd
import os
import sys


class EpibolyWorkTotalOnSystem(object):
    @property
    def demands(self):
        return self.__demands

    def __init__(self):
        self.__xls_table = None
        self.__demands = dict()

    def read_xls(self, file_path):
        if os.path.exists(file_path):
            data = xlrd.open_workbook(file_path)
            try:
                self.__xls_table = data.sheet_by_index(0)
            except xlrd.biffh.XLRDError:
                sys.exit(0)
            for row_idx in range(3, self.__xls_table.nrows):
                table_row = self.__xls_table.row_values(row_idx)
                self.__demands[table_row[1].strip()] = table_row[8].strip()


if __name__ == '__main__':
    a1 = EpibolyWorkTotalOnSystem()
    a1.read_xls('D:/2222.xls')
    print(a1.demands)
