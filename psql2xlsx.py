#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# 13.04.2018
#--------------------------------------------------------------------------------------------------
# python3-xlsxwriter
# python3-psycopg2
#--------------------------------------------------------------------------------------------------

import os
import re
import sys
import time
import datetime
from time import sleep
#
import argparse
import configparser
import psycopg2
import xlsxwriter

_DT_FORMAT   = 'MM.DD.YYYY HH:MM:SS'
_TIME_FORMAT = 'HH:MM:SS'
_DATE_FORMAT = 'MM.DD.YYYY'


def main():
    #______________________________________________________
    # Входящие аргументы
    try:
        parser = argparse.ArgumentParser(description='psql2xls - utility for saving Postgres SQL querys results to .xlsx file')
        parser.add_argument('-f', action='store', type=str, nargs='?', default='',
                            metavar='file', help="output file name")
        args = parser.parse_args()
    except SystemExit:
        return False
    #______________________________________________________
    # Подключение config файла
    config_name = 'config.ini'
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), config_name)
    config_ini = configparser.ConfigParser()
    config_ini.read(config_path)
    if not config_ini.sections() or 'default' not in config_ini.sections():
        print("[EE] Configuration failed: 'default'", file=sys.stderr)
        return False
    if not [x for x in config_ini.sections() if x != 'default']:
        print("[..] Nothing to do")
        return False
    #______________________________________________________
    config_default = {
        'host'    : "localhost",
        'port'    : 5432,
        'user'    : "postgres",
        'password': None,
        'dbname'  : "postgres",
        'file'    : None,
        'font_name': 'Liberation Sans',
        'font_size': 10,
    }
    for x in config_default:
        ### apply default config
        try:
            config_default[x] = config_ini['default'][x]
        except KeyError:
            pass
        except Exception as err:
            print("[EE] Exception Err: {}".format(err), file=sys.stderr)
            print("[EE] Exception Inf: {}".format(sys.exc_info()), file=sys.stderr)
            return False
    #______________________________________________________
    # Проверка пути файла
    if args.f:
        config_default['file'] = args.f
    if not config_default['file']:
        print("[EE] Configuration failed: 'file'", file=sys.stderr)
        return False
    #______________________________________________________
    # Проверка директорий
    if not check_access_dir('rw', os.path.dirname(config_default['file'])):
        return False
    #______________________________________________________
    # font_name
    if not config_default['font_name']:
        print("[EE] Configuration failed: 'font_name'", file=sys.stderr)
        return False
    #______________________________________________________
    # font_size
    if not config_default['font_size']:
        print("[EE] Configuration failed: 'font_size'", file=sys.stderr)
        return False
    #______________________________________________________
    # workbook
    workbook = xlsxwriter.Workbook(config_default['file'])
    # Форматы ячеек
    workbook_format_global = workbook.formats[0]
    workbook_format_global.set_font_name(config_default['font_name'])
    workbook_format_global.set_font_size(config_default['font_size'])
    cell_format_header = workbook.add_format({'bold': True,  'font_name': config_default['font_name'], 'font_size': config_default['font_size']})
    cell_format_dt     = workbook.add_format({'bold': False, 'font_name': config_default['font_name'], 'font_size': config_default['font_size'], 'num_format': _DT_FORMAT})
    cell_format_time   = workbook.add_format({'bold': False, 'font_name': config_default['font_name'], 'font_size': config_default['font_size'], 'num_format': _TIME_FORMAT})
    cell_format_date   = workbook.add_format({'bold': False, 'font_name': config_default['font_name'], 'font_size': config_default['font_size'], 'num_format': _DATE_FORMAT})
    #==============================================================================================
    #==============================================================================================
    # Start of the work cycle
    #==============================================================================================
    for page in [x for x in config_ini.sections() if x != 'default']:
        print("[..] Page '{}'...".format(page))
        config_page = config_default.copy()
        ### apply page config
        for x in [x for x in config_ini[page] if x not in config_default.keys()]:
            try:
                config_page[x] = config_ini[page][x]
            except KeyError:
                pass
            except Exception as err:
                print("[EE] Exception Err: {}".format(err), file=sys.stderr)
                print("[EE] Exception Inf: {}".format(sys.exc_info()), file=sys.stderr)
                return False
        #print(config_page) #### TEST
        #__________________________________________________
        # page
        re_simple_str = re.compile("^([\w\- ]+)$")
        if not re_simple_str.search(page):
            print("[EE] Page name is not simple string", file=sys.stderr)
            return False
        #__________________________________________________
        # query
        if 'query' not in config_page or not config_page['query']:
            print("[EE] Configuration failed: 'query'", file=sys.stderr)
            return False
        #__________________________________________________
        # POSTGRES_DSN
        POSTGRES_DSN = """host='{}' port={} dbname='{}' user='{}'""".format(config_page['host'],
                                                                            config_page['port'],
                                                                            config_page['dbname'],
                                                                            config_page['user'],)
        POSTGRES_DSN += """ password='{}'""".format(config_page['password'] if config_page['password'] else '')
        #print(POSTGRES_DSN) #### TEST
        #__________________________________________________
        # Подключение к БД
        db = psycopg2.connect(POSTGRES_DSN)
        db.set_client_encoding('UTF8')
        cursor = db.cursor()
        ### Выполнить запрос
        cursor.execute(config_page['query'])
        #__________________________________________________
        # Добавление страниц в документ
        worksheet = workbook.add_worksheet(page)
        column_names = [{'name': desc[0], 'length': len(desc[0])} for desc in cursor.description]
        #print(column_names) #### TEST
        ### Запись шапки
        for i, x in enumerate(column_names):
            worksheet.write(0, i, x['name'], cell_format_header)
        ### Запись данных
        for i, row in enumerate(cursor.fetchall()):
            row_num = i + 1
            for coll_num, value in enumerate(row):
                #print(row_num, coll_num, value) #### TEST
                #print(type(value), value) #### TEST
                # NOTE: Специальные форматы ячеек
                if isinstance(value, datetime.datetime):
                    # WARNING: datetime.datetime перед datetime.date
                    worksheet.write(row_num, coll_num, value, cell_format_dt)
                    length = len(_DT_FORMAT)
                elif isinstance(value, datetime.time):
                    worksheet.write(row_num, coll_num, value, cell_format_time)
                    length = len(_TIME_FORMAT)
                elif isinstance(value, datetime.date):
                    worksheet.write(row_num, coll_num, value, cell_format_date)
                    length = len(_DATE_FORMAT)
                else:
                    worksheet.write(row_num, coll_num, value)
                    length = len(str(value))
                # NOTE: Считаем макс длинну столбца
                if length > column_names[coll_num]['length']:
                    column_names[coll_num]['length'] = length
        #__________________________________________________
        # sets the column width
        #print(column_names) #### TEST
        for i, x in enumerate(column_names):
            if x['length'] < 100:
                length = x['length'] + 1
            else:
                length = 100
            #print(i, length, x['name']) ####TEST
            worksheet.set_column(i, i, length)
    #______________________________________________________
    # Запись файла
    workbook.close()
    print("[OK] Workbook saved: '{}'".format(config_default['file']))
    #______________________________________________________
    return True


#==================================================================================================
# Functions
#==================================================================================================
def check_access_dir(mode, *args):
    return_value = True
    modes_dict = {'ro': os.R_OK, 'rx': os.X_OK, 'rw': os.W_OK}
    for x in args:
        if not x:
            print("[EE] Directory is not specified", file=sys.stderr)
            return_value = False
            continue
        if os.path.exists(x):
            if os.path.isdir(x):
                if not os.access(x, modes_dict[mode]):
                    print("[EE] Access denied: '{}' ({})".format(x, mode), file=sys.stderr)
                    return_value = False
            else:
                print("[EE] Is not directory: '{}'".format(x), file=sys.stderr)
                return_value = False
        else:
            print("[EE] Does not exist: '{}'".format(x), file=sys.stderr)
            return_value = False
    #______________________________________________________
    return return_value


#%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
if __name__ == '__main__':
    rc = not main() # Compatible return code
    if os.name == 'nt':
        import msvcrt
        print("[..] Press any key to exit")
        msvcrt.getch()
    sys.exit(rc)
