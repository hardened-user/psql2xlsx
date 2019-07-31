#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# 31.07.2019
# ----------------------------------------------------------------------------------------------------------------------
import argparse
import configparser
import datetime
import os
import re
import sys
import traceback

import psycopg2
import xlsxwriter

_DT_FORMAT = 'MM.DD.YYYY HH:MM:SS'
_TIME_FORMAT = 'HH:MM:SS'
_DATE_FORMAT = 'MM.DD.YYYY'


def main():
    # __________________________________________________________________________
    # command-line options, arguments
    try:
        parser = argparse.ArgumentParser(
            description="psql2xls - utility for saving Postgres SQL queries results to .xlsx file")
        parser.add_argument('-c', '--config', action='store', default="config.ini",
                            metavar="<CONFIG_FILE>", help="configuration file path")
        parser.add_argument('-o', '--output', action='store', default=None,
                            metavar="<OUTPUT_FILE>", help="output file path")
        parser.add_argument('-f', '--overwrite', action='store_true', default=False,
                            help="allow overwrite output file")
        args = parser.parse_args()
    except SystemExit:
        return False
    # __________________________________________________________________________
    # read configuration file
    try:
        config_ini = configparser.ConfigParser()
        config_ini.read(os.path.abspath(args.config))
    except Exception as err:
        print("[!!] Exception :: {}\n{}".format(err, "".join(traceback.format_exc())), flush=True)
        return False
    # __________________________________________________________________________
    if not config_ini.sections() or 'default' not in config_ini.sections():
        print("[EE] Missing configuration section :: default", flush=True)
        return False
    if not [x for x in config_ini.sections() if x != 'default']:
        print("[..] Nothing to do", flush=True)
        return False
    # __________________________________________________________________________
    # generate default config
    config_default = {
        'output': None,
        'overwrite': False,
        'font_name': 'Liberation Sans',
        'font_size': 10,
        'host': "localhost",
        'port': 5432,
        'base': "postgres",
        'user': "postgres",
        'pass': None,
        'query': None
    }
    for x in config_default:
        try:
            config_default[x] = config_ini['default'][x]
        except KeyError:
            pass
        except Exception as err:
            print("[!!] Exception :: {}\n{}".format(err, "".join(traceback.format_exc())), flush=True)
            return False
    # inspect: output
    if args.output:
        config_default['output'] = args.output
    if not config_default['output']:
        print("[EE] Invalid option value :: output", flush=True)
        return False
    # inspect: overwrite
    if args.overwrite:
        config_default['overwrite'] = args.overwrite
    if isinstance(config_default['overwrite'], str):
        if config_default['overwrite'].lower() in ('true', 'yes', 'on'):
            config_default['overwrite'] = True
        else:
            config_default['overwrite'] = False
    # inspect: font_name
    if not config_default['font_name']:
        print("[EE] Invalid option value :: font_name", flush=True)
        return False
    # inspect: font_size
    if not config_default['font_size']:
        print("[EE] Invalid option value :: font_size", flush=True)
        return False
    # __________________________________________________________________________
    # check permission
    if not fs_check_access_file(config_default['output'], config_default['overwrite']):
        return False
    # __________________________________________________________________________
    # make workbook
    workbook = xlsxwriter.Workbook(config_default['output'])
    workbook_format_global = workbook.formats[0]
    workbook_format_global.set_font_name(config_default['font_name'])
    workbook_format_global.set_font_size(config_default['font_size'])
    # special cell formats
    cell_format_header = workbook.add_format(
        {'bold': True, 'font_name': config_default['font_name'], 'font_size': config_default['font_size']})
    cell_format_dt = workbook.add_format(
        {'bold': False, 'font_name': config_default['font_name'], 'font_size': config_default['font_size'],
         'num_format': _DT_FORMAT})
    cell_format_time = workbook.add_format(
        {'bold': False, 'font_name': config_default['font_name'], 'font_size': config_default['font_size'],
         'num_format': _TIME_FORMAT})
    cell_format_date = workbook.add_format(
        {'bold': False, 'font_name': config_default['font_name'], 'font_size': config_default['font_size'],
         'num_format': _DATE_FORMAT})
    # ==================================================================================================================
    # ==================================================================================================================
    # Start of the work cycle
    # ==================================================================================================================
    for page in [x for x in config_ini.sections() if x != 'default']:
        print("[..] Generated page :: {} ...".format(page))
        re_simple_str = re.compile(r"^([\w\- ]+)$")
        if not re_simple_str.search(page):
            print("[EE] Page name is not simple string", flush=True)
            return False
        # __________________________________________________________________________
        # generate page config
        config_page = config_default.copy()
        for x in config_page:
            try:
                config_page[x] = config_ini[page][x]
            except KeyError:
                pass
            except Exception as err:
                print("[!!] Exception :: {}\n{}".format(err, "".join(traceback.format_exc())), flush=True)
                return False
        # print(config_page)  # TEST
        # ______________________________________________________________________
        # inspect: query
        if 'query' not in config_page or not config_page['query']:
            print("[EE] Invalid option value :: query", flush=True)
            return False
        # ______________________________________________________________________
        # database connection
        postgres_dsn = "{}{}{}{}{}".format(
            "host='{}' ".format(config_page['host']) if config_page['host'] else '',
            "port='{}' ".format(config_page['port']) if config_page['port'] else '',
            "user='{}' ".format(config_page['user']) if config_page['user'] else '',
            "dbname='{}' ".format(config_page['base']) if config_page['base'] else '',
            "password='{}'".format(config_page['pass']) if config_page['pass'] else '',
        ).strip()
        # print(postgres_dsn)  # TEST
        db = pg_connect(postgres_dsn)
        if not db:
            return False
        # ______________________________________________________________________
        # make worksheet
        worksheet = workbook.add_worksheet(page)
        row_num = 0
        column_width = {}
        # ______________________________________________________________________
        # execute queries, supported multiple queries
        for query in config_page['query'].split(";\n"):
            cursor = pg_query(db, query.strip())
            if not cursor:
                return False
            for col_num, column in enumerate(cursor.description):
                # print(column)  # <class 'psycopg2.extensions.Column'>
                # print("HEAD >", row_num, col_num, column.name)  # TEST
                if column.name != "?column?":
                    worksheet.write(row_num, col_num, column.name, cell_format_header)
                    # NOTE: calculate the maximum length of a row in a column
                    length = len(column.name)
                    column_width.setdefault(col_num, 0)
                    if length > column_width[col_num]:
                        column_width[col_num] = length
            row_num += 1
            for column in cursor.fetchall():
                for col_num, value in enumerate(column):
                    # print("DATA >", row_num, col_num, value)  # TEST
                    # print(type(value), value) #### TEST
                    # special cell formats
                    if isinstance(value, datetime.datetime):
                        # WARNING: check <'datetime.datetime'> before <'datetime.date'>
                        worksheet.write(row_num, col_num, value, cell_format_dt)
                        length = len(_DT_FORMAT)
                    elif isinstance(value, datetime.time):
                        worksheet.write(row_num, col_num, value, cell_format_time)
                        length = len(_TIME_FORMAT)
                    elif isinstance(value, datetime.date):
                        worksheet.write(row_num, col_num, value, cell_format_date)
                        length = len(_DATE_FORMAT)
                    else:
                        worksheet.write(row_num, col_num, value)
                        length = len(str(value))
                    # NOTE: calculate the maximum length of a row in a column
                    column_width.setdefault(col_num, 0)
                    if length > column_width[col_num]:
                        column_width[col_num] = length
                # ______________________________________________________________
                row_num += 1
        # ______________________________________________________________________
        # set the width of the column
        # NOTE: limit the maximum width
        # print(column_width)  # TEST
        for x in column_width:
            length = column_width[x]
            if length < 80:
                length += 1
            else:
                length = 80
            # print(x, length)  # TEST
            worksheet.set_column(x, x, length)
    # __________________________________________________________________________
    # write file
    workbook.close()
    print("[OK] Workbook saved :: {}".format(config_default['output']), flush=True)
    # __________________________________________________________________________
    return True


# ======================================================================================================================
# Functions
# ======================================================================================================================
def fs_check_access_file(path, ignore_existing=False):
    """
    File permission check.
    """
    path = os.path.abspath(path)
    if os.path.exists(path):
        if os.path.isdir(path):
            print("[EE] Is a directory :: {}".format(path), flush=True)
            return False
        if not ignore_existing:
            print("[EE] File already exists :: {}".format(path), flush=True)
            return False
        if not os.access(path, os.W_OK):
            print("[EE] File access denied :: {}".format(path), flush=True)
            return False
    else:
        if not os.access(os.path.dirname(path), os.W_OK):
            print("[EE] Directory access denied :: {}".format(path), flush=True)
            return False
    # __________________________________________________________________________
    return True


def pg_connect(dsn):
    """
    Create a new database session and return a new connection object.
    """
    try:
        conn = psycopg2.connect(dsn)
        conn.set_client_encoding('UTF8')
        print("[OK] PostgreSQL successfully connected")
    except (psycopg2.OperationalError, psycopg2.ProgrammingError) as err:
        print("[EE] PostgreSQL Exception :: {}".format(err, flush=True))
        return None
    except Exception as err:
        print("[!!] Exception :: {}\n{}".format(err, "".join(traceback.format_exc())), flush=True)
        return None
    # __________________________________________________________________________
    return conn  # <class 'psycopg2.extensions.connection'>


def pg_query(conn, query):
    """
    Execute a database operation (query or command).
    """
    cursor = conn.cursor()
    try:
        cursor.execute(query)
    except (psycopg2.DataError, psycopg2.ProgrammingError) as err:
        print("[EE] PostgreSQL Exception :: {}".format(err), flush=True)
        conn.rollback()
        cursor.close()
        return None
    except Exception as err:
        print("[!!] Exception :: {}\n{}".format(err, "".join(traceback.format_exc())), flush=True)
        return None
    else:
        conn.commit()
    # __________________________________________________________________________
    return cursor  # <class 'psycopg2.extensions.cursor'>


# %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
if __name__ == '__main__':
    rc = main()
    # __________________________________________________________________________
    if os.name == 'nt':
        # noinspection PyUnresolvedReferences
        import msvcrt

        print("[..] Press any key to exit", flush=True)
        msvcrt.getch()
    # __________________________________________________________________________
    sys.exit(not rc)  # Compatible return code
