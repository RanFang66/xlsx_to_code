#!/usr/bin/python3
#-*- coding:UTF-8 -*-

import os
import os.path
import xlrd
import re
import sys
import getopt
import time
import importlib

importlib.reload(sys)
print (sys.getdefaultencoding)

PROJECT_NAME = "MhpsMDS"
COMPLIER = "CCS7.2"
FILE_NAME = "mbtcp_map"
PROTOCOL_EXCEL_FILE = "MDS_SC_modbus_tcp_protocol.xlsx"

HOLD_REG_MAP_START = 30000
INPUT_REG_MAP_START = 40000

ADDRESS_COL = 0
VARIABLE_EN_NAME_COL = 3
VARIABLE_NAME = 4

# holding registers map length
HOLDING_REGISTER_NUM = 128

# input registers map length
INPUT_REGISTER_NUM = 128

# the name of holding register map and input registers map
HOLDING_RE_MAP_NAME = "pHoldReAddrMap"
HOLDING_WR_MAP_NAME = "pHoldWrAddrMap"
INPUT_RE_MAP_NAME = "pInputReAddrMap"

# include files
INCLUDE_FILES = '''
#include "MDS_project"
#include "mb.h"
#include "mbtcp.h"
#include "DD250CanComm.h"
#include "DD250Ctrl.h"
'''

# holding regitsers map definition
HOLD_MAP_DEF = '''
const int16 *{hold_re_map}[{hold_map_length}];
const int16 *{hold_wr_map}[{hold_map_length}];
'''

# input register map definition
INPUT_MAP_DEF = '''
const int16 *{input_re_map}[{input_re_length}];
'''

# function name
FUNCTION_NAME = "MDSParaMapInit"

hold_reg_map = []
hold_reg_len = 0

# write file header
def write_file_header(file_handle, file_name):
    FILE_HEADER = '''
/********************************************************
*  Project:  {project_name}.prj
*  Filename: {file_name}
*  Complier: {complier}
*  Description: the register map of modbus tcp protocol
*  Created by: Ran.Fang
*  Created date: 2020.11.20
*  Updated by: {updater}
*  Updated date: {update_date}
*********************************************************
*       Copyright (c) 2020 Delta.
*       All rights reserved
*********************************************************/
'''
    file_handle.write(FILE_HEADER.format(project_name=PROJECT_NAME, file_name=FILE_NAME, \
        complier=COMPLIER, updater=os.getenv('USERNAME'), \
            update_date=time.strftime("%Y.%m.%d %H:%M:%S", (time.localtime()))))

# import register map from excel file
def import_register_map():
    # 1.open the excel file
    book = xlrd.open_workbook(PROTOCOL_EXCEL_FILE)
    table = book.sheet_by_index(1)
    hold_reg_len = table.nrows-2
    # 2.import holding register map
    for i in range(2, table.nrows):
        hold_reg_item = ["0" for j in range(3)]
        row = table.row_values(i)
        #address
        hold_reg_item[0] = str(int(row[ADDRESS_COL]) - HOLD_REG_MAP_START)
        #varible name
        hold_reg_item[1] = row[VARIABLE_EN_NAME_COL]
        #varible in program
        hold_reg_item[2] = row[VARIABLE_NAME]

        hold_reg_map.append(hold_reg_item)

def write_hold_reg_map(file_handler):
    f = file_handler

    write_file_header(f, FILE_NAME+'.c')
    f.write(INCLUDE_FILES)
    
    f.write(HOLD_MAP_DEF.format(hold_re_map=HOLDING_RE_MAP_NAME, hold_wr_map=HOLDING_WR_MAP_NAME, \
        hold_map_length=HOLDING_REGISTER_NUM))
    f.write(INPUT_MAP_DEF.format(input_re_map=INPUT_RE_MAP_NAME, input_re_length=INPUT_REGISTER_NUM))
    
    f.write("void {func_name}(void)\n".format(func_name=FUNCTION_NAME))
    f.write("{\n")

    for var in hold_reg_map:
        write_row = "\t{hold_map_name}[{address}] = (int16 *) &{variable};    \\\\{description}\n".format(hold_map_name=HOLDING_RE_MAP_NAME, \
            address=var[0], variable=var[2], description=var[1])
        f.write(write_row)
    
    f.write("}\n")


if __name__ == '__main__':
    c_file_handler = open("mbtcp_map.c", "w")
    import_register_map()
    write_hold_reg_map(c_file_handler)
