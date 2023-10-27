import xlwings as xw
import pandas as pd
import numpy as np
from datetime import datetime
import re 
import time

def xwAddAreaBlank():
    """
        为选定的区域的每行添加空白行
    """
    wb = xw.Book.caller()
    sht = wb.sheets.active
    app = xw.apps.active
    rng = app.selection 

    # 获取范围的具体形状，以及左上角的单元格的坐标，从而获取对于整个工作表来说，范围内第一行的行号
    r, c = rng.shape
    addition_r = rng.rows[0].row  # 范围内首行位置
    
    # 代码采用了Sheet[i, j].insert()，此时的索引变量i代表的是整个工作表中的行的索引
    # 如果采用 Range.rows[i].insert()，此时的索引变量i只代表Range范围内的行索引，和具体工作表中的行号无关
    # 而每次插入空白行后，Range范围内的行的行号都会产生变化，所以无法采用Range的方法
    
    # 另外，由于insert()方法在上方插入，因此范围内的第一行（索引对应0）不进行插入操作
    
    # 待插入的空白行的数量是范围的行数 - 1，
    for i in range(2 * r -1):
        if i % 2 == 0 and i != 0:
            # sht.range("A1").insert()
            sht[i + addition_r - 2, :].insert()      # col 值为:，否则只会对具体列单独增加空白列

def get_empty_rows(df: pd.DataFrame, axis):
    empty_rows = df[df.isnull().all(axis=axis)].index
    return empty_rows

def xwDelAreaBlank():
    """
        删掉所选区域的空白行
    """
    wb = xw.Book.caller()
    sht = wb.sheets.active
    app = xw.apps.active
    rng = app.selection 
    
    # 获取左上角位置
    left_upper_row = rng.rows[0].row
    
    data = rng.options(convert=pd.DataFrame, header=False, index=False).value
    blank_index_list = get_empty_rows(data, axis=1).to_list()
    delete_nums = 0 # 删除空行后，后续空行的行号索引随之发生变化，记录已删掉的行数，并在索引中减掉
    for idx in blank_index_list:
        row_num = idx + left_upper_row - delete_nums - 1
        sht[row_num, :].delete()
        delete_nums += 1

def xwConvStringToNumbers():
    wb = xw.Book.caller()
    sht = wb.sheets.active
    app = xw.apps.active
    rng = app.selection 

    data = rng.options(convert=pd.DataFrame, index=False, header=False).value 
    try:
        new_data = data.astype(float)
    except Exception as e:
        app.alert(repr(e))
    else:
        rng.options(convert=pd.DataFrame, index=False, header=False).value = new_data
    
