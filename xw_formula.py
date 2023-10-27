import xlwings as xw
import pandas as pd
import numpy as np
from datetime import datetime
import re 
import time

@xw.func
@xw.arg("area", convert=pd.DataFrame, header=False, index=False, doc="计算Describe的范围")
@xw.arg("header", doc="DataFrame的行头")
@xw.arg("index", doc="DataFrame的索引")
def xwUDFDataDesc(area:pd.DataFrame, header=False, index=False):
    """计算所选区域的DataFrame.Describe()

    Args:
        area (_type_): 选择的区域
        header (_type_): 行头
        index (_type_): 索引
    """
    headers = []
    indexes = []

    if header:
        if isinstance(header, list):
            headers.extend(header)
        else:
            headers.append(header)
        assert len(headers) == area.shape[1], "传入的header参数的长度与area的列数不等" 
        area.columns = headers
    if index:
        if isinstance(index, list):
            indexes.extend(index)
        else:
            indexes.append(index)
        assert len(indexes) == area.shape[0], "传入的index参数的长度与area的行数数不等"
        area.index = indexes
    
    return area.describe()


@xw.func
def xwUDFDataGrpByCols():
    pass  

@xw.func
def xwUDFDataGrpByFigs():
    pass


@xw.func
@xw.arg("area", convert=pd.DataFrame, header=True, index=False, doc="分组计算的范围，必须包括行头")
@xw.arg("cols_num", doc="进行group by的分组列数")
@xw.arg("method", doc="汇总方法的字符串，'mean', 'sum', 'size', 'count', 'std', 'var', 'describe', 'first', 'last', 'min', 'max'")
@xw.ret(index=False)
def xwUDFDataGrpBy(area:pd.DataFrame, cols_num:int, method, *header):
    support_num_only_methods = ['mean', 'sum', 'std', 'var', 'first', 'last', 'min', 'max']
    not_support_num_only_methods = ['size', 'count']
    if method not in support_num_only_methods and method not in not_support_num_only_methods:
        return "方法不支持"
    # 将header参数拼接为一个列表：
    # 如果是点击一个单元格，传进来的是一个普通对象，如果是连续选择多个单元格，传进来是个列表
    # 因此，要确保所有header元素都进入一个列表
    headers = []
    for hd in header:
        if isinstance(hd, list):
            headers.extend(hd)
        else:
            headers.append(hd)

    # 判断cols_num + figs_num是否与header的长度相同
    if cols_num > len(headers):
        return "cols_num输入错误，不允许大于area参数的列数"
    
    # 判断传入的header是否是area的列名的子集
    full_cols = set(area.columns)
    sub_cols = set(headers)
    if not sub_cols.issubset(full_cols):
        # xw.apps.active.alert("所选列名不在范围内")
        return "所选列名不在范围内"
    grp_cols = headers[:int(cols_num)]
    
    if method in support_num_only_methods:
        return area[headers].groupby(grp_cols).agg(method,numeric_only=True).reset_index()
    else:
        return area[headers].groupby(grp_cols).agg(method).reset_index()


