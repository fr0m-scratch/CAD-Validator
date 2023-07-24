import re
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Color, Alignment, Border, Side, PatternFill
from rich.progress import Progress
import os
import sys
def find_graph_id(coordinates, rectangles):
    #coordinates is a tuple of x and y
    #rectangles 是一个列表，每个元素是一个元组，元组的第一个元素是一个四元组，分别是左下角和右上角的坐标，第二个元素是图形的id
    for i in range(len(rectangles)):
        x1, x2, y1, y2 = rectangles[i][0]
        graph_id = rectangles[i][1].tag
        x, y = coordinates[0], coordinates[1]
        if x1 <= x <= x2 and y1 <= y <= y2:
            return graph_id, rectangles[i]
    return None, None

def parameter_retriving(entity):
    """entity is a dxf entity, parameters is a list of parameters to be retrived"""
    #entity是一个dxf实体，parameters是一个列表，每个元素是一个参数名，返回一个字典，key是参数名，value是参数值
    relevant_attributes = ['ObjectName', 'InsertionPoint', 'coordinates', 'Area', 'Length', 'TextString', 'Handle', 'height', 'Layer', 'TagString', 'color']
    relevant_member_function = ['Coordinate', 'GetBoundingBox']
    res_dict = {}
    for function in relevant_member_function:
        try: 
            res = getattr(entity, function)()
            res_dict[function] = res
        except:
            res_dict[function] = None
    for attribute in relevant_attributes:
        try:
            attr = getattr(entity, attribute)
            res_dict[attribute] = attr
        except:
            res_dict[attribute] = None
    if res_dict['coordinates'] is not None:
        coordinates = []
        for i in range(len(res_dict['coordinates'])//2):
            coordinates.append(entity.Coordinate(i))
        res_dict['coordinates'] = coordinates
    return res_dict

def remove_entity_by_string(entity_list, string):
    #去除entity_list中text中包含string的entity
    for entity in entity_list:
        if re.search(string, entity.text) is not None:
            entity_list.remove(entity)
    return entity_list

def enum_to_dict(enum):
    #把enum转换成字典，key是序号，value是值
    return {v:k for k,v in enum}

def read_sheet(filepath):
    #读取excel表格，返回一个元组，元组的每个元素是一个dataframe，分别是SENSOR_IO, SPECIAL_IO, ACTUATOR_IO
    with Progress() as progress:
        read = progress.add_task("[blue]Read[cyan2]file", total=14)
        sensors_control = pd.read_excel(filepath, sheet_name='SENSOR_IO')
        progress.update(read, advance=4)
        specials_control = pd.read_excel(filepath, sheet_name='SPECIAL_IO')
        progress.update(read, advance=5)
        actuators_control = pd.read_excel(filepath, sheet_name='ACTUATOR_IO')
        progress.update(read, advance=1)
        for _ in (sensors_control, specials_control, actuators_control):
            _.rename(mapper= lambda x: ''.join(''.join(''.join(x.split('_x000D_\n')).split('\n')).split('-')), axis=1, inplace=True)
            progress.update(read, advance=1)
        ref = enum_to_dict(enumerate(sensors_control.columns.tolist(),1))
        progress.update(read, advance=1)
    return (sensors_control, specials_control, actuators_control), ref

def unify(df):
    #统一dataframe中的数据顺序以进行比对
    df.sort_values(by=['信号位号idcode', '图号diagram number','扩展码extensioncode'], inplace=True, ignore_index=True)
    return df

def read_diff(extracted, control, ref, extension = True, designation = True):
    #比对两个dataframe，返回一个元组，元组的第一个元素是一个列表，每个元素是一个元组，元组的第一个元素是序号，第二个元素是列名
    diff = []
    diff_content = {}
    labels = ['扩展码extensioncode','信号位号idcode', '图号diagram number', '版本rev.', '信号说明designation', "安全分级/分组Safetyclass/division"]
    if not designation:
        labels.remove('信号说明designation')
    if not extension:
        labels.remove('扩展码extensioncode')
    for row in range(extracted.shape[0]):
        for col in labels:
            if col == '信号位号idcode':
                diff_check = extracted.loc[row, col][1:] != control.loc[row, col][1:]
            else:
                diff_check = extracted.loc[row, col] != control.loc[row, col]
            if diff_check:
                if col == '信号说明designation':
                    if control.loc[row, col] == extracted.loc[row, col].replace('\n', ''):
                        continue
                diff.append((control.loc[row,'序号serial number']+1, ref[col]))
                if col not in diff_content:
                    diff_content[col] = {}
                diff_content[col][(control.loc[row,'序号serial number']+1)] = extracted.loc[row, col].replace('\n', '') or '无描述'
    return diff, diff_content

def output_diff(filepath, ref, comparing_data):
    #把比对结果写入excel表格
    sensors, specials, actuators, sensors_control, specials_control, actuators_control = comparing_data
    wb = openpyxl.load_workbook(filepath)
    sensors_writer = wb['SENSOR_IO']
    specials_writer = wb['SPECIAL_IO']
    actuators_writer = wb['ACTUATOR_IO']
    col_dict = {
        "信号位号idcode": (1, "CAD信号位号"),
        "信号说明designation": (2, "CAD信号说明"),
        "安全分级/分组Safetyclass/division": (3, "CAD安全分级/分组"),
        "图号diagram number": (4, "CAD图号"),
        "版本rev.": (5, "CAD版本"),
    }
    
    def write_diff(writer, extracted, control, extension = True, designation = True ):
        pos, content = read_diff(extracted, control, ref, extension, designation)
        insert_col = writer.max_column + 1
        writer.insert_cols(insert_col, 1)
        for row, col in pos:
            writer.cell(row, col).font = Font(color='FF0000')
        for col in content.keys():
            for row in content[col].keys():
                writer.cell(row, insert_col+col_dict[col][0]).value = str(content[col][row])
        for col in col_dict.values():
            writer.cell(1, insert_col+col[0]).value = col[1]
            writer.cell(1, insert_col+col[0]).font = Font(bold=True, size=10)
            writer.cell(1, insert_col+col[0]).alignment = Alignment(horizontal='center', vertical='center')
        for i in range(insert_col, insert_col+5):
            writer.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 15
        return
    write_diff(sensors_writer, sensors, sensors_control, extension = False)
    write_diff(specials_writer, specials, specials_control, extension = False)
    write_diff(actuators_writer, actuators, actuators_control)
    wb.save(filepath)

def get_resource_path(relative_path):
    """解析编译后的数据文件路径"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def contain(words,entities):
        #根据关键词查找实体
        res = []
        for entity in entities:
            if entity.text is not None:
                text = re.search(words, entity.text)
                if text is not None:
                    res.append(entity)
            if entity.tag is not None:
                tag = re.search(words, entity.tag)
                if tag is not None:
                    res.append(entity)
        return res
    
def not_contain(words, entities):
        #根据关键词查找实体
        res = []
        for entity in entities:
            if entity.text is not None:
                text = re.search(words, entity.text)
                if text is None:
                    res.append(entity)
            if entity.tag is not None:
                tag = re.search(words, entity.tag)
                if tag is None:
                    res.append(entity)
        return set(res)
    
def unformat_mtext(s, exclude_list=('P', 'S')):
    """Returns string with removed format information

    :param s: string with multitext
    :param exclude_list: don't touch tags from this list. Default ('P', 'S') for
                         newline and fractions

    ::

        >>> text = ur'{\\fGOST type A|b0|i0|c204|p34;TEST\\fGOST type A|b0|i0|c0|p34;123}'
        >>> unformat_mtext(text)
        u'TEST123'

    """
    s = re.sub(r'\{?\\[^%s][^;]+;' % ''.join(exclude_list), '', s)
    s = re.sub(r'\}', '', s)
    return s



def mtext_to_string(s):
    """
    Returns string with removed format innformation as :func:`unformat_mtext` and
    `\\P` (paragraphs) replaced with newlines

    ::

        >>> text = ur'{\\fGOST type A|b0|i0|c204|p34;TEST\\fGOST type A|b0|i0|c0|p34;123}\\Ptest321'
        >>> mtext_to_string(text)
        u'TEST123\\ntest321'

    """

    return unformat_mtext(s).replace(u'\\P', u'\n')