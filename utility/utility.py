import re
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, colors
def find_graph_id(coordinates, rectangles):
    for i in range(len(rectangles)):
        x1, x2, y1, y2 = rectangles[i][0]
        graph_id = rectangles[i][1]
        x, y = coordinates[0], coordinates[1]
        if x1 <= x <= x2 and y1 <= y <= y2:
            return graph_id, rectangles[i]
    return None, None

def parameter_retriving(entity):
    """entity is a dxf entity, parameters is a list of parameters to be retrived"""
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
    for entity in entity_list:
        if re.search(string, entity.text) is not None:
            entity_list.remove(entity)
    return entity_list

def enum_to_dict(enum):
        return {v:k for k,v in enum}

def read_sheet(filepath):
    sensors_control = pd.read_excel(filepath, sheet_name='SENSOR_IO')
    specials_control = pd.read_excel(filepath, sheet_name='SPECIAL_IO')
    actuators_control = pd.read_excel(filepath, sheet_name='ACTUATOR_IO')
    for _ in (sensors_control, specials_control, actuators_control):
        _.rename(mapper= lambda x: ''.join(''.join(''.join(x.split('_x000D_\n')).split('\n')).split('-')), axis=1, inplace=True)
    ref = enum_to_dict(enumerate(sensors_control.columns.tolist(),1))
    return (sensors_control, specials_control, actuators_control), ref

def unify(df):
    df.sort_values(by=['信号位号idcode', '图号diagram number','扩展码extensioncode'], inplace=True, ignore_index=True)
    return df

def read_diff(extracted, control, ref, extension = True, designation = True):
    """marks diff cells in both dataframes as red"""
    diff = []
    diff_content = {}
    labels = ['扩展码extensioncode','信号位号idcode', '图号diagram number', '版本rev.', '信号说明designation']
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
                if (control.loc[row,'序号serial number']+1) not in diff_content.keys():
                    diff_content[(control.loc[row,'序号serial number']+1)] = {}
                diff_content[(control.loc[row,'序号serial number']+1)][col] = extracted.loc[row, col].replace('\n', '')
    return diff, diff_content

def output_diff(filepath, ref, comparing_data):
    sensors, specials, actuators, sensors_control, specials_control, actuators_control = comparing_data
    wb = openpyxl.load_workbook(filepath)
    sensors_writer = wb['SENSOR_IO']
    specials_writer = wb['SPECIAL_IO']
    actuators_writer = wb['ACTUATOR_IO']
    def write_diff(writer, extracted, control, extension = True, designation = True ):
        pos, content = read_diff(extracted, control, ref, extension, designation)
        insert_col = writer.max_column + 1
        writer.insert_cols(insert_col, 1)
        writer.cell(1, insert_col).value = 'CAD Content'
        for row, col in pos:
            writer.cell(row, col).font = Font(color='FF0000')
        for row in content.keys():
            writer.cell(row, insert_col).value = str(content[row])
        return
    write_diff(sensors_writer, sensors, sensors_control, extension = False)
    write_diff(specials_writer, specials, specials_control, extension = False)
    write_diff(actuators_writer, actuators, actuators_control)
    wb.save(filepath)