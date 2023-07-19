import re
import pandas as pd
import openpyxl
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

def read_diff(extracted, control, ref, designation = True, extension = True):
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
                diff_content[(row, col)] = (extracted.loc[row, col], control.loc[row, col])
    return diff, diff_content
