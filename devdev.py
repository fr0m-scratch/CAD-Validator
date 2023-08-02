from utility.utility import *
from cad_entity import CADEntity
from cad_handler import CADHandler
import shutil
from rich.progress import Progress
from rich.traceback import install
import sys
import os
install()

def main(graphfilepath, filepath):
    if re.search(r'\/', graphfilepath) is None:
        graphfilepath = get_resource_path(graphfilepath)
    acad = CADHandler(graphfilepath)

    control, ref = read_sheet(filepath)
    
    sensors, specials, actuators = acad.accessories.values()
    extracted = acad.accessories_to_map(sensors, specials, actuators)
    
  
    sensors, specials, actuators = extracted
    for _ in control:
        _['refindex'] = _.index
        print(_['refindex'])
    sensors_control, specials_control, actuators_control = control
    def read_diff(control_dataframe,extract_dict):
        output = []
        for control in control_dataframe.iterrows():
            control = control[1]
            key = (control.loc['信号位号idcode'][1:], control.loc['图号diagram number'])
            if key in extract_dict:
                output.append(check_diff(control, extract_dict[key]))
        return output
    
    def check_diff(control, extract, extension=False, designation=True):
        labels = ['扩展码extensioncode','信号位号idcode', '图号diagram number', '版本rev.', '信号说明designation', "安全分级/分组Safetyclass/division"]
        if not designation:
            labels.remove('信号说明designation')
        if not extension:
            labels.remove('扩展码extensioncode')
        for col in labels:
            if col == '信号位号idcode':
                diff_check = control[col][1:] != extract[col][1:]
            else:
                diff_check = control[col] != extract[col]
            if diff_check:
                print(control[col], extract[col])
        
    read_diff(sensors_control, sensors)
        
                
main('qqq.dxf', '2023-7-28清单校对.xlsx')
    
