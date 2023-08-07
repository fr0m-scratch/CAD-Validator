from utility.utility import *
from cad_entity import CADEntity
from cad_handler import CADHandler
from rich.progress import Progress
from rich.traceback import install
import sys
import os
from ezdxf.addons import odafc
import time
install()

def main(graphfilepath, filepath):
    if re.search(r'\/', graphfilepath) is None:
        graphfilepath = get_resource_path(graphfilepath)
    converted_path = graphfilepath.replace('.dwg', '.dxf')
    if not os.path.exists(converted_path):
        odafc.convert(graphfilepath, converted_path, version='R2010')
        while not os.path.exists(converted_path):
            time.sleep(500)
    acad = CADHandler(converted_path)
    
    os.remove(converted_path)
    control, ref = read_sheet(filepath)
    
    sensors, specials, actuators = acad.accessories.values()
    
    extracted = acad.accessories_to_map(sensors, specials, actuators)
    
  
    sensors, specials, actuators = extracted
    
    
    for _ in control:
        _['refindex'] = _.index+2
    sensors_control, specials_control, actuators_control = control
    output_diff(filepath, ref, (sensors, specials, actuators, sensors_control, specials_control, actuators_control))
    return filepath
    
# main(r"C:\Users\intern\Desktop\cad-ocr\QC测试\3\TFL SAMA(1).dwg",r"C:\Users\intern\Desktop\cad-ocr\QC测试\3\H801202S-KK-3-TFL-01-QC-00001^#3机组低压给水加热器系统（TFL）DCS IO清单^A4^HL317TFL001E05045GN.xlsx")
# main(r'C:\Users\intern\Desktop\cad-ocr\QC测试\2\TFP(1).dwg',r'C:\Users\intern\Desktop\cad-ocr\QC测试\2\H801202S-KK-3-TFP-01-QC-00001^#3机组电动主给水泵系统（TFP）DCS IO清单^A4^HL317TFP001E05045GN B版.xlsx' )
main("TFS逻辑图.dwg", "control.xlsx")
main(r"C:\Users\intern\Desktop\cad-ocr\t.dwg", r"C:\Users\intern\Desktop\cad-ocr\QC测试\1\H801202S-KK-3-TFR-01-QC-00001^#3机组给水加热器疏水回收系统（TFR）DCS IO清单^A4^HL317TFR001E05045GN.xlsx")