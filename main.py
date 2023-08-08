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
    #初始化输入文件
    if re.search(r'\/', graphfilepath) is None:
        graphfilepath = get_resource_path(graphfilepath)
    converted_path = graphfilepath.replace('.dwg', '.dxf')
    if not os.path.exists(converted_path):
        odafc.convert(graphfilepath, converted_path, version='R2010')
        while not os.path.exists(converted_path):
            time.sleep(500)
    acad = CADHandler(converted_path)
    
    os.remove(converted_path)
    
    #读取IO清单并unify数据比对格式
    control, ref = read_sheet(filepath)
    
    sensors, specials, actuators = acad.accessories.values()
    extracted = acad.accessories_to_map(sensors, specials, actuators)
    sensors, specials, actuators = extracted

    for _ in control:
        _['refindex'] = _.index+2
    sensors_control, specials_control, actuators_control = control

    #校对并输出结果
    output_diff(filepath, ref, (sensors, specials, actuators, sensors_control, specials_control, actuators_control))
    return filepath
    
if __name__ == "__main__":
    # print(sys.argv[1], '',sys.argv[2])
    main(sys.argv[1], sys.argv[2])