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

    path = filepath
    control, ref = read_sheet(filepath)
    
    sensors, specials, actuators = acad.accessories.values()
    extracted = acad.accessories_to_df(sensors, specials, actuators)
    
    for c, e in zip(control, extracted):
        unify(c)
        unify(e)
    sensors, specials, actuators = extracted
    sensors_control, specials_control, actuators_control = control
    for c,ce ,e, ee in zip(actuators_control['信号位号idcode'], actuators_control['扩展码extensioncode'], actuators['信号位号idcode'], actuators['扩展码extensioncode']):
        print(c, ce, e, ee)
    
    progress = Progress()
    uni = progress.add_task("[cyan2]Unifydata", total=1)
    progress.start()
    progress.update(uni, advance=1)
    progress.stop()
    #Mark and Write Diff to Master File
    mark = progress.add_task("[cyan2]Markdiff", total=1)
    output_diff(path, ref, (sensors, specials, 
                                    actuators, sensors_control, 
                                    specials_control, actuators_control))
    
    
    progress.start()
    progress.update(mark, advance=1)
    progress.stop()
    return filepath
    
if __name__ == "__main__":
    print(sys.argv[1], '#######',sys.argv[2])
    main(sys.argv[1], sys.argv[2])