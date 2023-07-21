from utility.utility import *
from cad_entity import CADEntity
from cad_handler import CADHandler
import shutil
from rich.progress import Progress
from rich.traceback import install
import sys
import os
install()

def main(filepath, mode):
    #CAD Info Extration
    if mode  == "load":
        acad = CADHandler(get_resource_path(".\data\entity_list.pkl"))
    elif mode == "read":
        acad = CADHandler()
         
    #Compared Data Unification
    targetpath = filepath
    destpath = filepath[:-5] + "_copy.xlsx"
    path = shutil.copy(targetpath, destpath)
    
    control, ref = read_sheet(path)
    
    sensors, specials, actuators = acad.accessories.values()
    extracted = acad.accessories_to_df(sensors, specials, actuators)
    for c, e in zip(control, extracted):
        unify(c)
        unify(e)
    sensors, specials, actuators = extracted
    sensors_control, specials_control, actuators_control = control
    progress = Progress()
    uni = progress.add_task("[blue]Unify[cyan2]data", total=1)
    progress.start()
    progress.update(uni, advance=1)

    #Mark and Write Diff to Master File
    mark = progress.add_task("[blue]Mark[cyan2]diff", total=1)
    output_diff(path, ref, (sensors, specials, 
                                    actuators, sensors_control, 
                                    specials_control, actuators_control))
    progress.start()
    progress.update(mark, advance=1)
    return destpath
    
if __name__ == "__main__":
    print(sys.argv[1])
    main(sys.argv[1], sys.argv[2])