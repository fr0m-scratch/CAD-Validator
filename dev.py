from utility.utility import *
from cad_entity import CADEntity
from cad_handler import CADHandler
import shutil
from rich.progress import Progress
from rich.traceback import install
import sys
import os
install()

acad = CADHandler(".\data\entity_list.pkl")


sensors, specials, actuators = acad.accessories.values()

extracted = acad.accessories_to_df(sensors, specials, actuators)

sensors, specials, actuators = extracted
print(actuators["信号说明designation"])