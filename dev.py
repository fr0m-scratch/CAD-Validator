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



a,b,c = acad.accessories_to_df(*acad.accessories.values())

