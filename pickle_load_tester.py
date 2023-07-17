
import pickle
from utility import *
from cad_entity import CADEntity
from cad_handler import CADHandler
import re

acad = CADHandler()
acad.set_params()
a,b,c = acad.identify_accessories()

acad.export_accessories_to_excel(a,b,c,"accessories.xlsx")