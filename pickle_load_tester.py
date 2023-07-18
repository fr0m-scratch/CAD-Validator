
import pickle
from utility import *
from cad_entity import CADEntity
from cad_handler import CADHandler
import re

acad = CADHandler("entity_list.pkl")
acad.set_params()
at,b,c = acad.identify_accessories()
illustration = acad.retrive_designation()
acad.trim_designation
acad.export_accessories_to_excel(at,b,c,"accessories.xlsx",)
