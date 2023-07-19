from cad_entity import CADEntity
from pyautocad import Autocad
import pandas as pd
import win32com.client
from utility.utility import *
import pickle
from rich.progress import Progress

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument
entity_list = []
progress = Progress()
find = progress.add_task("[cyan2]Loading Entities:", total=3680)
progress.start()
for entity in doc.ModelSpace:
    progress.update(find, advance=1)
    if entity.ObjectName != "AcDbZombieEntity":
        params = parameter_retriving(entity)
        cad_entity = CADEntity(params)
        entity_list.append(cad_entity)
        if entity.ObjectName == "AcDbBlockReference":
            for attribute in entity.GetAttributes():
                params = parameter_retriving(attribute)
                cad1_entity = CADEntity(params)
                entity_list.append(cad1_entity)
                cad1_entity.relevance = cad_entity
with open('entity_list.pkl', 'wb') as f:
    pickle.dump(entity_list, f)