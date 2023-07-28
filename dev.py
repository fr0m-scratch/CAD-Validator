from utility.utility import *
from cad_entity import CADEntity
from cad_handler import CADHandler
import shutil
from rich.progress import Progress
from rich.traceback import install
import sys
import ezdxf
import os
install()
from ezdxf.addons.dxf2code import entities_to_code

entities = ezdxf.readfile("TFS.dxf").modelspace()

defs = [(e.dxftype(), e.dxf.text) for e in entities if e.dxf.hasattr('tag') and 'YTFS' in e.dxf.text]
print(defs)
cent = [CADEntity(e.dxftype(), e.dxfattribs() ) for e in entities]





# def main(graphfilepath, filepath, mode):
#     #CAD Info Extration
#     if mode  == "load":
#         acad = CADHandler(get_resource_path(".\data\entity_list.pkl"), True)
#     elif mode == "read":
#         if re.search(r'\/', graphfilepath) is None:
#             graphfilepath = get_resource_path(graphfilepath)
#         acad = CADHandler(graphfilepath, False)
#     a = sum([1 for _ in acad.entities if _.bounding_box is not None])
#     print(a)
    
# main("TFS.dwg", "control.xlsx", "load")