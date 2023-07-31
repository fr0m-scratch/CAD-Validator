from utility.utility import *
from new_entity import CADEntity
from cad_handler import CADHandler
import shutil
from rich.progress import Progress
from rich.traceback import install
import sys
import ezdxf
from ezdxf.enums import TextEntityAlignment
import os
from ezdxf.entities import *
from ezdxf.addons import odafc
install()
from ezdxf.addons.dxf2code import entities_to_code
import datetime
start = datetime.datetime.now()
entities = ezdxf.readfile("t.dxf").modelspace()

defs = [(e.dxftype(), e.dxf.tag) for e in entities if e.dxf.hasattr('tag') and 'TFS' in e.dxf.tag]

ents = []
with Progress() as progress:
    loading = progress.add_task("[cyan2]Loading DXF", total=len(entities))
    for e in entities:
        if e.dxftype() != "ACAD_PROXY_ENTITY":
            ents.append(CADEntity(e, e.dxfattribs()))
        if e.dxftype() == "INSERT":
            for attr in e.attribs:
                ents.append(CADEntity(attr, attr.dxfattribs()))
        progress.update(loading, advance=1)

# for e in ents:
#     print(type(e.pos))
acad = CADHandler(ents)

acad.set_params()

for _ in acad.accessories.values():
    print(len(_))
                
