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

acad = CADHandler(get_resource_path("TFS.dxf"))
print(acad.accessories)
