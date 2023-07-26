from utility.utility import *
from cad_entity import CADEntity
from cad_handler import CADHandler
import shutil
from rich.progress import Progress
from rich.traceback import install
import sys
import os
install()
import aspose.cad as cad


if __name__ == '__main__':
    image = cad.fileformats.cad.CadImage.load("TFS.dwg")
    castedImage = cast(cad.fileformats.cad.CadImage, image)

    for entity in castedImage.entities:
        if entity.type_name == cad.fileformats.cad.cadconsts.CadEntityTypeName.TEXT:
            text = cast(cad.fileformats.cad.cadobjects.CadText, entity)
            print(text.default_value)
        if entity.type_name == cad.fileformats.cad.cadconsts.CadEntityTypeName.MTEXT:
            mtext = cast(cad.fileformats.cad.cadobjects.CadMText, entity)
            print(mtext.full_clear_text)
        if entity.type_name == cad.fileformats.cad.cadconsts.CadEntityTypeName.ATTRIB:
            attrib = cast(cad.fileformats.cad.cadobjects.attentities.CadAttrib, entity)
            print(attrib.default_text)
        if entity.type_name == cad.fileformats.cad.cadconsts.CadEntityTypeName.ATTDEF:
            attDef = cast(cad.fileformats.cad.cadobjects.attentities.CadAttDef, entity)
            print(attDef.definition_tag_string)
        if entity.type_name == cad.fileformats.cad.cadconsts.CadEntityTypeName.INSERT:
            insert = cast(cad.fileformats.cad.cadobjects.CadInsertObject, entity)
            for block in castedImage.block_entities.values:
                if block.original_block_name == insert.name:
                    for blockEntity in block.entities:
                        if blockEntity.type_name == cad.fileformats.cad.cadconsts.CadEntityTypeName.ATTDEF:
                            attDef = cast(cad.fileformats.cad.cadobjects.attentities.CadAttDef, blockEntity)
                            print(attDef.prompt_string)
            for e in insert.child_objects:
                if e.type_name == cad.fileformats.cad.cadconsts.CadEntityTypeName.ATTRIB:
                    attrib = cast(cad.fileformats.cad.cadobjects.attentities.CadAttrib, e)
                    print(attrib.definition_tag_string)
