import ezdxf
from ezdxf.addons import odafc
from utility.utility import *
#import dwg file
doc_path = dwg_to_dxf("TFS.dwg")
doc = ezdxf.readfile(doc_path)


output = open("output.txt", 'w')

#iterate all "Text" entities
figure_count = 0
modelspace = doc.modelspace()
for entity in modelspace:
    if entity.dxftype() == 'TEXT':
        print('TEXT:', entity.dxf.text, file=output)
    elif entity.dxftype() == 'MTEXT':
        print('MTEXT:', entity.plain_text())
    elif entity.dxftype() == 'ATTRIB':
        print('ATTRIB:', entity.dxf.text, file=output)
    elif entity.dxftype() == 'ATTDEF':
        print('ATTDEF:', entity.dxf.tag, file=output)
    elif entity.dxftype() == 'INSERT':
        
        aq = 0
        for attrib in entity.attribs:
            aq += 1
            if attrib.dxf.text.strip() is not None:
                print('ATTRIB:', attrib.dxf.tag, file=output)
        if aq > 0:
            figure_count += 1

print(f'figure count: {figure_count}')
output.close()