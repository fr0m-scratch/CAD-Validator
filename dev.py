from utility.utility import *
from cad_entity import CADEntity
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
import pandas as pd
install()
from ezdxf.addons.dxf2code import entities_to_code
import datetime
# Load a DWG file


acad = CADHandler("qq.dxf")

# q = []
# for _ in acad.entities:
#     if _.text and 0<len(_.text) < 3 and _.text[-1] in ['O', 'P', 'D', 'C'] and _.text != 'NC' and _.text != 'TD':
#         print(_.text)
#         q.append(_)
# print(len(q))

# # b = [_.text for _ in b]
# # print(sorted(b))

# # for _ in a:
# #     if len(_.text) < 6:
# #         _.text = _.sys + _.text 
# #         print(_.text)   
        
        
control, ref = read_sheet(r"C:\Users\intern\Desktop\QC测试\3\H801202S-KK-3-TFL-01-QC-00001^#3机组低压给水加热器系统（TFL）DCS IO清单^A4^HL317TFL001E05045GN.xlsx")
# control,ref = read_sheet(r"C:\Users\intern\Desktop\QC测试\1\H801202S-KK-3-TFR-01-QC-00001^#3机组给水加热器疏水回收系统（TFR）DCS IO清单^A4^HL317TFR001E05045GN.xlsx")
# control, ref = read_sheet("2023-7-28清单校对.xlsx")
# control, ref = read_sheet("control.xlsx")


a,b,c = acad.accessories.values()
sensors_control, specials_control, actuators_control = control

for _, n in zip((a,b,c), (sensors_control, specials_control, actuators_control)):
    print(len(_), len(n))
    
cont = list(specials_control['信号位号idcode'])
text = [i.text for i in b]

for i, j in zip(b, range(len(list(specials_control['信号位号idcode'])))):
    i.text = i.text[1:]
    
   
for j in range(len(cont)):
    k = cont[j]
    k = k[1:]
    cont[j] = k
deletion = []
for i in b:
    print(i.text)
    if i.text in cont:
       cont.remove(i.text)
       deletion.append(i)
        
for _ in deletion:
    b.remove(_)
print(cont)
print(b)
        
def actuators_check():
    
    a,b,c = acad.accessories.values()
    sensors_control, specials_control, actuators_control = control

    for _, n in zip((a,b,c), (sensors_control, specials_control, actuators_control)):
        print(len(_), len(n))
    a,b,c = acad.accessories_to_df(a,b,c)
    dict = {}
    for _ in actuators_control.iterrows():
        if not (_[1]['信号位号idcode'][1:], _[1]['扩展码extensioncode']) in dict:
            dict[(_[1]['信号位号idcode'][1:], _[1]['扩展码extensioncode'])] = 1
        else:
            dict[(_[1]['信号位号idcode'][1:], _[1]['扩展码extensioncode'])] += 1
    print(len(dict))

    dicte = {}
    for _ in c.iterrows():
        if (_[1]['信号位号idcode'][1:], _[1]['扩展码extensioncode']) == ('TFP220VL', 'FD'):
            print(_)
        if not (_[1]['信号位号idcode'][1:], _[1]['扩展码extensioncode']) in dicte:
            dicte[(_[1]['信号位号idcode'][1:], _[1]['扩展码extensioncode'])] = 1
        else:
            dicte[(_[1]['信号位号idcode'][1:], _[1]['扩展码extensioncode'])] += 1
    print(len(dicte))
    deletion = []
    for _ in dicte:
        if _ in dict:
            del dict[_]
            deletion.append(_)
    for _ in deletion:
        del dicte[_]
    print(dicte)
    print(dict)




# # spe = list(specials_control['信号位号idcode'])
# # print(len(spe))

# # for i in range(len(spe)):
# #     spe[i] = spe[i][1:]
    
# # for _ in a:
# #     if _.text.startswith('Y'):
# #         _.text = _.text[1:]
# #     if _.text in spe:
# #         spe.remove(_.text)
# # print(len(spe))
# # print(spe)

# a,b,c = acad.accessories.values()
# for _ in (a,b,c):
#     print(len(_))
# a,b,c = acad.accessories_to_df(a,b,c)





# dict = {}
# for _ in c.iterrows():
#     if not (_[1]['信号位号idcode'], _[1]['扩展码extensioncode']) in dict:
#         dict[(_[1]['信号位号idcode'], _[1]['扩展码extensioncode'])] = 1
#     else:
#         dict[(_[1]['信号位号idcode'], _[1]['扩展码extensioncode'])] += 1
# print(dict)
    