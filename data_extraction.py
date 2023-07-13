from pyautocad import Autocad
import pandas as pd
import win32com.client
from utility import *


vital_coordinates = []
sheet_name = []
# output = open("output.txt", 'w')
acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument
dataframe_map = {}
# try:
#     open("Coordinates.txt", 'r')
# except FileNotFoundError:
#     find_vital_coordinates(doc)
#     open("Coordinates.txt", 'r')
    
def find_vital_coordinates(doc):
    vital_coordinates = []
    handle = None
    for entity in doc.ModelSpace:
        if entity.objectName == "AcDbAttributeDefinition" and entity.TagString.startswith("H"):
            vital_coordinates.append((entity.InsertionPoint, entity.TagString))
            handle = hex(int(entity.Handle, 16) + 2)[2:].upper()
            
        elif entity.Handle == 0:
            
            sheet_name.append(entity.TagString)
    return vital_coordinates

datapoints = find_vital_coordinates(doc)


def generate_graph_rectangles(datapoints, width, height, w_correction, h_correction):
    rectangles = [((r1[0][0]-width+w_correction, r1[0][0]+w_correction, r1[0][1]+h_correction, r1[0][1]+height+h_correction), r1[1]) for r1 in datapoints]
    return rectangles
rectangles = generate_graph_rectangles(datapoints, 420, 297, 45.3, -12.6)

def set_graph_id(coordinates, rectangles):
    for i in range(len(rectangles)):
        x1, x2, y1, y2 = rectangles[i][0]
        graph_id = rectangles[i][1]
        x, y = coordinates[0], coordinates[1]
        if x1 <= x <= x2 and y1 <= y <= y2:
            return i, graph_id
    return None, None

def add_entry(entity, map):
    match entity.ObjectName:
        case 'AcDbText':
            num, graph_id = set_graph_id(entity.InsertionPoint, rectangles)
            map[entity.Handle] = [entity.ObjectName, entity.TextString, entity.InsertionPoint, num, graph_id, entity.Layer]
            # print(entity.TextString, file=output, end=' :')
            # print(getattr(entity, "InsertionPoint",None), entity.Layer, file=output)      
        case 'AcDbMText':
            num, graph_id = set_graph_id(entity.InsertionPoint, rectangles)
            map[entity.Handle] = [entity.ObjectName, entity.TextString, entity.InsertionPoint, num, graph_id, entity.Layer]
            # print(entity.TextString, file=output, end=' :')
            # print(getattr(entity, "InsertionPoint",None), entity.Layer, file=output)      
        case 'AcDbBlockReference':
            if not entity.HasAttributes:
                return
            for attrib in entity.GetAttributes():
                num, graph_id = set_graph_id(entity.InsertionPoint, rectangles)
                # if attrib.TagString == "INSDGN":
                #     print(attrib.TextString)
                #     print(type(attrib.TextString))
                #     print(len(attrib.TextString))
                #     print(attrib.TextString.strip()== "")
                if attrib.TextString != "":
                    print(f'{attrib.TagString}: {attrib.TextString}')
                map[attrib.Handle] = [attrib.ObjectName, f'{attrib.TagString}: {attrib.TextString}', attrib.InsertionPoint, num, graph_id, attrib.Layer]
                # print(f'{attrib.TagString}: {attrib.TextString}', file=output, end=' :')
                # print(getattr(entity, "InsertionPoint",None), entity.Layer, file=output)      
        case "AcDbAttributeDefinition":
            num, graph_id = set_graph_id(entity.InsertionPoint, rectangles)
            map[entity.Handle] = [entity.ObjectName, f'{entity.TagString}: {entity.TextString}', entity.InsertionPoint, num, graph_id, entity.Layer]
            # print(f'{entity.TagString}: {entity.TextString}', file=output, end=' :')
            # print(getattr(entity, "InsertionPoint",None), entity.Layer, file=output)   

for entity in acad.ActiveDocument.ModelSpace:
    add_entry(entity, dataframe_map)

df = pd.DataFrame.from_dict(dataframe_map, 
                            orient='index', 
                            columns=['Type', 'Text', 'InsertionPoint', 'diagram\nnumber', 'Graph_id', 'Layer']
                            )
# df.to_excel("raw_output.xlsx")
# df.sort_values(by=['diagram\nnumber', 'InsertionPoint'], inplace=True)

# df.to_excel("output.xlsx")

