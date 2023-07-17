from pyautocad import Autocad
import pandas as pd
import win32com.client
from utility import *
import time
from rich.progress import Progress

#Progress Bar 
with Progress() as progress:
    actuators = []
    arrows = []
    varrows = []
    #Arbitrary parameters though returned by programs; Need to refactor
    arrow_regions = [(14787.590118458344, 14967.590118458344, -1829.9884234355259, -1822.888423435526), (14787.590118458344, 14967.590118458344, -1835.957382778197, -1828.857382778197), (15233.970338667397, 15413.970338667397,
-1849.8405773153409, -1842.740577315341)]
    varrow_regions = [(15640.612836083143, 15648.612836083143, -3307.9525005456217, -3299.9525005456217), (16133.93924609066, 16141.93924609066, -3306.109394125974, -3298.109394125974)]  

    find = progress.add_task("[red]Finding Coordinates...", total=3680)
    add_en = progress.add_task("[green]Adding Entries...", total=3680)
    
    vital_coordinates = []
    sheet_name = []
    # output = open("output.txt", 'w')
    acad = win32com.client.Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    dataframe_map = {}
    
    
    
    def read_entity(entity, map):
        print(parameter_retriving(entity))
    
    def find_vital_coordinates(doc):
        vital_coordinates = []
        handle = None
        for entity in doc.ModelSpace:
            if entity.objectName == "AcDbAttributeDefinition" and entity.TagString.startswith("H"):
                vital_coordinates.append((tuple(entity.InsertionPoint), entity.TagString))
                handle = hex(int(entity.Handle, 16) + 2)[2:].upper()
            elif entity.Handle == 0:
                sheet_name.append(entity.TagString)
            progress.update(find, advance=1)
        return vital_coordinates
    
    datapoints = find_vital_coordinates(doc)
    rectangles = generate_graph_rectangles(datapoints, 420, 297, 45.3, -12.6)
    regions = [(14615.73452103801, 14655.73452103801, -1838.7400860374682, -1798.7400860374682), (14864.348194660093, 14904.348194660093, -1839.5206256509964, -1799.5206256509964), (15311.344909970081, 15351.344909970081, -1859.8555252168007, -1819.8555252168007), (15618.395970117243, 15658.395970117243, -3327.9702875672237, -3287.9702875672237), (16111.925881510117, 16151.925881510117, -3326.1271811475744, -3286.1271811475744)]
    
    def add_entry(entity, map):
        match entity.ObjectName:
            case 'AcDbText':
                num, graph_id = set_graph_id(tuple(entity.InsertionPoint), rectangles)
                map[entity.Handle] = [entity.ObjectName, entity.TextString, tuple(entity.InsertionPoint), num, graph_id, entity.Layer]
                # print(entity.TextString, file=output, end=' :')
                # print(getattr(entity, "InsertionPoint",None), entity.Layer, file=output)      
            case 'AcDbMText':
                if entity.color == 4:
                    coordinates = tuple(entity.InsertionPoint[0:2])
                    check = check_within(coordinates, regions)
                    if check:
                        actuators.append((check[0], entity.TextString, check[1]))
                arrow_check = check_within(tuple(entity.InsertionPoint[0:2]), arrow_regions, [1,1,2])
                if arrow_check and entity.TextString.isalpha() and entity.height == 2:
                    actuators.append((arrow_check[0], entity.TextString, arrow_check[1]))
                varrow_check = check_within(tuple(entity.InsertionPoint[0:2]), varrow_regions, [3,4])
                if varrow_check:
                    actuators.append((varrow_check[0], entity.TextString, varrow_check[1]))
                num, graph_id = set_graph_id(tuple(entity.InsertionPoint), rectangles)
                map[entity.Handle] = [entity.ObjectName, entity.TextString, tuple(entity.InsertionPoint), num, graph_id, entity.Layer]
                # print(entity.TextString, file=output, end=' :')
                # print(getattr(entity, "InsertionPoint",None), entity.Layer, file=output)      
            case 'AcDbBlockReference':
                num, graph_id = set_graph_id(tuple(entity.InsertionPoint), rectangles)
                map[entity.Handle] = [entity.ObjectName, "################", tuple(entity.InsertionPoint), num, graph_id, entity.Layer]
                if not entity.HasAttributes:
                    return
                for attrib in entity.GetAttributes():
                    if attrib.TextString != "":
                        num, graph_id = set_graph_id(tuple(entity.InsertionPoint), rectangles)
                        # if attrib.TagString == "INSDGN":
                        #     print(attrib.TextString)
                        #     print(type(attrib.TextString))
                        #     print(len(attrib.TextString))
                        #     print(attrib.TextString.strip()== "")
                        map[attrib.Handle] = [attrib.ObjectName, f'{attrib.TagString}: {attrib.TextString}', tuple(attrib.InsertionPoint), num, graph_id, attrib.Layer]
                        # print(f'{attrib.TagString}: {attrib.TextString}', file=output, end=' :')
                        # print(getattr(entity, "InsertionPoint",None), entity.Layer, file=output)      
            case "AcDbAttributeDefinition":
                num, graph_id = set_graph_id(tuple(entity.InsertionPoint), rectangles)
                map[entity.Handle] = [entity.ObjectName, f'{entity.TagString}: {entity.TextString}', tuple(entity.InsertionPoint), num, graph_id, entity.Layer]
                # print(f'{entity.TagString}: {entity.TextString}', file=output, end=' :')
                # print(getattr(entity, "InsertionPoint",None), entity.Layer, file=output)  
            case "AcDbHatch":
                if 1.29<entity.Area <1.3:
                    coordinates = get_center(tuple(entity.GetBoundingBox())[0:2])
                    check = check_within(coordinates, regions)
                    if check:
                       arrows.append(check) 
            case "AcDbPolyline":
                # print(dir(entity))                                                          
                if 3.8<entity.Length< 5:                                
                    coordinates = []
                    for i in range(len(entity.coordinates)//2):
                        coordinates.append(entity.Coordinate(i))
                    coordinates = get_center(coordinates)
                    check = check_within(coordinates, regions)
                    if check:
                        varrows.append(check)
                

    for entity in acad.ActiveDocument.ModelSpace:
        read_entity(entity, dataframe_map)
        add_entry(entity, dataframe_map)
        progress.update(add_en, advance=1)
        
    df = pd.DataFrame.from_dict(dataframe_map, 
                                orient='index', 
                                columns=['Type', 'Text', 'InsertionPoint', 'diagram\nnumber', 'Graph_id', 'Layer']
                                )
    df.to_excel("raw_output.xlsx")
    df.sort_values(by=['diagram\nnumber', 'InsertionPoint'], inplace=True)

    df.to_excel("output.xlsx")
    
    
    
    #TOBE Refactored
    arrow_regions = [create_regions(arrow[1], 4.1, 3, 90, 90) for arrow in arrows]
    
    arrow_indice = [arrow[0] for arrow in arrows]
    
    varrow_regions = [create_regions(varrow[1], 4,4,4,4) for varrow in varrows]
    varrow_indice = [varrow[0] for varrow in varrows]
    
    
    scattered = pd.DataFrame(actuators, columns=['indices', 'Text', 'InsertionPoint'])
    scattered.to_excel("scattered.xlsx", index=False)