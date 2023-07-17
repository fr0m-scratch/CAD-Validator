from cad_entity import CADEntity
import win32com.client
from utility import *
import pickle
from rich.progress import Progress
import re
import pandas as pd
class CADHandler:
    def __init__(self, filepath= None):
        self.acad = None
        self.doc = None
        self.modelspace = None
        self.entities = []
        if filepath is None:
            self.inspect_CAD()
            self.load_entities()
        else:
            self.load_entities_from_file(filepath)
            
    def set_pos(self):
        for entity in self.entities:
            entity.set_pos()
    
    def set_params(self):
        self.set_pos()
        self.set_diagram_number()
        self.set_sorting_pos()
        
    def inspect_CAD(self):
        self.acad = win32com.client.Dispatch("AutoCAD.Application")
        self.doc = self.acad.ActiveDocument
        self.modelspace = self.doc.modelspace 
    
    def set_sorting_pos(self):
        for entity in self.entities:
            entity.set_sorting_pos()
    
    def load_entities(self):
        progress = Progress()
        find = progress.add_task("[cyan2]Loading Entities:", total=3644)
        progress.start()
        for entity in self.modelspace:
            progress.update(find, advance=1)
            if entity.ObjectName != "AcDbZombieEntity":
                params = parameter_retriving(entity)
                cad_entity = CADEntity(params)
                self.entities.append(cad_entity)
            if entity.ObjectName == "AcDbBlockReference":
                for attribute in entity.GetAttributes():
                    params = parameter_retriving(attribute)
                    cad1_entity = CADEntity(params)
                    self.entities.append(cad1_entity)
                    cad1_entity.parent = cad_entity
        with open('entity_list.pkl', 'wb') as f:
            pickle.dump(self.entities, f)
        
    def load_entities_from_file(self, filepath):
        with open(filepath, 'rb') as f:
            self.entities = pickle.load(f)
    
    def find_vital_entities(self):
        vital_coordinates = []
        nums = []
        for entity in self.entities:
            if entity.type == "AcDbAttributeDefinition":
                if entity.tag.startswith("H"):
                    vital_coordinates.append((entity.pos, entity.tag))
                elif entity.text == "1":
                    nums.append((entity.pos, entity.tag))
        return vital_coordinates,nums
    
    def generate_graph_rectangles(self, anchors, width, height, w_correction, h_correction):
        rectangles = [((r1[0][0]-width+w_correction, r1[0][0]+w_correction, r1[0][1]+h_correction, r1[0][1]+height+h_correction), r1[1]) for r1 in anchors]
        return rectangles
    
    def generate_diagram_numbers(self, rectangles, tags):
        regions_with_diagrams = []
        for tag in tags:
            coordinates, tag = tag
            prefix, region = find_graph_id(coordinates, rectangles)
            if prefix[9] == "1":
                diagram_number = "FD-" + tag
            elif prefix[9] == "2":
                diagram_number = "SAMA-" + tag
            regions_with_diagrams.append((region[0], diagram_number))
        return regions_with_diagrams
                
    def set_diagram_number(self):
        anchors, tags = self.find_vital_entities()
        regions = self.generate_graph_rectangles(anchors,420, 297, 45.3, -12.6)
        regions_with_diagrams = self.generate_diagram_numbers(regions, tags)
        for entity in self.entities:
            for region, diagram_number in regions_with_diagrams:
                x,y = entity.pos
                x1, x2, y1, y2 = region
                if x1<=x<=x2 and y1<=y<=y2:
                    entity.diagram_number = diagram_number
                    break
    def identify_accessories(self):
        sensors = []
        special_ios = []
        actuators = []
        sys = sorted([entity for entity in self.entities if entity.tag and entity.text
                      and re.search("INSSYS", entity.tag)], key = lambda x: x.sorting_pos)
        ide = sorted([entity for entity in self.entities if entity.tag and entity.text
                      and re.search("INSIDE", entity.tag)], key = lambda x: x.sorting_pos)
        for sys_entity, ide_entity in zip(sys, ide):
            entity = "YTFS" + ide_entity.text + sys_entity.text
            if sys_entity.text == "SY":
                sys_entity.text = entity
                special_ios.append(sys_entity)
            elif len(ide_entity.text) == 3:
                sys_entity.text = entity
                sensors.append(sys_entity)
            else:
                ide_entity.text = 'YTFS' + ide_entity.text
                sys_entity.relevance = ide_entity
                actuators.append(sys_entity)
        scattered = [entity for entity in self.entities if entity.text and 
                     re.search("YTFS[^\s][^-]+$", entity.text)]
        unmatched = []
        for entity in scattered:
            if entity.text.endswith("SY"):
                special_ios.append(entity)
            else:
                unmatched.append(entity)
        unmatched_regions = self.get_target_regions(unmatched, 20, 20, 20, 20)
        arrows, varrows = self.get_arrows(unmatched_regions)
        arrow_regions = self.get_target_regions(arrows, 4.1, 3, 90, 90)
        varrow_regions = self.get_target_regions(varrows, 4,4,4,4)
        for entity in self.entities:
            if entity.type == "AcDbMText":
                if entity.color == 4:
                    relevance = self.check_within(entity.pos, unmatched_regions)
                    if relevance is not None:
                        entity.relevance = relevance
                        actuators.append(entity)
                elif entity.text.isalpha() and entity.height == 2:
                    relevance = self.check_within(entity.pos, arrow_regions)
                    if relevance is not None:
                        entity.relevance = relevance.relevance
                        actuators.append(entity)
                        continue
                    relevance = self.check_within(entity.pos, varrow_regions)
                    if relevance is not None:
                        entity.relevance = relevance.relevance
                        actuators.append(entity)
        return sensors, set(special_ios), actuators
        
    def get_target_regions(self, entities, up, down, left, right):
        res = []
        for entity in entities:
            region = self.generate_rectangle(entity.pos, up, down, left, right)
            res.append((region, entity))
        return res
        
    def generate_rectangle(self, coordinates, up, down, left, right):
        x, y = coordinates[0], coordinates[1]
        bounds = (x-left, x+right, y-down, y+up)
        return bounds
      
    def get_arrows(self, regions):
        arrows = []
        varrows = []
        for entity in self.entities:
            match entity.type:
                case "AcDbPolyline":
                    if 3.8<entity.length< 5:                                
                        relevance = self.check_within(entity.pos, regions)
                        if relevance is not None:
                            entity.relevance = relevance
                            varrows.append(entity)
                case "AcDbHatch":
                    if 1.29< entity.area <1.3:
                        relevance = self.check_within(entity.pos, regions)
                        if relevance is not None:
                            entity.relevance = relevance
                            arrows.append(entity)
        return arrows, varrows
    
    def check_within(self, coordinates, rectangles):
        x,y = coordinates
        for rec in rectangles:
            x1, x2, y1, y2 = rec[0]
            if x1 <= x <= x2 and y1 <= y <= y2:
                return rec[1]
        return None
    
    def export_to_excel(self, filepath):
        dataframe_map = {}
        for entity in self.entities:
            dataframe_map[entity.handle] = entity.entity_to_entry()
        df = pd.DataFrame.from_dict(dataframe_map, 
                                    orient = "index",
                                    columns = ["Handle", "Type", "Layer", 
                                               "Tag", "Text", "Diagram Number",
                                               "Relevance", "Position", "Sorting Position"])
        
        df.to_excel(filepath)
        
    def export_accessories_to_excel(self, sensors, special_ios, actuators, filepath):
        sensors_dict = {}
        special_ios_dict = {}
        actuators_dict = {}
        for sensor in sensors:
            sensors_dict[sensor.handle] = sensor.accessorie_to_entry()
        for special_io in special_ios:
            special_ios_dict[special_io.handle] = special_io.accessorie_to_entry()
        for actuator in actuators:
            actuators_dict[actuator.handle] = actuator.actuators_to_entry()
        df1 = pd.DataFrame.from_dict(sensors_dict, 
                                    orient = "index",
                                    columns = ["id_codes", "Diagram Number"])
        df2 = pd.DataFrame.from_dict(special_ios_dict,
                                     orient = "index",
                                     columns= ["id_codes", "Diagram Number"])
        df3 = pd.DataFrame.from_dict(actuators_dict,
                                     orient = "index",
                                     columns = ["id_codes", "ex_codes", "Diagram Number"])
        output = open(filepath, "w")
        writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
        df1.to_excel(writer, sheet_name='SENSOR_IO', index=False)
        df3.to_excel(writer, sheet_name='ACTUATOR_IO', index=False)
        df2.to_excel(writer, sheet_name='SPECIAL_IO', index=False)
        writer.close()
        output.close()
        
        
        