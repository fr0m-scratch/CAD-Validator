from typing import Any
from cad_entity import CADEntity
import win32com.client
from utility import *
import pickle
from rich.progress import Progress
import re
import pandas as pd
from accessories import accessories
class CADHandler:
    def __init__(self, filepath= None):
        self.acad = None
        self.doc = None
        self.modelspace = None
        self.entities = []
        self.accessories = {}
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
                    vital_coordinates.append(entity)
                elif entity.text == "1":
                    nums.append(entity)
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
        coordinates = [(anchor.pos, anchor.tag) for anchor in anchors]
        tags = [(tag.pos, tag.tag) for tag in tags]
        regions = self.generate_graph_rectangles(coordinates,420, 297, 45.3, -12.6)
        regions_with_diagrams = self.generate_diagram_numbers(regions, tags)
        for entity in self.entities:
            for region, diagram_number in regions_with_diagrams:
                x,y = entity.pos
                x1, x2, y1, y2 = region
                if x1<=x<=x2 and y1<=y<=y2:
                    entity.diagram_number = diagram_number
                    break
        #should be modifies into a seperate function--generate and find rev numbers
        alphas = []
        for entity in self.entities:
            if entity.type == "AcDbMText" and len(entity.text) == 1 and entity.text.isalpha():
                alphas.append(entity)
        revs = {}
        for alpha in alphas:
            if alpha.diagram_number not in revs:
                revs[alpha.diagram_number] = [alpha]
            else:
                revs[alpha.diagram_number].append(alpha)  
        revs = {anchor.diagram_number:sorted(revs[anchor.diagram_number], key=lambda x: 
            x.get_distance_by_c(anchor.pos))[0] for anchor in anchors}
        for entity in self.entities:
            if entity.diagram_number is not None:
                entity.rev = revs[entity.diagram_number].text

    
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
                ide_entity.accessories.append(sys_entity)
                actuators.append(sys_entity)
        scattered = [entity for entity in self.entities if entity.text and 
                     re.search("YTFS[^\s][^-]+$", entity.text)]
        unmatched = []
        for entity in scattered:
            if entity.text.endswith("SY"):
                special_ios.append(entity)
            else:
                unmatched.append(entity)
        unmatched_regions = self.get_target_regions(unmatched, 50, 50, 50, 50)
        arrows, varrows = self.get_arrows(unmatched_regions)
        arrow_regions = self.get_target_regions(arrows, 10, 10, 58, 58)
        varrow_regions = self.get_target_regions(varrows, 10,10,10,10)
        for entity in self.entities:
            if entity.type == "AcDbMText":
                if entity.color == 4:
                    relevance = self.check_within(entity.pos, unmatched_regions, overload=True)
                    if relevance:
                        entity.associates.extend(relevance)
                        entity.relevance = entity.relevance_from_associates()
                        entity.relevance.accessories.append(entity)
                        actuators.append(entity)
                elif entity.text.isalpha() and entity.height == 2:
                    relevance = self.check_within(entity.pos, arrow_regions, overload=True)
                    if relevance:
                        entity.associates.extend(relevance)
                        entity.relevance = entity.relevance_from_associates().relevance
                        entity.relevance.accessories.append(entity)
                        actuators.append(entity)
                        continue
                    relevance = self.check_within(entity.pos, varrow_regions,overload=True)
                    if relevance:
                        entity.associates.extend(relevance)
                        entity.relevance = entity.relevance_from_associates().relevance
                        entity.relevance.accessories.append(entity)
                        actuators.append(entity)
        special_ios = self.special_ios_duplicates_processing(special_ios)
        self.accessories["sensors"] = sensors
        self.accessories["special_ios"] = special_ios
        self.accessories["actuators"] = actuators
        return sensors, special_ios, actuators
        
    def get_target_regions(self, entities, up, down, left, right):
        res = []
        for entity in entities:
            region = self.generate_rectangle(entity.pos, up, down, left, right)
            res.append((region, entity))
        return res
    
    def special_ios_duplicates_processing(self, special_ios):
        text_dict = {}
        for spe in special_ios:
            if spe.text not in text_dict:
                text_dict[spe.text] = [spe]
            else:
                text_dict[spe.text].append(spe)
        duplicates = [item for item in text_dict.values() if len(item) > 1]
        
        for duo in duplicates:
            duo.sort(key=lambda x: x.sorting_pos, reverse = True)
            for spe in duo[1:]:
                special_ios.remove(spe)
        return special_ios
        
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
                            relevance.accessories.append(entity)
                            varrows.append(entity)
                case "AcDbHatch":
                    if 1.29< entity.area <1.3:
                        relevance = self.check_within(entity.pos, regions)
                        if relevance is not None:
                            entity.relevance = relevance
                            relevance.accessories.append(entity)
                            arrows.append(entity)
        return arrows, varrows
    
    def check_within(self, coordinates, rectangles, overload = False):
        x,y = coordinates
        res = []
        for rec in rectangles:
            x1, x2, y1, y2 = rec[0]
            if x1 <= x <= x2 and y1 <= y <= y2:
                if overload:
                    res.append(rec[1])
                    continue
                return rec[1]
        return res or None
    
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
        def extract_text_sort(accessories):
            accessories.sort(key = lambda x: x.diagram_number)
            for asso in accessories:
                asso.designation = [x.text for x in asso.trim_designation()]
                asso.designation.sort(key=lambda x: x, reverse = True)
            asso.designation.sort(reverse = True)
        extract_text_sort(sensors)
        extract_text_sort(special_ios)
        extract_text_sort(actuators)
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
                                    columns = ["id_codes", "Diagram Number", "designation", "rev."])
        df2 = pd.DataFrame.from_dict(special_ios_dict,
                                     orient = "index",
                                     columns= ["id_codes", "Diagram Number", "designation", "rev."])
        df3 = pd.DataFrame.from_dict(actuators_dict,
                                     orient = "index",
                                     columns = ["id_codes", "ex_codes", "Diagram Number", "designation", "rev."])
        output = open(filepath, "w")
        writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
        df1.to_excel(writer, sheet_name='SENSOR_IO', index=False)
        df3.to_excel(writer, sheet_name='ACTUATOR_IO', index=False)
        df2.to_excel(writer, sheet_name='SPECIAL_IO', index=False)
        writer.close()
        output.close()
        
        
    def search_by_handle(self, handle):
        for entity in self.entities:
            if entity.handle == handle:
                return entity
        return None
    
    def contain(self, words):
        res = []
        for entity in self.entities:
            if entity.text is not None:
                text = re.search(words, entity.text)
                if text is not None:
                    res.append(entity)
            if entity.tag is not None:
                tag = re.search(words, entity.tag)
                if tag is not None:
                    res.append(entity)
        return res
    
    def not_contain(self, words):
        res = []
        for entity in self.entities:
            if entity.text is not None:
                text = re.search(words, entity.text)
                if text is None:
                    res.append(entity)
            if entity.tag is not None:
                tag = re.search(words, entity.tag)
                if tag is None:
                    res.append(entity)
        return set(res)
    
    def get_entity_by_handle(self, handle):
        for entity in self.entities:
            if entity.handle == handle:
                return entity
        return None
    
    def retrive_designation(self):
        accessories = []
        for items in self.accessories.values():
            accessories.extend(items)
        regions = self.get_target_regions(accessories, 5.97,70,70,70)
        chinese = self.contain(r'[\u3400-\u9FFF]+')
        chinese_illustration = []
        for chi in chinese:
            relevance = self.check_within(chi.pos, regions, overload=True)
            if relevance is not None and chi.type == "AcDbMText":
                chi.associates.extend(relevance)
                chi.relevance = chi.relevance_from_associates()
                chinese_illustration.append(chi)
        for chiilu in chinese_illustration:
            #can be modified to take a regex pattern as paramter
            if "过程" in chiilu.text:
                chinese_illustration.remove(chiilu)
                continue
            chiilu.relevance.designation.append(chiilu)
        relevant_chinese_illustration = []
        for acco in accessories:
            if acco.designation:
                relevant_chinese_illustration.append(acco.get_closest_designation())
        relevant_chinese_illustration = set(relevant_chinese_illustration)
        illu_regions = self.get_target_regions(relevant_chinese_illustration, 0,20,10,10)
        non_chinese = self.not_contain(r'[\u3400-\u9FFF]+')
        english_illustration = []
        for entity in non_chinese:
            if entity.type == "AcDbMText":
                relevance = self.check_within(entity.pos, illu_regions,overload=True)
                if relevance is not None:
                    entity.associates.extend(relevance)
                    english_illustration.append(entity)
        for engilu in english_illustration:
            if engilu.associates:
                for asso in engilu.associates:
                    if asso.handle == engilu.handle:
                        engilu.associates.remove(asso)
                engilu.relevance = engilu.relevance_from_associates()
                if engilu.relevance.relevance == engilu:
                    engilu.relevance = None
                    continue
                engilu.relevance.relevance.designation.append(engilu)
        illustration = list(relevant_chinese_illustration) + english_illustration
        return illustration
        
    def trim_designation(self):
        for entity in self.entities:
            entity.designation = [x for x in entity.trim_designation()]
            entity.designation.sort(key = lambda x: x.text, reverse = True)
        
        
        