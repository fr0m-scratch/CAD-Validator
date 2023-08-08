from typing import Any
from cad_entity import CADEntity
from pyautocad import Autocad
from utility.utility import *  
import pickle
from rich.progress import Progress
import re
import pandas as pd
from numba import jit
import ezdxf
from ezdxf.entities import *
from ezdxf.enums import TextEntityAlignment
from ezdxf.addons import odafc
class CADHandler:
    def __init__(self, filepath):
        self.doc = ezdxf.readfile(filepath)
        #实体列表，里面村的是CADEntity
        self.entities = []
        self.systems = []
        #用来进行比对的关键词，可以根据数据内容进行添加来优化比对效果
        self.extension = {'OD', 'CO', 'P', 'C', 'CP', 'FP', 'CD', 'RC', 'OP'}
        self.special_postfix = {'SY', 'MC', 'MY'}
        self.actuators_postfix = {'VV', 'VL'}
        self.accessories = {}
        self.load_entities()
        self.set_params()
        
            
    def set_params(self):
        #初始化CAD实体的各个参数 
        self.set_diagram_number()
        self.set_sorting_pos()
        self.set_rev_safty_entity()
        self.identify_accessories()
        self.retrive_designation()
        self.bind_designation_to_idcode()
  
    def set_sorting_pos(self):
        for entity in self.entities:
            entity.set_sorting_pos()
    
    def load_entities(self):
        #从CAD中读取实体信息，存入self.entities
        progress = Progress()
        find = progress.add_task("[cyan2]Loading Entities:", total=3644)
        progress.start()  
        for i in range(len(self.doc.modelspace())):
            obj = self.doc.modelspace()[i]
            if obj.dxftype() == "INSERT":
                object = CADEntity(obj, obj.dxfattribs())
                self.entities.append(object)
                for attr in obj.attribs:
                    attrib = CADEntity(attr, attr.dxfattribs())
                    attrib.parent = object
                    object.children[attrib.tag] = attrib
                    self.entities.append(attrib)
            elif obj.dxftype() != "ACAD_PROXY_ENTITY":
                self.entities.append(CADEntity(obj, obj.dxfattribs()))
            progress.update(find, advance=1)
        progress.stop()
                       
    def find_vital_entities(self):
        #找到图号相关信息的实体
        vital_coordinates = []
        nums = []
        for entity in self.entities:
            if entity.type == "ATTDEF":
                if entity.tag.startswith("H"):
                    vital_coordinates.append(entity)
                elif entity.text == "1":
                    nums.append(entity)
        return vital_coordinates,nums
    
    def set_diagram_number(self):
        #与generate_graph_rectangles和generate_diagram_numbers配合，找到图号相关信息的实体
        anchors, tags = self.find_vital_entities()
        tags = [(tag.pos, tag.tag) for tag in tags]
        regions = self.generate_graph_rectangles(anchors,420, 297, 45.3, -12.6)
        regions_with_diagrams = self.generate_diagram_numbers(regions, tags)
        for entity in self.entities:
            for region, diagram_number in regions_with_diagrams:
                x,y = entity.pos
                x1, x2, y1, y2 = region
                if x1<=x<=x2 and y1<=y<=y2:
                    entity.diagram_number = diagram_number
                    break
                
    def generate_graph_rectangles(self, anchors, width, height, w_correction, h_correction, true_point = False):
        #根据图号实体，生成一个元组列表， 每个元组[0]是该图号实体的的矩形区域，元组[1]是该图号实体本身
        if true_point:
            #write a for loop try except to identify the error
            rectangles = [((r1.insert[0]-width+w_correction, r1.insert[0]+w_correction, r1.insert[1]+h_correction, r1.insert[1]+height+h_correction), r1) for r1 in anchors]
        else:
            rectangles = [((r1.pos[0]-width+w_correction, r1.pos[0]+w_correction, r1.pos[1]+h_correction, r1.pos[1]+height+h_correction), r1) for r1 in anchors]
        return rectangles
    
    def generate_diagram_numbers(self, rectangles, tags):
        #根据图号信息实体以及前缀生成图号
        regions_with_diagrams = []
        for tag in tags:
            coordinates, tag = tag
            prefix, region = find_graph_id(coordinates, rectangles)
            if prefix[8:11] == "016":
                diagram_number = "FD-" + tag
            elif prefix[8:11] == "026":
                diagram_number = "SAMA-" + tag
            regions_with_diagrams.append((region[0], diagram_number))
        return regions_with_diagrams
            
    def set_rev_safty_entity(self):
        #找到并设置版本,安全分级信息以及工程图唯一实体
        anchors, _ = self.find_vital_entities()
        rev = self.generate_graph_rectangles(anchors, 8,5, -4, -5, True)
        safe = self.generate_graph_rectangles(anchors, 6,4, -31.5, -5, True)
        div = self.generate_graph_rectangles(anchors,11,4, -72,-5, True)
        ent = self.generate_graph_rectangles(anchors, 25,15 , -92, 10, True )
        sys = self.generate_graph_rectangles(anchors, 25,5 , -100, -5, True )
        revs, safes, divs, ents, syss = {}, {}, {}, {}, {}
        for entity in self.entities:
            if entity.type == "MTEXT":
                if self.check_within(entity.insert, rev):
                    revs[self.check_within(entity.insert, rev).diagram_number] = entity
                if self.check_within(entity.insert, safe):
                    safes[self.check_within(entity.insert, safe).diagram_number] = entity
                if self.check_within(entity.insert, div):
                    divs[self.check_within(entity.insert, div).diagram_number] = entity
                if self.check_within(entity.insert, sys):
                    syss[self.check_within(entity.insert, sys).diagram_number] = entity
        bounds = self.contain('-')
        for entity in bounds:
            if self.check_within(entity.insert, ent):
                if self.check_within(entity.insert, ent).diagram_number not in ents:
                    ents[self.check_within(entity.insert, ent).diagram_number] = []
                ents[self.check_within(entity.insert, ent).diagram_number].append(entity)
        for entity in self.entities:
            if entity.diagram_number is not None:
                entity.rev = entity.diagram_number in revs and revs[entity.diagram_number].text or None
                
                entity.safety_class = safes[entity.diagram_number].text + '/' + divs[entity.diagram_number].text
                
                entity.diagram_entity = entity.diagram_number in ents and ents[entity.diagram_number] or None
                if entity.diagram_entity and entity.diagram_entity[0]:
                    entity.diagram_entity = entity.diagram_entity[0].text.split('-')[0]
                
                entity.sys = entity.diagram_number in syss and syss[entity.diagram_number].text or None
                self.systems.append(entity.sys)
        self.bounds = ents
        self.systems = list(set(self.systems))
        
        return divs, revs, safes, ents, syss
    
    def bind_designation_to_idcode(self):
        #将识别出的图标实体信号位号绑定该图表的每一个实体
        v_bind = {}
        for _ in self.bounds.values():
            for bound in _:
                diagram_number = bound.diagram_number
                idcode, designation = bound.text.split('-')
                if idcode not in v_bind:
                    v_bind[idcode] = {}
                if diagram_number not in v_bind[idcode]:
                    v_bind[idcode][diagram_number] = {"Chinese": "", "English": ""}
                if designation is not None:
                    if re.search('[\u4e00-\u9fff]', designation):
                        v_bind[idcode][diagram_number]["Chinese"] += designation.strip()
                    else:
                        v_bind[idcode][diagram_number]["English"] += designation.strip()
        self.bounds = v_bind
        return v_bind
    
    
    def identify_accessories(self):
        #识别sensors, special_ios, actuators，这个函数的逻辑比较复杂，可以进一步refactor
        sensors = []
        special_ios = []
        actuators = []
        
        accessories = sorted([entity.parent for entity in self.entities if entity.tag and entity.text
                      and re.search("INSIDE", entity.tag)], key = lambda x: x.sorting_pos)
        for accessory in accessories:
            accessory.sys = accessory.sys or ''
            if len(accessory.children['INSIDE'].text) > 3:
                accessory.children['INSIDE'].text = accessory.sys + accessory.children['INSIDE'].text
                accessory.children['INSSYS'].relevance = accessory.children['INSIDE']
                actuators.append(accessory.children['INSSYS'])
            elif accessory.children['INSSYS'].text[-2:] in self.special_postfix:
                accessory.children['INSSYS'].text = accessory.sys + accessory.children['INSIDE'].text + accessory.children['INSSYS'].text
                special_ios.append(accessory.children['INSSYS'])
            else:
                accessory.children['INSSYS'].text = accessory.sys + accessory.children['INSIDE'].text + accessory.children['INSSYS'].text
                sensors.append(accessory.children['INSSYS'])
        scattered = []
        for sys in self.systems:
            scattered.extend([entity for entity in self.entities if entity.text and 
                     re.search("[0-9]{3}[A-Za-z]{2,3}$", entity.text)])
        unmatched = []
        for entity in scattered:
            if entity.text[-2:] in self.special_postfix:
                special_ios.append(entity)
                if len(entity.text) < 6:
                    entity.text = entity.sys + entity.text
                if len(entity.text) == 8:
                    entity.text = 'X' + entity.text
            else:
                if entity.text[-2:] in self.actuators_postfix:
                    unmatched.append(entity)
        unmatched_regions = self.get_target_regions(unmatched, 50, 50, 50, 50)
        arrows, varrows = self.get_arrows(unmatched_regions)
        arrow_regions = self.get_target_regions(arrows, 10, 10, 58, 58)
        # varrow_regions = self.get_target_regions(varrows, 10,10,10,10)
        non_chinese = self.not_contain(r'[\u3400-\u9FFF]+') 
        for entity in non_chinese:
            if entity.type == "MTEXT" and entity.text and len(entity.text) < 3 and entity.text in self.extension:
                if entity.color == 4 or entity.text == "P" or entity.text == "C":
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
                        if entity.relevance.text != entity.relevance.diagram_entity:
                            entity.relevance.exception.add('信号位号idcode', (entity.relevance.diagram_entity, entity.relevance.text) )
                            entity.relevance.text = entity.relevance.diagram_entity
                        entity.relevance.accessories.append(entity)
                        actuators.append(entity)
                        continue
                    # relevance = self.check_within(entity.pos, varrow_regions,overload=True)
                    # if relevance:
                    #     entity.associates.extend(relevance)
                    #     entity.relevance = entity.relevance_from_associates().relevance
                    #     entity.relevance.accessories.append(entity)
                    #     actuators.append(entity)
        special_ios = self.special_ios_duplicates_processing(special_ios)
        self.accessories["sensors"] = sensors
        self.accessories["special_ios"] = special_ios
        self.accessories["actuators"] = actuators
        return sensors, special_ios, actuators
        
    def get_target_regions(self, entities, up, down, left, right):
        #根据实体的位置，生成一个矩形区域，用于后续的判断
        res = []
        for entity in entities:
            region = self.generate_rectangle(entity.pos, up, down, left, right)
            res.append((region, entity))
        return res
        
    def generate_rectangle(self, coordinates, up, down, left, right):
        #根据实体的位置，生成一个矩形区域，用于后续的判断
        x, y = coordinates[0], coordinates[1]
        bounds = (x-left, x+right, y-down, y+up)
        return bounds
      
    def get_arrows(self, regions):
        #识别箭头
        arrows = []
        varrows = []
        for entity in self.entities:
            match entity.type:
                case "LWPOLYLINE":
                    if 2.9<entity.length< 5:                                
                        relevance = self.check_within(entity.pos, regions)
                        if relevance is not None:
                            entity.relevance = relevance
                            relevance.accessories.append(entity)
                            varrows.append(entity)
                    if 1.2< entity.area <1.8:
                        relevance = self.check_within(entity.pos, regions)
                        if relevance is not None:
                            entity.relevance = relevance
                            relevance.accessories.append(entity)
                            arrows.append(entity)
        return arrows, varrows
    
    def check_within(self, coordinates, rectangles, overload = False):
        #判断一个坐标是否在一个矩形区域内，并返回矩形区域对应的实体
        x,y = coordinates[0:2]
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
        #将实体信息导出到excel文件，还未完成！！！
        dataframe_map = {}
        for entity in self.entities:
            dataframe_map[entity.handle] = entity.entity_to_entry()
        df = pd.DataFrame.from_dict(dataframe_map, 
                                    orient = "index",
                                    columns = ["Handle", "Type", "Layer", 
                                               "Tag", "Text", "Diagram Number",
                                               "Relevance", "Position", "Sorting Position"])
        
        df.to_excel(filepath)
        
    def accessories_to_map(self, sensors, special_ios, actuators):
        #将识别出的sensors, special_ios, actuators导出到pandas dataframe以进行后续比对
        sensors_dict, special_ios_dict, actuators_dict = {}, {}, {}
        for sensor in sensors:
            sensor.dict = sensor.to_dict(self.bounds)
            sensors_dict[(sensor['信号位号idcode'][1:], sensor['图号diagram number'])] = sensor
        for special_io in special_ios:
            special_io.dict = special_io.to_dict(self.bounds)
            special_ios_dict[(special_io['信号位号idcode'][1:], special_io['图号diagram number'])] = special_io
        for actuator in actuators:
            actuator.dict = actuator.to_dict(self.bounds)
            actuators_dict[(actuator['信号位号idcode'][1:], actuator['扩展码extensioncode'], actuator['图号diagram number'])] = actuator
        return (sensors_dict, special_ios_dict, actuators_dict)
    #need to be modifies to only take copiess
    def export_accessories_to_excel(self, sensors, special_ios, actuators, filepath):
        #将识别出的sensors, special_ios, actuators导出到excel文件
        def extract_text_sort(accessories):
            accessories.sort(key = lambda x: x.diagram_number)
            for asso in accessories:
                asso.designation = [x.text for x in asso.trim_designation()]
                asso.designation.sort(key=lambda x: x, reverse = True)
        extract_text_sort(sensors)
        extract_text_sort(special_ios)
        extract_text_sort(actuators)
        sensors_dict = []
        special_ios_dict = []
        actuators_dict = []
        for sensor in sensors:
            sensors_dict.append(sensor.to_dict())
        for special_io in special_ios:
            special_ios_dict.append(special_io.to_dict())
        for actuator in actuators:
            actuators_dict.append(actuator.to_dict())
        df1 = pd.DataFrame.from_dict(sensors_dict)
        df2 = pd.DataFrame.from_dict(special_ios_dict)
        df3 = pd.DataFrame.from_dict(actuators_dict)
        for df in (df1, df2, df3):
            df.insert(0, "序号serial number", df.index+1)
        output = open(filepath, "w")
        writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
        df1.to_excel(writer, sheet_name='SENSOR_IO', index=False)
        df3.to_excel(writer, sheet_name='ACTUATOR_IO', index=False)
        df2.to_excel(writer, sheet_name='SPECIAL_IO', index=False)
        writer.close()
        output.close()
    
    def special_ios_duplicates_processing(self, special_ios):
        #处理special_ios中的重复项
        text_dict = {}
        for spe in special_ios:
            if spe.text not in text_dict:
                text_dict[spe.text] = [spe]
            else:
                text_dict[spe.text].append(spe)
        duplicates = [item for item in text_dict.values() if len(item) > 1]
        
        for duo in duplicates:
            duo.sort(key=lambda x: x.pos, reverse = True)
            for spe in duo[1:]:
                special_ios.remove(spe)
        return special_ios
    
    def search_by_handle(self, handle):
        #根据handle查找实体
        for entity in self.entities:
            if entity.handle == handle:
                return entity
        return None
    
    def contain(self, words):
        #根据关键词查找实体
        return contain(words, self.entities)
    
    def not_contain(self, words):
        #根据关键词查找实体
        return not_contain(words, self.entities)
    
    def retrive_designation(self):
        #识别每个实体相关的图例列表
        accessories = []
        for items in self.accessories.values():
            accessories.extend(items)
        regions = self.get_target_regions(accessories, 18, 18,77,77)
        chinese = self.contain(r'[\u3400-\u9FFF]+')
        chinese_illustration = []
        for chi in chinese:
            relevance = self.check_within(chi.pos, regions, overload=True)
            if relevance is not None and (chi.type == "MTEXT" or chi.type == "TEXT") and (not chi in self.bounds):
                chi.associates.extend(relevance)
                chi.relevance = chi.relevance_from_associates(True)
                chinese_illustration.append(chi)
        for chiilu in chinese_illustration:
            #can be modified to take a regex pattern as paramter
            if "过程" in chiilu.text:
                chinese_illustration.remove(chiilu)
                continue
            for _ in chiilu.associates:
                _.designation.append(chiilu)
            # chiilu.relevance.designation.append(chiilu)
            
        relevant_chinese_illustration = []
        for acco in accessories:
            if acco.designation:
                relevant_chinese_illustration.append(acco.get_closest_designation())
        relevant_chinese_illustration = set(relevant_chinese_illustration)
        illu_regions = self.get_target_regions(relevant_chinese_illustration, 0,20,10,10)
        non_chinese = self.not_contain(r'[\u3400-\u9FFF]+')
        english_illustration = []
        directs = []
        for entity in non_chinese:
            if entity.type == "MTEXT" and (not entity in self.bounds):
                relevance = self.check_within(entity.pos, illu_regions,overload=True)
                if relevance is not None:
                    entity.associates.extend(relevance)
                    english_illustration.append(entity)
                direct = self.check_within(entity.pos, regions, overload=True)
                if direct is not None:
                    entity.directs.extend(direct)
                    directs.append(entity)
        for engilu in directs:
            for _ in engilu.directs:
                if _.handle != engilu.handle:
                    _.designation.append(engilu)
        for engilu in set(english_illustration):
            engilu.text = engilu.text.replace('\n', '')   
            if engilu.associates:
                for asso in engilu.associates:
                    if asso.handle == engilu.handle:
                        engilu.associates.remove(asso)
                engilu.relevance = engilu.relevance_from_associates(True)
                if engilu.relevance.relevance == engilu:
                    engilu.relevance = None
                    continue
                engilu.relevance.relevance.designation.append(engilu)
                for _ in engilu.relevance.associates:
                    if engilu.handle != _.handle:
                        _.designation.append(engilu)
        illustration = list(relevant_chinese_illustration) + english_illustration
        
        return illustration
        
    def trim_designation(self):
        #去除irrelevant的图例
        for entity in self.entities:
            entity.designation = [x for x in entity.trim_designation()]
            entity.designation.sort(key = lambda x: x.text, reverse = True)
        
        
        