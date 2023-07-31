import re
import ezdxf
import math
from ezdxf.entities import *
class CADEntity:
    def __init__(self, entity, params):
        self.entity = entity
        self.params = params
        #插入点一般是图形的四个角，所以需要set_pos计算中心点
        self.type = entity.dxftype()
        #图形几何中心
        self.pos = None
        self.relevance = None
        #图号
        self.diagram_number = None  
        self.sorting_pos = None
        self.parent = None
        self.associates = []
        self.designation = []
        self.accessories = []
        #版本
        self.rev = None
        #安全分级
        self.safety_class = None
        #一些特殊的CAD图包含单一实体的信号，这些信号的信号位号是一样的，所以需要一个变量来记录
        self.diagram_entity = None
        self.init_attribs()
        
             
    def set_pos(self):
        match self.type:
            case "TEXT":
                self.pos = self.align_point or self.insert
            case "CIRCLE":
                self.insert = self.center
                self.pos = self.center
                self.area = math.pi * self.radius**2
            case "LINE":
                self.insert = self.start
                self.pos = self.get_center([self.start, self.end])
                self.length = ezdxf.math.distance(self.start, self.end)
            case "ARC":
                self.pos = self.center
                self.insert = self.center
            case "LWPOLYLINE":
                with self.entity.points('xy') as points:
                    self.area = ezdxf.math.area(points)
                    self.pos = self.get_center(points)
                self.insert = self.pos
            case "INSERT":
                self.pos = self.insert
            case "MTEXT":
                self.pos = self.attachment_to_center()
                # return self.attachment_to_center()
            case "ATTDEF":
                self.pos = self.insert
                # return self.insert
            case "ATTRIB":
                self.pos = self.insert
                # return self.insert
            case "HATCH":
                self.pos = self.entity.seeds[0]
                self.insert = self.entity.seeds[0]
                paths = self.entity.paths.external_paths()
                vertices = []
                if paths:
                    path = next(iter(paths))
                if (path.type == BoundaryPathType.EDGE) and (path.edges[0].type == EdgeType.LINE):
                    for edge in path.edges:
                        vertices.append(edge.start)
                        vertices.append(edge.end) 
                self.area = ezdxf.math.area(vertices)
        self.insert = tuple(self.insert)
        self.pos = tuple(self.pos)       
                
    def attachment_to_center(self):
        """MText.dxf.attachment_point	Value
            MTEXT_TOP_LEFT	1
            MTEXT_TOP_CENTER	2
            MTEXT_TOP_RIGHT	3
            MTEXT_MIDDLE_LEFT	4
            MTEXT_MIDDLE_CENTER	5
            MTEXT_MIDDLE_RIGHT	6
            MTEXT_BOTTOM_LEFT	7
            MTEXT_BOTTOM_CENTER	8
            MTEXT_BOTTOM_RIGHT	9"""
        match self.attachment: 
            case 1:
                return (self.insert[0] + self.width/2, self.insert[1] - self.height/2)
            case 2:
                return (self.insert[0], self.insert[1] - self.height/2)
            case 3:
                return (self.insert[0] - self.width/2, self.insert[1] - self.height/2)
            case 4:
                return (self.insert[0] + self.width/2, self.insert[1])
            case 5:
                return (self.insert[0], self.insert[1])
            case 6:
                return (self.insert[0] - self.width/2, self.insert[1])
            case 7: 
                return (self.insert[0] + self.width/2, self.insert[1] + self.height/2)
            case 8:
                return (self.insert[0], self.insert[1] + self.height/2)
            case 9:
                return (self.insert[0] - self.width/2, self.insert[1] + self.height/2)

                
            
    def align_to_center(self):
        """dxf.valign
        Vertical alignment flag as int value, use the set_placement() and get_align_enum() methods to handle text alignment, the default value is 0.

        0	Baseline
        1	Bottom
        2	Middle
        3	Top"""
        """dxf.halign
        Horizontal alignment flag as int value, use the set_placement() and get_align_enum() methods to handle text alignment, the default value is 0.

        0	Left
        2	Right
        3	Aligned (if vertical alignment = 0)
        4	Middle (if vertical alignment = 0)
        5	Fit (if vertical alignment = 0)"""
        return
        
    
    def init_attribs(self):
        try:
            self.height = self.params['char_height']
        except KeyError:
            try:
                self.height = self.params['height']
            except KeyError:
                self.height = None
        try:
            self.attachment = self.params['attachment_point']
        except KeyError:
            self.attachment = None
        self.set_pos()
        
    
    def __getattribute__(self, __name: str):
        """Get Attribute from Params"""
        try:
            return object.__getattribute__(self, __name)
        except AttributeError:
            try:
                return self.params.get(__name)
            except KeyError:
                return None
    
    def relevance_from_associates(self):
        """Find the most relevant(In terms of distance) associate from the associates list"""
        if len(self.associates) == 1:
            self.relevance = self.associates[0]
            return self.associates[0]
        elif len(self.associates) > 1:
            self.associates.sort(key=lambda x: self.get_distance(x))
            self.relevance = self.associates[0]
            return self.associates[0]
        
    def get_closest_designation(self):
        """Get the most relevant(In terms of distance) designation from the designation list"""
        if len(self.designation) > 1:
            self.designation.sort(key=lambda x: self.get_distance(x))
            return self.designation[0]
        else:
            return self.designation[0]
    
    def trim_designation(self):
        """Remove Irrelevant Designation and Sort the Designation by Distance"""
        for designation in self.designation:
            designation.text = designation.text.replace('\\P', '')
        if len(self.designation) > 3:
            self.designation.sort(key=lambda x: self.get_distance(x))
            chinese = [x for x in self.designation if re.search('[\u4e00-\u9fff]', x.text)]
            english = [x for x in self.designation if not re.search('[\u4e00-\u9fff]', x.text)]
            self.designation = chinese[:1] + english[:1]
        else:
            self.designation.sort(key=lambda x: self.get_distance(x))
        return self.designation
    
    def designation_to_string(self, bounds):
        """Convert Designation List to Output String"""
        if self.designation is None:
            return '-'
        des_list = self.trim_designation()
        id_code = (self.relevance and self.relevance.text) or self.text
        chinese_addition, english_addition = None, None
        if id_code in bounds:
            if self.diagram_number in bounds[id_code]:
                chinese_addition = bounds[id_code][self.diagram_number]['Chinese']
                english_addition = bounds[id_code][self.diagram_number]['English']
        chinese = [x.text for x in des_list if re.search('[\u4e00-\u9fff]', x.text)]
        english = [x.text for x in des_list if not re.search('[\u4e00-\u9fff]', x.text)]
        if not (chinese or english):
            return '无描述' 
        chinese.sort(key=lambda x: len(x), reverse=True)
        english.sort()
        chinese, english = ''.join(chinese), ''.join(english)
        if chinese_addition and (not re.search(chinese_addition, chinese)):
            chinese = chinese_addition + chinese
        if english_addition and (not re.search(english_addition, english)):
            english = english_addition + english
        return ((chinese or english) and chinese +'\n' + english) or ''
        
    def to_dict(self, bounds):
        """Convert Entity to Dictionary Entry for Dataframe Output"""
        return {'信号位号idcode': (self.relevance and self.relevance.text) or self.text, 
                '扩展码extensioncode': (self.relevance and self.text) or '无扩展码',
                "信号说明designation": self.designation_to_string(bounds),
                "图号diagram number": self.diagram_number,
                "版本rev.": self.rev,
                "备注remark": "-",
                "安全分级/分组Safetyclass/division": self.safety_class or '无安全分级/分组',
                }
        
    def get_center(self, coordinates):
        x, y = (coordinates[0][0]+coordinates[1][0]) /2 , (coordinates[0][1]+ coordinates[1][1]) /2
        return (x, y)
        
    def __repr__(self):
        return f"{self.handle} {self.type} {self.text}"
    
    def set_sorting_pos(self):
        self.sorting_pos =  (int(round(float(self.pos[0])/5)*5), round(float(self.pos[1])))
    
    def get_distance(self, other= None):
        return ((self.pos[0]-other.pos[0])**2 + (self.pos[1]-other.pos[1])**2)**0.5
    
    def get_distance_by_c(self, coordinates):
        return ((self.pos[0]-coordinates[0])**2 + (self.pos[1]-coordinates[1])**2)**0.5
    
