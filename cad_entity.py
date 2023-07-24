import re
class CADEntity:
    def __init__(self, param_dict):
        self.params = param_dict
        self.handle = param_dict['Handle']
        #图形中心点
        self.pos = param_dict['InsertionPoint']
        #插入点一般是图形的四个角，所以需要set_pos计算中心点
        self.insertionpoint = param_dict['InsertionPoint']
        self.layer = param_dict['Layer']
        self.type = param_dict['ObjectName']
        self.area = param_dict['Area']
        self.length = param_dict['Length']
        self.height = param_dict['height']
        #一个CAD实体的文字内容可能是text也可能是tag
        self.text = param_dict['TextString']
        self.tag = param_dict['TagString'] 
        self.color = param_dict['color']
        #bounding_box是一个四元组，分别是左下角和右上角的坐标
        self.bounding_box = param_dict['GetBoundingBox']
        #coordinates是一个列表，每个元素是一个坐标元组
        self.coordinates = param_dict['coordinates']
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
             
    def set_pos(self):
        """Get Geometric Center of the Entity"""
        if self.bounding_box is not None:
            self.pos = self.get_center(tuple(self.bounding_box)[0:2])
        elif self.pos is not None:
            self.pos = self.pos[0:2]
        elif self.coordinates is not None:
            self.pos = self.get_center(self.coordinates[0:2])
        return self.pos
    
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
        return f"{self.diagram_number} {self.type} {self.text}"
    
    def set_sorting_pos(self):
        self.sorting_pos =  (int(round(float(self.pos[0])/5)*5), round(float(self.pos[1])))
    
    def get_distance(self, other= None):
        return ((self.pos[0]-other.pos[0])**2 + (self.pos[1]-other.pos[1])**2)**0.5
    
    def get_distance_by_c(self, coordinates):
        return ((self.pos[0]-coordinates[0])**2 + (self.pos[1]-coordinates[1])**2)**0.5
    
