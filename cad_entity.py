import re
class CADEntity:
    class handle:
        def __init__(self, handle, entity):
            self.handle = handle
            self.entity = entity
    def __init__(self, param_dict):
        self.params = param_dict
        self.handle = param_dict['Handle']
        self.pos = param_dict['InsertionPoint']
        self.layer = param_dict['Layer']
        self.type = param_dict['ObjectName']
        self.area = param_dict['Area']
        self.length = param_dict['Length']
        self.height = param_dict['height']
        self.text = param_dict['TextString']
        self.tag = param_dict['TagString'] 
        self.color = param_dict['color']
        self.bounding_box = param_dict['GetBoundingBox']
        self.coordinates = param_dict['coordinates']
        self.relevance = None
        self.diagram_number = None  
        self.sorting_pos = None
        self.parent = None
        self.associates = []
        self.designation = []
        self.accessories = []
        self.rev = None
             
    def set_pos(self):
        if self.bounding_box is not None:
            self.pos = self.get_center(tuple(self.bounding_box)[0:2])
        elif self.pos is not None:
            self.pos = self.pos[0:2]
        elif self.coordinates is not None:
            self.pos = self.get_center(self.coordinates[0:2])
        return self.pos
    
    def get_center(self, coordinates):
        x, y = (coordinates[0][0]+coordinates[1][0]) /2 , (coordinates[0][1]+ coordinates[1][1]) /2
        return (x, y)
        
    def __repr__(self):
        return f"{self.diagram_number} {self.type} {self.text}"
    
    def set_sorting_pos(self):
        self.sorting_pos =  (int(round(float(self.pos[0])/5)*5), round(float(self.pos[1])))
    
    def entity_to_entry(self):
        return [self.diagram_number, self.type, self.text, self.pos, self.layer, self.area, self.length, self.height, self.tag, self.color, self.bounding_box, self.coordinates, self.relevance, self.sorting_pos, self.parent]
    
    def accessorie_to_entry(self):
        return [self.text, self.diagram_number, self.designation, self.rev]
   
    def actuators_to_entry(self):
        return [self.relevance.text, self.text, self.diagram_number, self.designation, self.rev]
    
    def get_distance(self, other= None, coordinates = False):
        return ((self.pos[0]-other.pos[0])**2 + (self.pos[1]-other.pos[1])**2)**0.5
    def get_distance_by_c(self, coordinates):
        return ((self.pos[0]-coordinates[0])**2 + (self.pos[1]-coordinates[1])**2)**0.5
    def relevance_from_associates(self):
        if len(self.associates) == 1:
            self.relevance = self.associates[0]
            return self.associates[0]
        elif len(self.associates) > 1:
            self.associates.sort(key=lambda x: self.get_distance(x))
            self.relevance = self.associates[0]
            return self.associates[0]
        
    def get_closest_designation(self):
        if len(self.designation) > 1:
            self.designation.sort(key=lambda x: self.get_distance(x))
            return self.designation[0]
        else:
            return self.designation[0]
    
    def trim_designation(self):
        if len(self.designation) > 3:
            self.designation = [self.get_closest_designation()]
        else:
            self.designation.sort(key=lambda x: self.get_distance(x))
        return self.designation
    
    # def to_dict(self):
    #     return {'Handle': self.handle, 'InsertionPoint': self.pos, 'Layer': self.layer, 'ObjectName': self.type, 'Area': self.area, 'Length': self.length, 'height': self.height, 'TextString': self.text, 'TagString': self.tag, 'color': self.color, 'GetBoundingBox': self.bounding_box, 'coordinates': self.coordinates}
    def designation_to_string(self):
        if self.designation is None:
            return '-'
        des_list = self.trim_designation()
        chinese = [x.text for x in des_list if re.search('[\u4e00-\u9fff]', x.text)]
        english = [x.text for x in des_list if not re.search('[\u4e00-\u9fff]', x.text)]
        chinese.sort(key=lambda x: len(x), reverse=True)
        english.sort()
        return ''.join(chinese) +'\n' + ''.join(english)
        
        
    def to_dict(self):
        return {'信号位号idcode': (self.relevance and self.relevance.text) or self.text, 
                '扩展码extensioncode': (self.relevance and self.text) or '-',
                "信号说明designation": self.designation_to_string(),
                "图号diagram number": self.diagram_number,
                "版本rev.": self.rev,
                "备注remark": "-",
                }
    