class CADEntity:
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
             
    def set_pos(self):
        if self.pos is not None:
            self.pos = self.pos[0:2]
        elif self.bounding_box is not None:
            self.pos = self.get_center(tuple(self.bounding_box)[0:2])
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
        return [self.text, self.diagram_number]
    def actuators_to_entry(self):
        return [self.relevance.text, self.text, self.diagram_number]