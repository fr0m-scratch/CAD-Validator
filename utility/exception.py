class Exception:
    def __init__(self):
        self.bool = False
        self.mismatch = {}
        
    def __bool__(self):
        return self.bool
    
    def add(self, type, item):
        self.bool = True
        key, value = item
        self.mismatch[key] = value
                