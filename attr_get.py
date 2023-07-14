import re
file = open('module.txt', 'r')

p = re.compile("entity\.[^,\)\s\[\(\]\}\<]*")
lines = file.readlines()
print(p.match('entity.a'))
attrs = []
for line in lines:
    line = line.strip()
    attrs.extend(p.findall(line))
    
res = [attr.split('.')[1] for attr in attrs]
print(set(res))
