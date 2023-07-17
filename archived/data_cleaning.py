import pandas as pd
import xlsxwriter

df = pd.read_excel("output.xlsx")
additions = pd.read_excel("scattered.xlsx")

filtered = df[df.Text.str.match('\d+:\ 1')].index
entities = []
for row in filtered:
    if df['Graph_id'][row][9] == '1':
        diagram_number = 'FD-'+df['Text'][row].split(':')[0]
    if df['Graph_id'][row][9] == '2':
        diagram_number = 'SAMA-'+df['Text'][row].split(':')[0]
    df['diagram\nnumber'].replace(df['diagram\nnumber'][row], diagram_number, inplace=True)

df['InsertionPoint'] = df['InsertionPoint'].apply(lambda x: tuple(x[1:].split(',')))
df['region'] = df['InsertionPoint'].apply(lambda x: (int(round(float(x[0])/5)*5), round(float(x[1]))))
df.sort_values(by=['diagram\nnumber', 'region', 'InsertionPoint'], inplace=True)

#find Train
# Train is placed 40 px right to SYSTEM ID

sys = df[df['Text'].str.match("INSSYS*")].loc[:,['Text', 'diagram\nnumber', 'InsertionPoint']].itertuples()
ide = df[df['Text'].str.match("INSIDE*")].loc[:,['Text', 'diagram\nnumber', 'InsertionPoint']].itertuples()
scattered = df[df['Text'].str.match("YTFS[^\s][^-]+$")].loc[:,['Text', 'diagram\nnumber', 'InsertionPoint']].itertuples()


special_ios = []
sensors = []
actuators = []
unmatched = []
for s, i in zip(sys, ide):
    entity = "YTFS" + i[1][8:] + s[1][8:]
    if s[1][8:] == "SY":
        special_ios.append((entity, s[2]))
    elif len(i[1][8:]) == 3:
        sensors.append((entity, s[2]))
    else:
        entity = "YTFS" + i[1][8:] 
        ex_code = s[1][8:]
        actuators.append((entity,ex_code, s[2]))


for row in scattered:
    index, entity, id, pos = row
    if entity[7:] == "SY":
        special_ios.append((entity,id))
    else:
        unmatched.append((entity, id, pos))

scatter_pos = [] 

for row in additions.itertuples():
    entity = unmatched[row[1]]
    actuators.append((entity[0], row[2], entity[1]))

   
for entity in unmatched:
    pos = entity[2]
    bounds = (float(pos[0])-20, float(pos[0])+20, float(pos[1])-20, float(pos[1])+20)
    scatter_pos.append(bounds)
       
special_ios = set(special_ios)
special_ios = pd.DataFrame(special_ios)
sensors = pd.DataFrame(sensors)
actuators = pd.DataFrame(actuators)


output = open("first_trial.xlsx", "w")

writer = pd.ExcelWriter("first_trial.xlsx", engine='xlsxwriter')

sensors.to_excel(writer, sheet_name='SENSOR_IO', index=False)
actuators.to_excel(writer, sheet_name='ACTUATOR_IO', index=False)
special_ios.to_excel(writer, sheet_name='SPECIAL_IO', index=False)

writer.close()
output.close()
actuators.to_excel("actuators.xlsx", index=False)
sensors.to_excel("sensors.xlsx", index=False)
special_ios.to_excel("special_ios.xlsx", index=False)
# actuators.to_excel("entities.xlsx", sheet_name='actuators', index=False)
# sensors.to_excel("entities.xlsx", sheet_name='sensors', index=False)
# special_ios.to_excel("entities.xlsx",sheet_name='spcial_ios', index=False)
df.to_excel("test1.xlsx")    