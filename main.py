from utility.utility import *
from cad_entity import CADEntity
from cad_handler import CADHandler

#CAD Info Extration
acad = CADHandler("entity_list.pkl")
acad.set_params()
at,b,c = acad.identify_accessories()
illustration = acad.retrive_designation()

#Compared Data Unification
control, ref = read_sheet('control.xlsx')
extracted = acad.accessories_to_df(at, b,c)

for c, e in zip(control, extracted):
    unify(c)
    unify(e)
sensors, specials, actuators = extracted
sensors_control, specials_control, actuators_control = control

#Mark and Write Diff to Master File
output_diff('control.xlsx', ref, (sensors, specials, 
                                actuators, sensors_control, 
                                specials_control, actuators_control))
