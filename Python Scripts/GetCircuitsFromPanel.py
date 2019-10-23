import clr

# Import Revit API + APIUI
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import *

OUT = []

def tolist(obj):
	if isinstance(obj, list):
		return UnwrapElement(obj)
	else:
		return [UnwrapElement(obj)]

panels = tolist(IN[0])
circuitsInPanel = []
circuits = []
panel = []
try:
	for x in panels:
		circuits = []
		for y in x.MEPModel.ElectricalSystems:
			circuits.append(y)
		circuitsInPanel.append(circuits)
		panel.append(x)
except: 
	pass
	
OUT = circuitsInPanel, panel

#Assign your output to the OUT variable.
if len(OUT) == 1:
	OUT = OUT[0]
else:
	OUT = OUT 