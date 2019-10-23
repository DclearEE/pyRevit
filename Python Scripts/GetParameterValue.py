#Copyright (c) 2014, Nathan Miller
#The Proving Ground http://theprovingground.org

# Default imports
import clr

# Import RevitAPI
clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *

# Import DocumentManager and TransactionManager
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

# Import ToDSType(bool) extension method
clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)

#The input to this node will be stored in the IN[0] variable.

doc =  DocumentManager.Instance.CurrentDBDocument
app =  DocumentManager.Instance.CurrentUIApplication.Application

elements = UnwrapElement(IN[0])
parameter = IN[1]

values = []
if hasattr(elements, "__iter__"):
	output = []
	for elem in elements:
		if hasattr(elem, "__iter__"):
			vals = []
			for e in elem:
				for p in elem.Parameters:
					if p.Definition.Name == parameter:		
						parm = p.AsValueString()
						if (parm is None):
							parm = p.AsString()
				vals.append(parm)
			values.append(vals)
		else:
			for p in elem.Parameters:
				if p.Definition.Name == parameter:		
					parm = p.AsValueString()
					if (parm is None):
							parm = p.AsString()
			values.append(parm)
	output.append(values)
else:
	parm = 	elements.Parameter[parameter].AsValueString()
	output = parm



#Assign your output to the OUT variable
OUT = output