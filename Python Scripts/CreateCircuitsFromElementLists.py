import clr
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *

clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.Elements)

clr.AddReference('RevitAPI')
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import *
from Autodesk.Revit.DB import MEPSystem

clr.AddReference('DSCoreNodes')
import DSCore
from DSCore.List import *

import sys
pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)

# Import List ( ICollection(ElementId) = List[ElementId]() )
clr.AddReference("System")
from System.Collections.Generic import List 

#The inputs to this node will be stored as a list in the IN variables.
input = UnwrapElement(IN[0])

NewCircuit = []

elementList = []

ElectComponents = []
for panel in input:	
	list = []
	for sublist in panel:
		sub = []
		for item in sublist:
			sub.Add(item.Id)
		list.append(List[ElementId](sub))
	ElectComponents.append(list)



doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument

listNewCircuit = []

TransactionManager.Instance.EnsureInTransaction(doc)

for circuit in ElectComponents:
	newCircuit = []
	for e in circuit:
		circuit = ElectricalSystem.Create(doc, e, ElectricalSystemType.PowerCircuit)
		newCircuit.append(circuit)
	listNewCircuit.append(newCircuit)
	

TransactionManager.Instance.TransactionTaskDone()


OUT = listNewCircuit




