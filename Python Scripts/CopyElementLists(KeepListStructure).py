import clr
from System.Collections.Generic import *
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
import Autodesk

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

doc = DocumentManager.Instance.CurrentDBDocument
input = UnwrapElement(IN[0])
xyz = UnwrapElement(IN[1]).ToXyz()


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


TransactionManager.Instance.EnsureInTransaction(doc)
collection = []
for elec in ElectComponents:
	elementList = []
	for e in elec:
		newitemList = []
		newitems = ElementTransformUtils.CopyElements(doc,e,doc,Transform.CreateTranslation(xyz),None)
		for n in newitems:	
			newitemList.append(doc.GetElement(n).ToDSType(False))
		elementList.append(newitemList)
	collection.append(elementList)
		

	


OUT = collection