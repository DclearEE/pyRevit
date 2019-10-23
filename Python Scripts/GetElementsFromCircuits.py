import sys
#pyt_path = r"C:\Program Files (x86)\IronPython 2.7\DLLs"
#sys.path.append(pyt_path)
import clr

#clr.AddReference('IronPython.SQLite.dll')
pyt_path = r"C:\Program Files (x86)\IronPython 2.7\Lib"
sys.path.append(pyt_path)
#import sqlite3

# Import ToDSType(bool) extension method
clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.GeometryConversion)
clr.ImportExtensions(Revit.Elements)

# Import DocumentManager and TransactionManager
clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

doc = DocumentManager.Instance.CurrentDBDocument
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
app = DocumentManager.Instance.CurrentUIApplication.Application

# Import List ( ICollection(ElementId) = List[ElementId]() )
clr.AddReference('System')
from System.Collections.Generic import List

# Import Revit API + APIUI
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import *

OUT = []

def tolist(obj):
	if isinstance(obj, list):
		return UnwrapElement(obj)
	else:
		return [UnwrapElement(obj)]

# Start Transaction
# TransactionManager.Instance.EnsureInTransaction(doc)

# End Transaction
# TransactionManager.Instance.TransactionTaskDone()

systems = tolist(IN[0])
elements = []
	

if hasattr(systems[0],'__iter__'):
	newlist = []
	elements = []
	for x in systems:
		collection = []
		for y in x:
			newlist = []
			try:
				for z in y.Elements:
					newlist.append(z)
			except:
				if y.IsEmpty:
					newlist.append(y.Category.Name)
				else:
					newlist.append("Failed")
			collection.append(newlist)
		elements.append(collection)
	
else:
	for x in systems:
		sublist = []
		try:
			elementset = x.Elements
			for x in elementset:
				sublist.append(x)
		except:
			if x.IsEmpty:
				sublist.append(x.Category.Name)
			else:
				sublist.append("Failed")
		elements.append(sublist)

#Assign your output to the OUT variable.
OUT = elements
if len(OUT) == 1:
	OUT = OUT[0]