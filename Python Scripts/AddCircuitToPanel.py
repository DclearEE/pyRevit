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


circuit = UnwrapElement(IN[0])
panel = UnwrapElement(IN[1])

doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument

TransactionManager.Instance.EnsureInTransaction(doc)


for index, c in enumerate(circuit):
	for _ in c:
		try:
			_.SelectPanel(panel[index])
		except:
			pass
TransactionManager.Instance.TransactionTaskDone()

#Assign your output to the OUT variable.
OUT = []
OUT = circuit, panel