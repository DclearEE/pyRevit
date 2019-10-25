import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('IronPython.Wpf')

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
import Autodesk

clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

#pylint: disable=E0401,W0703,C0103
# noinspection PyUnresolvedReferences
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
# noinspection PyUnresolvedReferences
from Autodesk.Revit.DB import Transaction
# noinspection PyUnresolvedReferences
from Autodesk.Revit.Exceptions import InvalidOperationException
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.UI.Selection import ISelectionFilter
import rpw
import wpf
from System import Windows
from itertools import imap

from pyrevit import revit, UI, DB
from pyrevit import forms
from pyrevit.forms import WPFWindow
from pyrevit import script
xamlfile = script.get_bundle_file('ui.xaml')

doc = rpw.revit.doc
uidoc = rpw.revit.uidoc

__doc__ = "Select Elements of a Cateogry with Parameter Filter"
__title__ = "Pick'em"
__author__ = "Daniel Cleary"


logger = script.get_logger()



class MyWindow(Windows.Window):

    def __init__(self):
        wpf.LoadComponent(self, xamlfile)

    def SelectPanel_Click(self, sender, args):

        def tolist(obj1):
            if hasattr(obj1,"__iter__"): return obj1
            else: return [obj1]

        selection = revit.get_selection()
        self.Close()
        class MassSelectionFilter(UI.Selection.ISelectionFilter):
            # standard API override function

            def AllowElement(self, element):
                if not element.ViewSpecific:
                    return True
                else:
                    return False

            # standard API override function
            def AllowReference(self, refer, point):
                return False


        try:
            msfilter = MassSelectionFilter()
            selection_list = revit.pick_rectangle(pick_filter=msfilter)
            for el in selection_list:
                filtered_list.append(el.Id)
            selection.set_to(filtered_list)
            revit.uidoc.RefreshActiveView()
        except Exception:
            pass
        
        in1 = []
        out1 = []
        filter = map(str.lower, imap(str, tolist("Electrical Equipment") ) ) 
        for e in selection_list:
            c1 = e.Category
            if c1 is None:
                nulls.append(e)
            else:
                n1 = c1.Name.lower()
                if any(f in n1 for f in filter):
                    in1.append(e)
                else:
                    out1.append(e)
        for x in in1:
            print(x.Name)
        MyWindow().ShowDialog()


    def Origin_Click(self, sender, args):
        self.Close()
        origin = revit.pick_element()
        print(origin.Name)
        MyWindow().ShowDialog()


    def End_Click(self, sender, args):
        self.Close()
        dest = revit.pick_element()
        print(dest.Name)
        MyWindow().ShowDialog()

    def RUN_Click(self, sender, args):
        findPanel = self.PanelFind.Text
        replacePanel = self.PanelReplace.Text
        findCircuit = self.CircuitFind.Text
        replaceCircuit = self.CircuitReplace.Text
        self.Close()

        
MyWindow().ShowDialog()


