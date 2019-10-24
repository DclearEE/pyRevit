import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('IronPython.Wpf')

#pylint: disable=E0401,W0703,C0103
# noinspection PyUnresolvedReferences
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
# noinspection PyUnresolvedReferences
from Autodesk.Revit.DB import Transaction
# noinspection PyUnresolvedReferences
from Autodesk.Revit.Exceptions import InvalidOperationException
from Autodesk.Revit.UI.Selection import ISelectionFilter
import rpw
import wpf
from System import Windows
from pyrevit import revit
from pyrevit import forms
from pyrevit.forms import WPFWindow
from pyrevit import UI
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
        
        try:
            selection = revit.get_selection()
            msfilter = CustomSelectionFilter("Electrical Equipment")
            selection_list = revit.pick_rectangle(pick_filter=msfilter)
            filtered_list = []
            for element in selection_list:
                filtered_list.append(element.Id)
            selection.set_to(filtered_list)
            print(msfilter)
        except Exception as err:
            logger.debug(err)

        UI.TaskDialog.Show("Hello World", "Hello{}")



    def RUN_Click(self, sender, args):
        findPanel = self.PanelFind.Text
        replacePanel = self.PanelReplace.Text
        findCircuit = self.CircuitFind.Text
        replaceCircuit = self.CircuitReplace.Text
        self.Close()
        
MyWindow().ShowDialog()

class CustomSelectionFilter(ISelectionFilter):
    def __init__(self):
        self.category = catname

    # standard API override function
    def AllowElement(self, element):
        if self.category in element.Category.Name:
            return True
        else:
            return False

    # standard API override function
    def AllowReference(self, refer, point): #pylint: disable=W0613
        return False

    

def pickbycategory(catname):
    try:
        selection = revit.get_selection()
        msfilter = CustomSelectionFilter(catname)
        selection_list = revit.pick_rectangle(pick_filter=msfilter)

        filtered_list = []
        for element in selection_list:
            filtered_list.append(element.Id)
            
        selection.set_to(filtered_list)
        print(msfilter)
    except Exception as err:
        logger.debug(err)



if __shiftclick__:  #pylint: disable=E0602
    options = sorted([x.Name for x in revit.doc.Settings.Categories])
else:
    options = sorted(['Area',
                      'Area Boundary',
                      'Column',
                      'Dimension',
                      'Door',
                      'Floor',
                      'Framing',
                      'Furniture',
                      'Grid',
                      'Rooms',
                      'Room Tag',
                      'Truss',
                      'Wall',
                      'Window',
                      'Ceiling',
                      'Section Box',
                      'Elevation Mark',
                      'Parking'])

selected_switch = \
    forms.CommandSwitchWindow.show(options,
                                   message='Pick only elements of type:')

if selected_switch:
    pickbycategory(selected_switch)

    
