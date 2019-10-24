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

    def RUN_Click(self, sender, args):
        findPanel = self.PanelFind.Text
        replacePanel = self.PanelReplace.Text
        findCircuit = self.CircuitFind.Text
        replaceCircuit = self.CircuitReplace.Text
        self.Close()

        
MyWindow().ShowDialog()

""" CategoryOption = namedtuple('CategoryOption', ['name', 'revit_cat'])

category_opt = DB.BuiltInCategory.OST_Grids """

""" class PickByCategorySelectionFilter(UI.Selection.ISelectionFilter):
    def __init__(self, category_opt):
        self.category_opt = category_opt

    # standard API override function
    def AllowElement(self, element):
        if element.Category \
                and self.category_opt.revit_cat.Id == element.Category.Id:
            return True
        else:
            return False

    # standard API override function
    def AllowReference(self, refer, point):  # pylint: disable=W0613
        return False


def pick_by_category(category_opt):
    try:
        selection = revit.get_selection()
        msfilter = PickByCategorySelectionFilter(category_opt)
        selection_list = revit.pick_rectangle(pick_filter=msfilter)
        filtered_list = []
        for element in selection_list:
            filtered_list.append(element.Id)
        selection.set_to(filtered_list)
    except Exception as err:
        logger.debug(err)



source_categories = \
    [revit.query.get_category(x) for x in FREQUENTLY_SELECTED_CATEGORIES]
if __shiftclick__:  # pylint: disable=E0602
    source_categories = revit.doc.Settings.Categories

# cleanup source categories
source_categories = filter(None, source_categories)
category_opts = \
    [CategoryOption(name=x.Name, revit_cat=x) for x in source_categories]
selected_category = \
    forms.CommandSwitchWindow.show(
        sorted([x.name for x in category_opts]),
        message='Pick only elements of type:'
    )

if selected_category:
    selected_category_opt = \
        next(x for x in category_opts if x.name == selected_category)
    logger.debug(selected_category_opt) 
    pick_by_category(selected_category_opt) """
