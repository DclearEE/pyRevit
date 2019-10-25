# region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

using Autodesk.Revit;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI.Selection;
using XYZ = Autodesk.Revit.DB.XYZ;

using Reference = Autodesk.Revit.DB.Reference;
using Exceptions = Autodesk.Revit.Exceptions;
using Creation = Autodesk.Revit.Creation;
using DialogResult = System.Windows.Forms.DialogResult;
# endregion Namespaces

namespace NBK.Spool
{
    /// <summary>
    /// Allows the selection of connectable MEP Elements
    /// </summary>
    internal class MEPConnectableSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element element)
        {
            if (element.Category.Name == "Air Terminals" ||
                element.Category.Name == "Cable Trays" ||
                element.Category.Name == "Cable Tray Fittings" ||
                element.Category.Name == "Conduits" ||
                element.Category.Name == "Conduit Fittings" ||
                element.Category.Name == "Ducts" ||
                element.Category.Name == "Duct Accessories" ||
                element.Category.Name == "Duct Fittings" ||
                element.Category.Name == "Flex Ducts" ||
                element.Category.Name == "Flex Pipes" ||
                element.Category.Name == "Mechanical Equipment" ||
                element.Category.Name == "Pipes" ||
                element.Category.Name == "Pipe Accessories" ||
                element.Category.Name == "Pipe Fittings" ||
                element.Category.Name == "Plumbing Fixtures" ||
                element.Category.Name == "Sprinklers")
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            return false;
        }
    }

    /// <summary>
    /// Allows the selection of connectable MEP Elements, preventing the selection of a previously selected element.
    /// </summary>
    internal class MEPConnectableNewSelectionFilter : ISelectionFilter
    {

        public ElementId PreviousElementID { get; set; }

        public bool AllowElement(Element e)
        {
            if ((e.Category.Name == "Air Terminals" ||
                e.Category.Name == "Cable Trays" ||
                e.Category.Name == "Cable Tray Fittings" ||
                e.Category.Name == "Conduits" ||
                e.Category.Name == "Conduit Fittings" ||
                e.Category.Name == "Ducts" ||
                e.Category.Name == "Duct Accessories" ||
                e.Category.Name == "Duct Fittings" ||
                e.Category.Name == "Flex Ducts" ||
                e.Category.Name == "Flex Pipes" ||
                e.Category.Name == "Mechanical Equipment" ||
                e.Category.Name == "Pipes" ||
                e.Category.Name == "Pipe Accessories" ||
                e.Category.Name == "Pipe Fittings" ||
                e.Category.Name == "Plumbing Fixtures" ||
                e.Category.Name == "Sprinklers")
                && e.Id != PreviousElementID)
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            return false;
        }
    }


    /// <summary>
    /// Allows the selection of cable tray, conduit, duct and pipe.
    /// </summary>
    internal class MEPCurveSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element e)
        {
            if (e.Category.Name == "CableTray"
                || e.Category.Name == "Conduits"
                || e.Category.Name == "Ducts"
                || e.Category.Name == "Pipes")
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference r, XYZ p)
        {
            return true;
        }
    }


    /// <summary>
    /// Allows the selection of cable tray, conduit, duct and pipe as well as any associated fittings.
    /// </summary>
    public class MEPCurveOrFittingSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element e)
        {
            if (e.Category.Name == "CableTray"
                || e.Category.Name == "CableTray Fittings"
                || e.Category.Name == "Conduits"
                || e.Category.Name == "Conduit Fittings"
                || e.Category.Name == "Ducts"
                || e.Category.Name == "Duct Fittings"
                || e.Category.Name == "Duct Accessories"
                || e.Category.Name == "Flex Ducts"
                || e.Category.Name == "Pipes"
                || e.Category.Name == "Pipe Fittings"
                || e.Category.Name == "Pipe Accessories"
                || e.Category.Name == "Flex Pipes")
            {
                return true;
            }
            else if (e.Category.Equals(BuiltInCategory.OST_PipeCurves))
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference r, XYZ p)
        {
            return true;
        }
    }


    /// <summary>
    /// Allows the selection of duct or pipe, preventing the selection of a previously selected element.
    /// </summary>
    internal class DuctPipeSelectionFilter : ISelectionFilter
    {
        public ElementId PreviousElementID { get; set; }

        public bool AllowElement(Element e)
        {
            if ((e.Category.Name == "Ducts"
                || e.Category.Name == "Pipes")
                && e.Id != PreviousElementID)
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference r, XYZ p)
        {
            return true;
        }
    }


    /// <summary>
    /// Allows the selection of duct, preventing the selection of a previously selected element.
    /// </summary>
    internal class DuctSelectionFilter : ISelectionFilter
    {
        public ElementId PreviousElementID { get; set; }

        public bool AllowElement(Element e)
        {
            if (e.Category.Name == "Ducts"
                && e.Id != PreviousElementID)
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference r, XYZ p)
        {
            return true;
        }
    }


    /// <summary>
    /// Allows the selection of pipe.
    /// </summary>
    internal class PipeSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element e)
        {
            if (e.Category.Name == "Pipes")
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference r, XYZ p)
        {
            return false;
        }
    }


    /// <summary>
    /// Allows the selection of pipe, preventing the selection of a previously selected elements.
    /// </summary>
    internal class PipeNewSelectionFilter : ISelectionFilter
    {
        public ElementId PreviousElementID0 { get; set; }
        public ElementId PreviousElementID1 { get; set; }

        public bool AllowElement(Element e)
        {
            if (e.Category.Name == "Pipes"
                && e.Id != PreviousElementID0
                && e.Id != PreviousElementID1)
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference r, XYZ p)
        {
            return true;
        }
    }


    /// <summary>
    /// Allows the selection of fittings.
    /// </summary>
    internal class FittingSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element e)
        {
            if (e.Category.Name == "Conduit Fittings"
                || e.Category.Name == "Pipe Fittings"
                || e.Category.Name == "Pipe Accessories"
                || e.Category.Name == "Duct Fittings")
            {
                return true;
            }
            else if (e.Category.Equals(BuiltInCategory.OST_PipeCurves))
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference r, XYZ p)
        {
            return true;
        }
    }


    /// <summary>
    /// Allows the selection of sprinkler heads
    /// </summary>
    internal class FabricationSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element e)
        {
#if (!R2016)
            if (e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_FabricationDuctwork
                || e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_FabricationPipework
                || e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeAccessory)
            {
                return true;
            }
#endif 
            return false;
        }

        public bool AllowReference(Reference r, XYZ p)
        {
            return false;
        }
    }


    /// <summary>
    /// Allows only the selection of a previously selected element.
    /// </summary>
    internal class PreviousElementSelectionFilter : ISelectionFilter
    {
        public ElementId PreviousElementID { get; set; }

        public bool AllowElement(Element e)
        {
            if (e.Id == PreviousElementID)
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference r, XYZ p)
        {
            return true;
        }
    }


    /// <summary>
    /// Allows the selection of pipe or sprinkler heads, preventing the selection of a previously selected elements..
    /// </summary>
    internal class PipeorSprinkSelectionFilter : ISelectionFilter
    {
        public ElementId PreviousElementID0 { get; set; }
        public ElementId PreviousElementID1 { get; set; }

        public bool AllowElement(Element e)
        {
            if ((e.Category.Name == "Pipes" || e.Category.Name == "Sprinklers")
                && e.Id != PreviousElementID0
                && e.Id != PreviousElementID1)
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference r, XYZ p)
        {
            return true;
        }
    }
}
