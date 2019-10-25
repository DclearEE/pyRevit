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

namespace NBK.Spool
{
    //This command lets the user select pipe and create a spool draing from it

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    public class CalculateElevation : IExternalCommand
    {

        Autodesk.Revit.ApplicationServices.Application m_application;
        Document m_document;
        UIDocument s_document;

        protected int spoolInt;

        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            try
            {
                m_application = commandData.Application.Application;
                m_document = commandData.Application.ActiveUIDocument.Document;
                s_document = commandData.Application.ActiveUIDocument;


                Transaction transaction = new Transaction(m_document, "Update Invert Elevations");


                transaction.Start();

                Autodesk.Revit.DB.View activeView = m_document.ActiveView;

                FilteredElementCollector fabParts = new FilteredElementCollector(m_document).OfClass(typeof(FabricationPart));

                foreach (FabricationPart fabPart in fabParts)
                {
                    if (fabPart.ProductName != "Pipe") continue;

                    Level level = m_document.GetElement(fabPart.LevelId) as Level;

                    //if (level == null) continue;

                    double localInvert = fabPart.get_Parameter(BuiltInParameter.FABRICATION_BOTTOM_ELEVATION_OF_PART).AsDouble();

                    double gloablInvert = level.Elevation + localInvert;

                    Parameter globalInvertParam = fabPart.LookupParameter("Absolute Invert");
                    globalInvertParam.Set(gloablInvert);

                    //#region log
                    //using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                    //{
                    //    writer.WriteLine("Pipe: {0}", fabPart.Id);
                    //    writer.WriteLine("Level: {0}", level.Name);
                    //    writer.WriteLine("Level: {0}", level.Elevation);
                    //    writer.WriteLine("localInvert: {0}", localInvert);
                    //    writer.WriteLine("gloablInvert: {0}", gloablInvert);
                    //    writer.WriteLine("gloablInvert: {0}", fabPart.LookupParameter("Absolute Invert").AsDouble());
                    //}
                    //#endregion log
                }

                TaskDialog.Show("Complete!", "Absolute Inverts have been updated!");

                transaction.Commit();

                return Autodesk.Revit.UI.Result.Succeeded;
            }
            //If the user right-clicks or presses Esc, handle the exception
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (System.Exception e)
            {
                message += e.ToString();
                return Autodesk.Revit.UI.Result.Failed;
            }
        }
    }
}
