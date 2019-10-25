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
    public class RotateIso : IExternalCommand
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


                Transaction transaction = new Transaction(m_document, "Rotate Iso");


                transaction.Start();

                Autodesk.Revit.DB.View activeView = m_document.ActiveView;

                FilteredElementCollector viewports = new FilteredElementCollector(m_document, activeView.Id).OfClass(typeof(Viewport));
                View3D isoView = null;
                Viewport viewport = null;
                foreach (Viewport vP in viewports)
                {
                    isoView = m_document.GetElement(vP.ViewId) as View3D;

                    if(isoView != null)
                    {
                        viewport = vP;
                        break;
                    }
                }

                if(isoView == null) return Result.Cancelled;

                XYZ insertion = viewport.GetBoxCenter();

                // Unpin the view
                isoView.Unlock();


                ViewOrientation3D orientation = isoView.GetOrientation();
                XYZ eyePos = orientation.EyePosition;
                XYZ forwardDir = orientation.ForwardDirection;
                XYZ up = orientation.UpDirection;

                // Get the center
                XYZ center = (isoView.GetSectionBox().Max + isoView.GetSectionBox().Min) / 2.0;

                // Rotate the eye position
                double transX = (eyePos.X - center.X);
                double transY = (eyePos.Y - center.Y);
                double newEyeX = transX * Math.Cos(Math.PI / 2.0) - transY * Math.Sin(Math.PI / 2.0) + center.X;
                double newEyeY = transY * Math.Cos(Math.PI / 2.0) + transX * Math.Sin(Math.PI / 2.0) + center.Y;

                // Rotate the forward Direction
                double newForX = forwardDir.X * Math.Cos(Math.PI / 2.0) - forwardDir.Y * Math.Sin(Math.PI / 2.0);
                double newForY = forwardDir.Y * Math.Cos(Math.PI / 2.0) + forwardDir.X * Math.Sin(Math.PI / 2.0);

                // Rotate the forward Direction
                double newUpX = up.X * Math.Cos(Math.PI / 2.0) - up.Y * Math.Sin(Math.PI / 2.0);
                double newUpY = up.Y * Math.Cos(Math.PI / 2.0) + up.X * Math.Sin(Math.PI / 2.0);

                ViewOrientation3D newOrientation = new ViewOrientation3D(new XYZ(newEyeX, newEyeY, eyePos.Z), new XYZ(newUpX, newUpY, up.Z), new XYZ(newForX, newForY, forwardDir.Z));
                isoView.SetOrientation(newOrientation);

                ElementTransformUtils.MoveElement(m_document, viewport.Id, insertion - viewport.GetBoxCenter());



                // Lock the view
                isoView.SaveOrientationAndLock();

                // Rename the View
                Parameter nameParam = isoView.get_Parameter(BuiltInParameter.VIEW_NAME);
                string name = nameParam.AsString();
                if (name.Contains("SouthWest"))
                    nameParam.Set(name.Replace("SouthWest", "SouthEast"));
                else if (name.Contains("SouthEast"))
                    nameParam.Set(name.Replace("SouthEast", "NorthEast"));
                else if (name.Contains("NorthEast"))
                    nameParam.Set(name.Replace("NorthEast", "NorthWest"));
                else if (name.Contains("NorthWest"))
                    nameParam.Set(name.Replace("NorthWest", "SouthWest"));

                // Find the North Arrow and 'rotate' it
                FilteredElementCollector symbols = new FilteredElementCollector(m_document, activeView.Id)
                    .OfCategory(BuiltInCategory.OST_GenericAnnotation);

                foreach(FamilyInstance fI in symbols)
                {
                    if(fI.Symbol.Name.ToUpper().Contains("NORTH ARROW"))
                    {
                        try
                        {
                            if(fI.LookupParameter("SouthWest").AsInteger() == 1)
                            {
                                fI.LookupParameter("SouthWest").Set(0);
                                fI.LookupParameter("SouthEast").Set(1);
                                fI.LookupParameter("NorthEast").Set(0);
                                fI.LookupParameter("NorthWest").Set(0);
                            }
                            else if (fI.LookupParameter("SouthEast").AsInteger() == 1)
                            {
                                fI.LookupParameter("SouthWest").Set(0);
                                fI.LookupParameter("SouthEast").Set(0);
                                fI.LookupParameter("NorthEast").Set(1);
                                fI.LookupParameter("NorthWest").Set(0);
                            }
                            else if (fI.LookupParameter("NorthEast").AsInteger() == 1)
                            {
                                fI.LookupParameter("SouthWest").Set(0);
                                fI.LookupParameter("SouthEast").Set(0);
                                fI.LookupParameter("NorthEast").Set(0);
                                fI.LookupParameter("NorthWest").Set(1);
                            }
                            else if (fI.LookupParameter("NorthWest").AsInteger() == 1)
                            {
                                fI.LookupParameter("SouthWest").Set(1);
                                fI.LookupParameter("SouthEast").Set(0);
                                fI.LookupParameter("NorthEast").Set(0);
                                fI.LookupParameter("NorthWest").Set(0);
                            }
                        }
                        catch { }
                    }
                }

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
