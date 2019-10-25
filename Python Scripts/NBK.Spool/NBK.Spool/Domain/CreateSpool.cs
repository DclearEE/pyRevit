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
    public class CreateSpool : IExternalCommand
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





                
                //***************  Required Family Types  ***************






                // Get a Tag Family Type for Pipe
                ElementId tagIdPipe = NBKHelpers.GetTypeId(m_document, BuiltInCategory.OST_FabricationPipeworkTags, typeof(FamilySymbol), AppKeys.SpoolPipeTagName);
                if (null == tagIdPipe)
                {
                    TaskDialog.Show("Warning", "This command requires an Annotation Symbol Family named " + AppKeys.SpoolPipeTagName + " in order to tag.");
                    return Result.Cancelled;
                }

                // Get a Tag Family Type for Fittings
                ElementId tagIdFitting = NBKHelpers.GetTypeId(m_document, BuiltInCategory.OST_FabricationPipeworkTags, typeof(FamilySymbol), AppKeys.SpoolFittingTagName);
                if (null == tagIdFitting)
                {
                    TaskDialog.Show("Warning", "This command requires an Annotation Symbol Family named " + AppKeys.SpoolFittingTagName + " in order to tag.");
                    return Result.Cancelled;
                }

                // Get a North Arrow Generic Annotation Family Type
                ElementId northArrowId = NBKHelpers.GetTypeId(m_document, BuiltInCategory.OST_GenericAnnotation, typeof(FamilySymbol), AppKeys.SpoolNorthArrow);
                if (null == northArrowId)
                {
                    TaskDialog.Show("Warning", "This command requires an Annotation Symbol Family named " + AppKeys.SpoolNorthArrow + ".");
                    return Result.Cancelled;
                }

                // Get a North Arrow Iso Family Type
                ElementId northArrowIsoId = NBKHelpers.GetTypeId(m_document, BuiltInCategory.OST_GenericAnnotation, typeof(FamilySymbol), AppKeys.SpoolNorthArrowIso);
                if (null == northArrowIsoId)
                {
                    TaskDialog.Show("Warning", "This command requires a Generic Model Family named " + AppKeys.SpoolNorthArrowIso + ".");
                    return Result.Cancelled;
                }

                // Get a View3D Family Type
                ElementId view3DTypeId = NBKHelpers.GetViewTypeId(m_document, AppKeys.SpoolView3DType);
                if (null == view3DTypeId)
                {
                    TaskDialog.Show("Warning", "This command requires an View 3D Family named " + AppKeys.SpoolView3DType + " in order to create a floor plan.");
                    return Result.Cancelled;
                }

                // Find a ViewSection type
                ElementId sectionTypeId = NBKHelpers.GetViewTypeId(m_document, AppKeys.SpoolViewSectionType);

                if (null == sectionTypeId)
                {
                    TaskDialog.Show("Warning", "This command requires an Section Plan Family named " + AppKeys.SpoolViewSectionType + " in order to create a floor plan.");
                    return Result.Cancelled;
                }

                // Find a plan view type
                ElementId planTypeId = NBKHelpers.GetViewTypeId(m_document, AppKeys.SpoolViewPlanType);
                if (null == planTypeId)
                {
                    TaskDialog.Show("Warning", "This command requires an View Plan Family named " + AppKeys.SpoolViewPlanType + " in order to create a floor plan.");
                    return Result.Cancelled;
                }

                // Find a schedule type
                ElementId fittingSchedTypeId = NBKHelpers.GetViewTypeId(m_document, AppKeys.SpoolFittingScheduleType);
                if (null == fittingSchedTypeId)
                {
                    TaskDialog.Show("Warning", "This command requires an Schedule Family Type named " + AppKeys.SpoolFittingScheduleType + " in order to create a fitting schedule.");
                    return Result.Cancelled;
                }

                // Find a schedule type
                ElementId pipeSchedTypeId = NBKHelpers.GetViewTypeId(m_document, AppKeys.SpoolPipeScheduleType);
                if (null == pipeSchedTypeId)
                {
                    TaskDialog.Show("Warning", "This command requires an Schedule Family Type named " + AppKeys.SpoolPipeScheduleType + " in order to create a pipe schedule.");
                    return Result.Cancelled;
                }

                // Find a schedule type
                ElementId weightSchedTypeId = NBKHelpers.GetViewTypeId(m_document, AppKeys.SpoolWeightScheduleType);
                if (null == weightSchedTypeId)
                {
                    TaskDialog.Show("Warning", "This command requires an Schedule Family Type named " + AppKeys.SpoolWeightScheduleType + " in order to create a weight schedule.");
                    return Result.Cancelled;
                }

                // Find a schedule view template
                ElementId fittingSchedTemplateId = NBKHelpers.GetViewTemplate(m_document, AppKeys.SpoolFittingSchedViewTemplate);
                if (null == fittingSchedTemplateId)
                {
                    TaskDialog.Show("Warning", "This command requires an Schedule View Template named " + AppKeys.SpoolFittingSchedViewTemplate + " in order to create a schedule.");
                    return Result.Cancelled;
                }

                // Find a schedule view template
                ElementId pipeSchedTemplateId = NBKHelpers.GetViewTemplate(m_document, AppKeys.SpoolPipeSchedViewTemplate);
                if (null == pipeSchedTemplateId)
                {
                    TaskDialog.Show("Warning", "This command requires an Schedule View Template named " + AppKeys.SpoolPipeSchedViewTemplate + " in order to create a schedule.");
                    return Result.Cancelled;
                }

                // Find a schedule view template
                ElementId weightSchedTemplateId = NBKHelpers.GetViewTemplate(m_document, AppKeys.SpoolWeightSchedViewTemplate);
                if (null == weightSchedTemplateId)
                {
                    TaskDialog.Show("Warning", "This command requires an Schedule View Template named " + AppKeys.SpoolWeightSchedViewTemplate + " in order to create a schedule.");
                    return Result.Cancelled;
                }



                Transaction transaction = new Transaction(m_document, "Create Spool");
                transaction.Start();







                //***************  Select Fabrication Parts  ***************








                ISelectionFilter selFilter = new FabricationSelectionFilter();
                #region log
                using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), false))
                {
                    writer.WriteLine("Starting");
                }
                #endregion log
                ICollection<ElementId> selectedIds = s_document.Selection.GetElementIds();
               
                List<RevitElement> selectedElements = NBKHelpers.GetElementsFromIds(m_document, selectedIds);

                if (selectedIds.Count == 0)
                {
                    IList<Reference> selectedReferences = s_document.Selection.PickObjects(ObjectType.Element, selFilter, "Please select parts to spool");
                    //List<Element> selectedElements = selectedReferences.Select(x => m_document.GetElement(x)).ToList();
                    selectedElements = NBKHelpers.GetElementsFromReferences(m_document, selectedReferences);
                }

                // Sort the selected elements by their name and size
                selectedElements = selectedElements.OrderBy(w => w.PartType).ThenBy(w => w.FamilyName).ThenBy(w => w.Size).ThenBy(w => w.FamilySizeandLength).ToList();

                if(selectedElements.Count == 0)
                {
                    TaskDialog.Show("Cancel", "Selection contains no parts, " +
                        "or all parts in the selection belong to an existing spool.  " +
                        "Please ensure that selected elements belong to the MEP Fabrication Pipework category," +
                        "and that their SPOOL NUMBER parameter is blank.");
                    transaction.Commit();
                    return Result.Failed;
                }
                //
                // Get the Bounding Box of all the selected elements
                BoundingBoxXYZ elementExtents = NBKHelpers.GetBoudingBox(m_document, selectedElements);

                





                //***************  Number Items in Spool  ***************






                // Get all the spool numbers in the model
                List<SpoolNumber> existingSpoolNumbers = NBKHelpers.GetExistingSpoolNumbers(m_document);

                Forms.SpoolForm spoolForm = new Forms.SpoolForm(existingSpoolNumbers);
                DialogResult formResult = spoolForm.ShowDialog();

                if (formResult == DialogResult.Cancel)
                {
                    transaction.Commit();
                    return Result.Failed;
                }


                string spoolNo = "1";
                if (existingSpoolNumbers.Count() > 0)
                {
                    spoolNo = (existingSpoolNumbers.Last().Number + 1).ToString();
                }

                spoolNo = spoolForm.NewSpoolNumber;

                // Update selected elements with a spool number and part type
                NBKHelpers.AddSpoolNumberToElements(m_document, selectedElements, spoolNo);

                // Number the selected elements
                NBKHelpers.NumberSpoolElements(selectedElements);


                // Get the Best orientation for the spool
                List<RevitElement> directions = new List<RevitElement>();
                foreach (RevitElement rE in selectedElements)
                {
                    foreach(XYZ xyz in rE.PartDirections)
                    {
                        string directionName = xyz.X.ToString() + "|" + xyz.Y.ToString();
                        directions.Add(new RevitElement(directionName, xyz));
                    }
                }

                //There may be no directions detected if every selected pipe is vertical
                XYZ spoolDirection = new XYZ(1, 0, 0);

                if (directions.Count != 0)
                {
                    spoolDirection = directions.GroupBy(x => x.Name).OrderByDescending(g => g.Count())
                       .SelectMany(x => x).ToList().First().PartDirection;

                    spoolDirection = -1.0 * spoolDirection;
                }

                #region log
                using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                {
                    writer.WriteLine("spoolDirection: {0}, {1}", Math.Round(spoolDirection.X, 3), Math.Round(spoolDirection.Y, 3));
                }
                #endregion log


                // Tag fabrication parts adjactent to the selected spool
                NBKHelpers.TagAdjacentSpools(m_document, selectedElements, spoolNo);



                //***************  3D View  ***************








                View3D spool3DView = NBKHelpers.Create3DView(m_document, view3DTypeId, elementExtents, spoolNo, northArrowIsoId);











                //***************  Floor Plan View  ***************




                ViewPlan spoolPlan = NBKHelpers.CreatePlan(m_document, planTypeId, elementExtents, spoolNo, null, northArrowId);

                //ViewPlan spoolPlanOrtiented = NBKHelpers.CreatePlan(m_document, planTypeId, elementExtents, spoolNo + "a", spoolDirection, northArrowId);

                if (null == spoolPlan)
                {
                    transaction.Commit();
                    return Result.Cancelled;
                }
                







                //***************  Section Views  ***************






                ViewSection spoolSectionX = NBKHelpers.CreateSection(m_document, sectionTypeId, elementExtents, new XYZ(-1, 0, 0), spoolNo, "Section North");

                ViewSection spoolSectionAligned = NBKHelpers.CreateSection(m_document, sectionTypeId, elementExtents, spoolDirection, spoolNo, "Section Aligned");

                ViewSection spoolSectionPerp = NBKHelpers.CreateSection(m_document, sectionTypeId, elementExtents, spoolDirection.CrossProduct(new XYZ(0,0,1)), spoolNo, "Section Perpendicular");
                
                if (null == spoolSectionX ||
                    null == spoolSectionAligned ||
                    null == spoolSectionPerp)
                {
                    transaction.Commit();
                    return Result.Cancelled;
                }








                //***************  View Filter  ***************





                //Add filters to all the Views
                ParameterFilterElement spoolNumberFalseFilter = NBKHelpers.CreateFilter(m_document, spoolNo, false);
                NBKHelpers.AddFiltertoView(m_document, spool3DView, spoolNumberFalseFilter);
                NBKHelpers.AddFiltertoView(m_document, spoolPlan, spoolNumberFalseFilter);
                NBKHelpers.AddFiltertoView(m_document, spoolSectionX, spoolNumberFalseFilter);
                NBKHelpers.AddFiltertoView(m_document, spoolSectionAligned, spoolNumberFalseFilter);
                NBKHelpers.AddFiltertoView(m_document, spoolSectionPerp, spoolNumberFalseFilter);


                //Create a filter for color-theming assembly plans
                ParameterFilterElement spoolNumberTrueFilter = NBKHelpers.CreateFilter(m_document, spoolNo, true);






                //***************  Schedule View  ***************



                // Pipe Schedule
                ScheduleManager scheduleManagerPipe = 
                    new ScheduleManager(m_application, 
                    m_document, 
                    BuiltInCategory.OST_FabricationPipework, 
                    selectedElements.First().Element,
                    spoolNo,
                    AppKeys.SpoolPipe,
                    pipeSchedTypeId,
                    pipeSchedTemplateId);
                ViewSchedule vSPipe = scheduleManagerPipe.ViewSchedule;


                // Fitting Schedule
                ScheduleManager scheduleManagerFitting =
                    new ScheduleManager(m_application,
                    m_document,
                    BuiltInCategory.OST_FabricationPipework,
                    selectedElements.First().Element,
                    spoolNo,
                    AppKeys.SpoolFitting,
                    fittingSchedTypeId,
                    fittingSchedTemplateId);
                ViewSchedule vSFitting = scheduleManagerFitting.ViewSchedule;


                // Weight Schedule
                ScheduleManager scheduleManagerWeight =
                    new ScheduleManager(m_application,
                    m_document,
                    BuiltInCategory.OST_FabricationPipework,
                    selectedElements.First().Element,
                    spoolNo,
                    null,
                    weightSchedTypeId,
                    weightSchedTemplateId);
                ViewSchedule vSWeight = scheduleManagerWeight.ViewSchedule;



                //*************** Commit Transaction  ***************






                transaction.Commit();
                transaction.Start();







                //***************  OrientPlan  ***************

                double angle = (new XYZ(0, 1, 0)).AngleTo(spoolDirection);
                angle = angle - Math.PI / 2.0;


                //Rotate the floor plan if it is not orthagonal
                if (angle != 0
                    && Math.Abs(Math.Round(angle, 3)) != Math.Round(Math.PI, 3) //180
                    && Math.Abs(Math.Round(angle, 3)) != Math.Round(Math.PI / 2, 3) //90
                    && Math.Abs(Math.Round(angle, 3)) != Math.Round(3* Math.PI / 2, 3)) //270
                {
                    NBKHelpers.RotateBoundingBox(m_document, spoolPlan, angle);
                }


                //***************  Tagging  ***************
                // Tagging has to occur after the Views 
                // are created and the transaction is commited.
                // Some Views can be tagged without commiting
                // the transaction, but 3D Views cannot.


                

                //NBKHelpers.TagAllInView(m_document, tagIdPipe, tagIdFitting, selectedElements, spool3DView, spool3DView, true);
                //NBKHelpers.TagAllInView(m_document, tagIdPipe, tagIdFitting, selectedElements, spoolSectionAligned, spool3DView, false);
                //NBKHelpers.TagAllInView(m_document, tagIdPipe, tagIdFitting, selectedElements, spoolPlan, spool3DView, false);





                //***************  Sheet  ***************







                ElementId titleblockId = NBKHelpers.GetTypeId(m_document, BuiltInCategory.OST_TitleBlocks, typeof(FamilySymbol), AppKeys.SpoolTitleblock);

                if(titleblockId == null)
                {
                    TaskDialog.Show("Warning", "Titleblock " + AppKeys.SpoolTitleblock + "not Loaded");
                }
                ViewSheet newSheet = ViewSheet.Create(m_document, titleblockId);


                //
                // Place all the views on the sheet
                //
                
                XYZ scheduleFittingInsertion = new XYZ(1.0125 / 12.0, 10.0125 / 12.0, 0);
                XYZ schedulePipeInsertion = new XYZ(7.7625 / 12.0, 10.0125 / 12.0, 0);
                XYZ scheduleWeightInsertion = new XYZ(8.9 / 12.0, 0.908 / 12.0, 0);
                XYZ isoCenter = new XYZ(5.0 / 12.0, 4.0 / 12.0, 0);
                XYZ isoNorthArrow = new XYZ(isoCenter.X - 1.0 / 12.0, isoCenter.Y + 1.0 / 12.0, 0);
                XYZ planCenter = new XYZ(13.375 / 12.0, 6.25 / 12.0, 0);
                XYZ elevCenter = new XYZ(13.375 / 12.0, 2.5 / 12.0, 0);
                XYZ elev2Center = new XYZ(19.0 / 12.0, 2.5 / 12.0, 0);
                XYZ elev3Center = new XYZ(19.0 / 12.0, 6.25 / 12.0, 0);
                //XYZ isoInsertion = NBKHelpers.Calc3DViewInsertion(projectedHeightWidth, new XYZ(2/12, 1.5/12, 0), scale3D);

                Viewport viewport3D = Viewport.Create(m_document, newSheet.Id, spool3DView.Id, isoCenter);
                Viewport viewportPlan = Viewport.Create(m_document, newSheet.Id, spoolPlan.Id, planCenter);
                Viewport viewportSection = Viewport.Create(m_document, newSheet.Id, spoolSectionAligned.Id, elevCenter);
                Viewport viewportSection2 = Viewport.Create(m_document, newSheet.Id, spoolSectionPerp.Id, elev2Center);
                Viewport viewportSection3 = Viewport.Create(m_document, newSheet.Id, spoolSectionX.Id, elev3Center);
                ScheduleSheetInstance.Create(m_document, newSheet.Id, vSFitting.Id, scheduleFittingInsertion);
                ScheduleSheetInstance.Create(m_document, newSheet.Id, vSPipe.Id, schedulePipeInsertion);
                ScheduleSheetInstance.Create(m_document, newSheet.Id, vSWeight.Id, scheduleWeightInsertion);

                // Place the north arrow for the isometric view
                m_document.Create.NewFamilyInstance(isoNorthArrow, m_document.GetElement(northArrowIsoId) as FamilySymbol, newSheet);

                // Get all the viewport Types in the project
                List<Element> viewportTypes = NBKHelpers.GetViewportTypes(m_document);


                // Set the Viewtype of the 3D View
                NBKHelpers.SetViewportType(m_document, viewport3D, viewportTypes, AppKeys.SpoolViewport3DType);
                NBKHelpers.SetViewportType(m_document, viewportPlan, viewportTypes, AppKeys.SpoolViewportPlanType);
                NBKHelpers.SetViewportType(m_document, viewportSection, viewportTypes, AppKeys.SpoolViewportSectionType);
                NBKHelpers.SetViewportType(m_document, viewportSection2, viewportTypes, AppKeys.SpoolViewportSectionType);
                NBKHelpers.SetViewportType(m_document, viewportSection3, viewportTypes, AppKeys.SpoolViewportSectionType);

                //
                // Set titleblock information
                //
                newSheet.get_Parameter(BuiltInParameter.SHEET_NAME).Set(spoolNo);
                newSheet.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set("SP-" + spoolNo);

                //LinksManager linksManager = new LinksManager(m_application, m_document);
                //LevelManager levelManager = new LevelManager(m_application, m_document);
                //SleevesForm form = new SleevesForm(linksManager, levelManager);
                //form.ShowDialog();
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
