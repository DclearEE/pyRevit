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

using Excel = Microsoft.Office.Interop.Excel;

namespace NBK.Spool
{
    //This command lets the user select pipe and create a spool draing from it

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    public class ExportBOM : IExternalCommand
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

                #region log
                using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), false))
                {
                    writer.WriteLine("Starting");
                }
                #endregion log

                //Create Spreadsheets
                object misValue = System.Reflection.Missing.Value;

                Excel.Application xlApp = new Excel.Application();
                xlApp.DisplayAlerts = false;


                //Select an output folder
                string outputFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\BOM2\\";


                Autodesk.Revit.DB.View activeView = m_document.ActiveView;

                FabricationConfiguration fabconfig = FabricationConfiguration.GetFabricationConfiguration(m_document);
                

                FilteredElementCollector schedules = new FilteredElementCollector(m_document).OfClass(typeof(ViewSchedule));
                FilteredElementCollector fabParts = new FilteredElementCollector(m_document).OfCategory(BuiltInCategory.OST_FabricationPipework).OfClass(typeof(FabricationPart));

                List<RevitElement> revitElements = new List<RevitElement>();

                foreach (FabricationPart fabPart in fabParts)
                {
                    string spoolNo = fabPart.LookupParameter(AppKeys.SpoolParameterName).AsString();

                    if (spoolNo == "" || spoolNo == null)
                        continue;

                    string alias = fabPart.Alias;


                    //fabconfig.get

                    #region log
                    using (StreamWriter writer = new StreamWriter(MicrodeskHelpers.Log(), true))
                    {
                        writer.WriteLine(fabPart.Name);
                        writer.WriteLine(fabPart.Alias);
                        writer.WriteLine(fabPart.GetTypeId());
                    }
                    #endregion log

                    FabricationPartType fabType = m_document.GetElement(fabPart.GetTypeId()) as FabricationPartType;
                    string familyTypeName = fabType.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString();
                    string familyName = fabType.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString();
                    string manufacturer = fabPart.get_Parameter(BuiltInParameter.FABRICATION_PRODUCT_DATA_OEM).AsString();
                    string longDescription = fabPart.ProductLongDescription;
                    string shortDescription = fabPart.ProductShortDescription;
                    string product = fabPart.ProductName;
                    string material = fabPart.get_Parameter(BuiltInParameter.FABRICATION_PART_MATERIAL).AsValueString().Replace("BD Piping Material: ", "");
                    


                    //Classify the part as a pipe or fitting
                    string partType = "Pipe Fitting";
                    if (fabPart.IsAStraight()) partType = "Pipe";

                    double length = 0;
                    string pipeLength = "";
                    Parameter lengthParam = fabPart.get_Parameter(BuiltInParameter.FABRICATION_PART_LENGTH);
                    if (lengthParam != null)
                    {
                        length = lengthParam.AsDouble() * 12.0;
                        pipeLength = lengthParam.AsValueString();
                    }

                    string size = fabPart.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString();

                    RevitElement revitElement = new RevitElement(fabPart, 
                        familyName, 
                        familyTypeName, 
                        partType, 
                        size,
                        length,
                        pipeLength, 
                        spoolNo,
                        manufacturer,
                        "product line",
                        shortDescription,
                        product,
                        material);

                    revitElements.Add(revitElement);
                }

                // Sort the list by spool No
                revitElements = revitElements.OrderBy(w => w.SpoolNumber).ThenBy(w => w.PartType).ThenBy(w => w.Size).ThenBy(w => w.FamilySizeandLength).ToList();

                IOrderedEnumerable<IGrouping<string, RevitElement>> spoolEnumerable =
                    from rE in revitElements
                    group rE by rE.SpoolNumber into newGroup
                    orderby newGroup.Key
                    select newGroup;

                foreach (IGrouping<string, RevitElement> nameGroup in spoolEnumerable)
                {
                    if (nameGroup.Count() == 0) continue;


                    string spoolNo = nameGroup.Key;
                    string pathPipeBOM = outputFolder + spoolNo + " - BOM - Pipes.xls";
                    string pathFittingBOM = outputFolder + spoolNo + " - BOM - Fittings.xls";

                    Excel.Workbook xlWorkBookPipe = xlApp.Workbooks.Add(misValue);
                    Excel.Worksheet xlWorkSheetPipe = (Excel.Worksheet)xlWorkBookPipe.Worksheets.get_Item(1);

                    xlWorkBookPipe.SaveAs(pathPipeBOM, 
                        Excel.XlFileFormat.xlWorkbookNormal, 
                        misValue, 
                        misValue, 
                        misValue, 
                        misValue, 
                        Excel.XlSaveAsAccessMode.xlExclusive, 
                        misValue, 
                        misValue, 
                        misValue, 
                        misValue, 
                        misValue);

                    Excel.Workbook xlWorkBookFitting = xlApp.Workbooks.Add(misValue);
                    Excel.Worksheet xlWorkSheetFitting = (Excel.Worksheet)xlWorkBookFitting.Worksheets.get_Item(1);

                    xlWorkBookFitting.SaveAs(pathFittingBOM,
                        Excel.XlFileFormat.xlWorkbookNormal,
                        misValue,
                        misValue,
                        misValue,
                        misValue,
                        Excel.XlSaveAsAccessMode.xlExclusive,
                        misValue,
                        misValue,
                        misValue,
                        misValue,
                        misValue);

                    int pipeRow = 1;
                    int fittingRow = 1;
                    pipeRow = NBKHelpers.WriteColumnHeadersPipe(xlWorkSheetPipe, pipeRow);
                    fittingRow = NBKHelpers.WriteColumnHeadersFittings(xlWorkSheetFitting, fittingRow);
                    foreach (RevitElement rE in nameGroup)
                    {
                        if(rE.PartType == "Pipe")
                        {
                            
                            //xlWorkSheetPipe.Cells[pipeRow, 1] = rE.FamilyTypeName;
                            xlWorkSheetPipe.Cells[pipeRow, 1] = rE.FamilyName;
                            xlWorkSheetPipe.Cells[pipeRow, 2] = rE.Size;
                            xlWorkSheetPipe.Cells[pipeRow, 3] = "1";
                            xlWorkSheetPipe.Cells[pipeRow, 4] = rE.Length;
                            xlWorkSheetPipe.Cells[pipeRow, 5] = rE.PipeLength;
                            xlWorkSheetPipe.Cells[pipeRow, 6] = rE.Manufacturer;
                            xlWorkSheetPipe.Cells[pipeRow, 7] = rE.ProductLine;
                            xlWorkSheetPipe.Cells[pipeRow, 8] = rE.ShortDescription;
                            xlWorkSheetPipe.Cells[pipeRow, 9] = rE.Product;
                            xlWorkSheetPipe.Cells[pipeRow, 10] = rE.Material;
                            pipeRow++;
                        }
                        else if(rE.PartType == "Pipe Fitting")
                        {
                            //xlWorkSheetFitting.Cells[fittingRow, 1] = rE.FamilyTypeName;
                            xlWorkSheetFitting.Cells[fittingRow, 1] = rE.FamilyName; //Long Description
                            xlWorkSheetFitting.Cells[fittingRow, 2] = rE.Size; //Size
                            xlWorkSheetFitting.Cells[fittingRow, 3] = "1"; //Count
                            xlWorkSheetFitting.Cells[fittingRow, 4] = rE.Manufacturer; //Manufacturer Name
                            xlWorkSheetFitting.Cells[fittingRow, 5] = rE.ProductLine;
                            xlWorkSheetFitting.Cells[fittingRow, 6] = rE.ShortDescription;
                            xlWorkSheetFitting.Cells[fittingRow, 7] = rE.Product;
                            xlWorkSheetFitting.Cells[fittingRow, 8] = rE.Material; //Material
                            fittingRow++;
                        }
                    }

                    //Add formulas
                    xlWorkSheetPipe.Cells[pipeRow, 4] = "=sum(D2:D" + (pipeRow - 1).ToString();
                    xlWorkSheetPipe.Cells[pipeRow, 5] = "=D" + (pipeRow - 1).ToString() + "/12";

                    xlWorkSheetFitting.Cells[fittingRow, 3] = "=sum(C2:C" + (fittingRow - 1).ToString();

                    //Add Grid Lines
                    NBKHelpers.AddBorders(xlWorkSheetPipe);
                    NBKHelpers.AddBorders(xlWorkSheetFitting);

                    //
                    //Sort the rows
                    //Excel.Range usedCells = xlWorkSheetPipe.UsedRange;
                    //Excel.Range rangeAllData = xlWorkSheetPipe.Range[xlWorkSheetPipe.Cells[6, 1], xlWorkSheetPipe.Cells[usedCells.Rows.Count, usedCells.Columns.Count]];
                    //object missing = System.Type.Missing;
                    //rangeAllData.Sort(
                    //                    rangeAllData.Columns[1], Excel.XlSortOrder.xlAscending,
                    //                    rangeAllData.Columns[2], missing, Excel.XlSortOrder.xlAscending,
                    //                    rangeAllData.Columns[3], Excel.XlSortOrder.xlAscending,
                    //                    Excel.XlYesNoGuess.xlNo, missing, missing,
                    //                    Excel.XlSortOrientation.xlSortColumns,
                    //                    Excel.XlSortMethod.xlPinYin,
                    //                    Excel.XlSortDataOption.xlSortNormal,
                    //                    Excel.XlSortDataOption.xlSortNormal,
                    //                    Excel.XlSortDataOption.xlSortNormal
                    //                    );

                    xlWorkBookPipe.Save();
                    xlWorkBookPipe.Close(true, misValue, misValue);
                    xlWorkBookFitting.Save();
                    xlWorkBookFitting.Close(true, misValue, misValue);
                    //List<RevitElement> grouped = nameGroup.Select(w => w.PartType == "Pipe").ToList();

                }


                TaskDialog.Show("Export BOM's", "Export Complete!");




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
